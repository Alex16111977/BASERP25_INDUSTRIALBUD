"""ERP query wrapper via Python COM (V83.COMConnector).

Same interface as _mcp_query.execute_query but uses COM directly.
Handles 1C-specific type conversions (dates, None, refs → dict fields).
"""
import datetime
from pathlib import Path
import sys

# Скрипт живе в knowledge_PL/Python/PL/. Конфіг (з CONN_ERP) — у _Rarzrabotki/Python/PnL/config.py.
_PNL_DIR = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL")
sys.path.insert(0, str(_PNL_DIR))
import config  # noqa: E402


_CONN = None


def _connect():
    global _CONN
    if _CONN is not None:
        return _CONN
    import win32com.client
    v8 = win32com.client.Dispatch("V83.COMConnector")
    _CONN = v8.Connect(config.CONN_ERP)
    return _CONN


def _convert_param(v):
    """Convert Python value to COM-safe form.

    Dates: str 'YYYY-MM-DDTHH:MM:SS' → datetime.datetime
    Others: pass through.
    """
    if isinstance(v, str) and len(v) == 19 and v[10] == "T":
        return datetime.datetime.strptime(v, "%Y-%m-%dT%H:%M:%S")
    return v


def _extract_columns(result_table):
    """Get column names from 1C ТаблицаЗначений."""
    cols_coll = result_table.Колонки
    out = []
    for i in range(cols_coll.Количество()):
        out.append(cols_coll.Получить(i).Имя)
    return out


def _row_to_dict(row, columns):
    rec = {}
    for c in columns:
        val = getattr(row, c)
        if val is None:
            rec[c] = None
            continue
        # datetime passthrough
        if isinstance(val, (datetime.datetime, datetime.date)):
            rec[c] = val.isoformat()
            continue
        # bool/int/float passthrough
        if isinstance(val, (bool, int, float)):
            rec[c] = val
            continue
        # COM ref with .Наименование — take string representation
        if hasattr(val, "Наименование"):
            try:
                naim = val.Наименование
                if naim:
                    rec[c] = str(naim)
                    continue
            except Exception:
                pass
        # Fallback
        rec[c] = str(val) if val != "" else ""
    return rec


def execute_query(query_text, parameters=None, max_rows=None):
    """Run query via COM. Returns list[dict].

    parameters: {name: value} with dates as 'YYYY-MM-DDTHH:MM:SS' strings.
    max_rows: optional truncation.
    """
    conn = _connect()
    q = conn.NewObject("Запрос")
    q.Текст = query_text
    for name, val in (parameters or {}).items():
        q.УстановитьПараметр(name, _convert_param(val))
    tz = q.Выполнить().Выгрузить()
    cols = _extract_columns(tz)
    n = tz.Количество()
    if max_rows:
        n = min(n, max_rows)
    rows = []
    for i in range(n):
        rows.append(_row_to_dict(tz.Получить(i), cols))
    return rows
