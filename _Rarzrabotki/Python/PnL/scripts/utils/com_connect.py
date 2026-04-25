"""V83.COMConnector wrapper for ERP."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import config  # noqa: E402


def connect_erp():
    import win32com.client
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(config.CONN_ERP)


def run_query(conn, text, params=None):
    query = conn.NewObject("Запрос")
    query.Текст = text
    for k, v in (params or {}).items():
        query.УстановитьПараметр(k, v)
    return query.Выполнить().Выгрузить()


def table_to_records(tz, columns):
    rows = []
    n = tz.Количество()
    for i in range(n):
        r = tz.Получить(i)
        rec = {}
        for col in columns:
            val = getattr(r, col)
            rec[col] = val
        rows.append(rec)
    return rows


def uuid_str(conn, ref):
    """Returns string UUID of a ref via conn.XMLСтрока (which accepts UUID)."""
    uid = ref.УникальныйИдентификатор()
    return str(conn.XMLСтрока(uid))


def is_manually_locked(obj) -> bool:
    """True, если элемент Справочник.А_Статьи_PL помечен финансистом как «Ручная корректировка»
    (реквизит А_РучнаяКорректировка = Истина). Такие элементы пайплайн не трогает."""
    try:
        return bool(obj.А_РучнаяКорректировка)
    except Exception:
        return False
