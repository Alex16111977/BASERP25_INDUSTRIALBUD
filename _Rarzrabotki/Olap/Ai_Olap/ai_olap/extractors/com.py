"""ComExtractor — fallback for 1С virtual tables (Остатки, Обороты, ОстаткиИОбороты).

Direct SQL backend cannot reproduce the platform's virtual-table aggregations.
This extractor calls Запрос via V83.COMConnector and returns rows.

Config shape:
    {
        "type": "com",
        "com_query": "ВЫБРАТЬ ... ИЗ РегистрНакопления.X.ОстаткиИОбороты(...)",
        "params": {"НачалоПериода": "2026-02-01", "КонецПериода": "2026-02-28"}
    }

Date params (str) are converted to BSL dates inside the COM session.
"""
from __future__ import annotations

import datetime as dt
from typing import Any

import structlog

from ..core.connections import get_com_connection
from ..core.decorators import measure_time, retry
from ..core.exceptions import ExtractionError
from .base import Extractor

log = structlog.get_logger().bind(component="com_extractor")


def _to_bsl_date(erp, value: Any):
    """Convert a Python date/string to a BSL Date inside the connected ERP session."""
    if isinstance(value, dt.datetime):
        return erp.NewObject("Дата", value.year, value.month, value.day, value.hour, value.minute, value.second)
    if isinstance(value, dt.date):
        return erp.NewObject("Дата", value.year, value.month, value.day, 0, 0, 0)
    if isinstance(value, str):
        d = dt.date.fromisoformat(value[:10])
        return erp.NewObject("Дата", d.year, d.month, d.day, 0, 0, 0)
    return value


class ComExtractor(Extractor):
    @measure_time("com_extract")
    @retry(max_attempts=2, backoff=3.0, exceptions=(ExtractionError,))
    def extract(self) -> list[dict]:
        cfg = self.config
        text = cfg.get("com_query")
        if not text:
            raise ExtractionError("com extractor missing 'com_query'")

        try:
            erp = get_com_connection()
            q = erp.NewObject("Запрос")
            q.Текст = text
            for k, v in (cfg.get("params") or {}).items():
                if isinstance(v, (dt.date, str)) and ("Период" in k or "Дата" in k):
                    v = _to_bsl_date(erp, v)
                q.УстановитьПараметр(k, v)
            tv = q.Выполнить().Выгрузить()
        except Exception as exc:
            raise ExtractionError(f"COM query failed: {exc!r}") from exc

        cols = [str(c.Имя) for c in tv.Колонки]
        rows: list[dict] = []
        for i in range(tv.Количество()):
            row = tv.Получить(i)
            rows.append({c: row[c] for c in cols})
        log.info("com rows", rows=len(rows), cols=len(cols))
        return rows
