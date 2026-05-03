"""FactLoader — idempotent per period (DELETE WHERE Period_Month=? + INSERT)."""
from __future__ import annotations

import datetime as dt

import structlog

from ..core.connections import get_olap_sql
from ..core.decorators import measure_time
from ..core.exceptions import LoadError
from .base import Loader, bulk_insert

log = structlog.get_logger().bind(component="fact_loader")


class FactLoader(Loader):
    @measure_time("fact_load")
    def load(self, rows: list[dict]) -> int:
        table = self.config["target_table"]
        period_col = self.config.get("period_column", "Period_Month")
        period_value = self.config.get("period")  # set by orchestrator
        if period_value is None:
            raise LoadError(f"FactLoader for {table} requires 'period' (YYYY-MM-DD)")

        # Normalize period to a 1st-of-month date.
        if isinstance(period_value, str):
            period_value = dt.date.fromisoformat(period_value[:10])
        if isinstance(period_value, dt.datetime):
            period_value = period_value.date()
        if isinstance(period_value, dt.date) and period_value.day != 1:
            period_value = dt.date(period_value.year, period_value.month, 1)

        try:
            with get_olap_sql() as c:
                cur = c.cursor()
                cur.execute(f"DELETE FROM [{table}] WHERE [{period_col}] = ?", period_value)
                deleted = cur.rowcount
                c.commit()
        except Exception as exc:
            raise LoadError(f"DELETE from {table} failed: {exc!r}") from exc

        # Force Period_Month onto each row in case extractor missed it.
        for r in rows:
            r.setdefault(period_col, period_value)

        n = bulk_insert(table, rows)
        log.info("fact period reload", table=table, period=str(period_value), deleted=deleted, inserted=n)
        return n
