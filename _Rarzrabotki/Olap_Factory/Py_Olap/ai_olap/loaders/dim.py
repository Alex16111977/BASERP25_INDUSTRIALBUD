"""DimLoader — full reload (TRUNCATE + INSERT) for Dim_* tables."""
from __future__ import annotations

import structlog

from ..core.connections import get_olap_sql
from ..core.decorators import measure_time
from ..core.exceptions import LoadError
from .base import Loader, bulk_insert

log = structlog.get_logger().bind(component="dim_loader")


class DimLoader(Loader):
    @measure_time("dim_load")
    def load(self, rows: list[dict]) -> int:
        table = self.config["target_table"]
        try:
            with get_olap_sql() as c:
                c.cursor().execute(f"TRUNCATE TABLE [{table}]")
                c.commit()
        except Exception as exc:
            raise LoadError(f"TRUNCATE {table} failed: {exc!r}") from exc
        n = bulk_insert(table, rows)
        log.info("dim full reload", table=table, rows=n)
        return n
