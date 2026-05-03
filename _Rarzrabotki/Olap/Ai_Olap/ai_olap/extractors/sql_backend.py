"""SqlBackendExtractor — read from 1С MSSQL backend directly via pyodbc.

Config shape (extracted from a pipeline step):
    {
        "type": "sql_backend",
        "object": "Справочник.Организации",        # 1С metadata full name
        "fields": ["Ссылка", "Код", "Наименование"],  # optional whitelist; default = all mapped
        "extra_columns": ["_Marked"],                 # optional: raw SQL columns the mapping omits
        "where": "_Marked = 0x00",                    # optional raw WHERE fragment
        "where_params": [],                           # parameters for the fragment
        "limit": 100,                                 # optional TOP N
    }

Or to bypass mapping entirely:
    {
        "type": "raw_sql",
        "sql": "SELECT _IDRRef, _Description FROM _Reference329",
        "params": [],
    }
"""
from __future__ import annotations

import structlog

from ..core.connections import get_baserp_sql
from ..core.decorators import measure_time, retry
from ..core.exceptions import ExtractionError
from ..utils.mapping_resolver import resolve
from .base import Extractor

log = structlog.get_logger().bind(component="sql_backend_extractor")


class SqlBackendExtractor(Extractor):
    @measure_time("sql_backend_extract")
    @retry(max_attempts=3, backoff=2.0, exceptions=(ExtractionError,))
    def extract(self) -> list[dict]:
        cfg = self.config
        kind = cfg.get("type", "sql_backend")

        if kind == "raw_sql":
            sql = cfg.get("raw_sql") or cfg.get("sql")
            params = cfg.get("params") or []
            if not sql:
                raise ExtractionError("raw_sql extractor missing 'raw_sql'")
            return self._run_query(sql, params)

        meta_full = cfg.get("object")
        if not meta_full:
            raise ExtractionError("sql_backend extractor missing 'object'")
        table, field_map = resolve(meta_full)

        wanted_1c = cfg.get("fields")
        select_cols: list[str] = []
        if wanted_1c:
            for fname in wanted_1c:
                col = field_map.get(fname)
                if not col:
                    raise ExtractionError(
                        f"{meta_full}: field '{fname}' not in mapping. Known: {list(field_map)[:8]}…"
                    )
                select_cols.append(f'{col} AS "{fname}"')
        else:
            select_cols = [f'{col} AS "{fname}"' for fname, col in field_map.items()]

        for raw in cfg.get("extra_columns", []) or []:
            select_cols.append(raw)

        sql = "SELECT "
        if cfg.get("limit"):
            sql += f"TOP {int(cfg['limit'])} "
        sql += ", ".join(select_cols)
        sql += f" FROM {table}"
        if cfg.get("where"):
            sql += " WHERE " + cfg["where"]
        if cfg.get("order_by"):
            sql += " ORDER BY " + cfg["order_by"]

        return self._run_query(sql, cfg.get("where_params") or [])

    @staticmethod
    def _run_query(sql: str, params: list) -> list[dict]:
        try:
            with get_baserp_sql() as c:
                cur = c.cursor()
                cur.execute(sql, params)
                cols = [d[0] for d in cur.description]
                rows = [dict(zip(cols, r)) for r in cur.fetchall()]
            log.info("sql_backend rows", rows=len(rows), sql_preview=sql[:120])
            return rows
        except Exception as exc:
            raise ExtractionError(f"SQL failed: {exc!r}; query={sql[:200]}") from exc
