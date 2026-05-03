"""Loader base + shared helpers (column discovery, batch insert)."""
from __future__ import annotations

import abc

import structlog

from ..core.connections import get_olap_sql
from ..core.exceptions import LoadError

log = structlog.get_logger().bind(component="loader")

# Columns that should never be written by ETL — DB fills them.
SKIP_COLUMNS_DEFAULT = ("Loaded_At",)


def get_table_columns(
    table: str,
    *,
    skip_identity: bool = True,
    skip: tuple[str, ...] = SKIP_COLUMNS_DEFAULT,
) -> list[str]:
    """Return INSERT-eligible column names for a table in OlapBASERP."""
    with get_olap_sql() as c:
        cur = c.cursor()
        cur.execute(
            """SELECT c.name FROM sys.tables t JOIN sys.columns c ON t.object_id = c.object_id
               WHERE t.name = ? AND (? = 0 OR c.is_identity = 0)
               ORDER BY c.column_id""",
            (table, 1 if skip_identity else 0),
        )
        cols = [r[0] for r in cur.fetchall()]
    return [col for col in cols if col not in skip]


def bulk_insert(
    table: str,
    rows: list[dict],
    *,
    columns: list[str] | None = None,
    batch_size: int = 1000,
) -> int:
    """Insert dicts into table; missing keys -> NULL.

    Returns the number of rows inserted. Uses pyodbc fast_executemany.
    """
    if not rows:
        return 0
    cols = columns or get_table_columns(table)
    placeholders = ", ".join(["?"] * len(cols))
    col_list = ", ".join(f"[{c}]" for c in cols)
    sql = f"INSERT INTO [{table}] ({col_list}) VALUES ({placeholders})"
    inserted = 0
    try:
        with get_olap_sql() as c:
            cur = c.cursor()
            cur.fast_executemany = True
            for i in range(0, len(rows), batch_size):
                chunk = rows[i : i + batch_size]
                params = [[r.get(col) for col in cols] for r in chunk]
                cur.executemany(sql, params)
                inserted += len(chunk)
            c.commit()
        log.info("bulk insert", table=table, rows=inserted, cols=len(cols))
        return inserted
    except Exception as exc:
        raise LoadError(f"bulk_insert into {table} failed: {exc!r}") from exc


class Loader(abc.ABC):
    def __init__(self, config: dict) -> None:
        self.config = config

    @abc.abstractmethod
    def load(self, rows: list[dict]) -> int:
        """Return the number of rows persisted."""
