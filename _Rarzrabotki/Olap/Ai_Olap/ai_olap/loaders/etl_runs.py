"""ETL_Runs logging — open/close lifecycle for each pipeline run."""
from __future__ import annotations

import datetime as dt

import structlog

from ..core.connections import get_olap_sql
from ..core.exceptions import LoadError

log = structlog.get_logger().bind(component="etl_runs")


def open_run(script: str, period_month: dt.date | None = None) -> int:
    """Insert a 'Running' row into ETL_Runs and return its Run_ID."""
    try:
        with get_olap_sql() as c:
            cur = c.cursor()
            cur.execute(
                """INSERT INTO ETL_Runs (Script, Period_Month, Started_At, Status)
                   OUTPUT INSERTED.Run_ID
                   VALUES (?, ?, SYSDATETIME(), 'Running')""",
                (script, period_month),
            )
            run_id = cur.fetchval()
            c.commit()
    except Exception as exc:
        raise LoadError(f"ETL_Runs open failed: {exc!r}") from exc
    log.info("etl run opened", run_id=run_id, script=script, period=str(period_month))
    return int(run_id)


def close_run(run_id: int, *, rows_loaded: int | None, status: str, error: str | None = None) -> None:
    try:
        with get_olap_sql() as c:
            c.cursor().execute(
                """UPDATE ETL_Runs
                   SET Finished_At = SYSDATETIME(),
                       Rows_Loaded = ?,
                       Status      = ?,
                       Error       = ?
                   WHERE Run_ID = ?""",
                (rows_loaded, status, error[:1900] if error else None, run_id),
            )
            c.commit()
    except Exception as exc:
        raise LoadError(f"ETL_Runs close failed: {exc!r}") from exc
    log.info("etl run closed", run_id=run_id, status=status, rows=rows_loaded)
