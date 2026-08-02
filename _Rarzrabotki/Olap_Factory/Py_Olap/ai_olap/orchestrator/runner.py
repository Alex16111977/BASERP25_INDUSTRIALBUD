"""PipelineRunner — load JSON, validate, wrap with ETL_Runs lifecycle, dispatch."""
from __future__ import annotations

import datetime as dt
import traceback
from pathlib import Path

import structlog

from ..config.loader import load_pipeline_config
from ..config.validator import validate_pipeline
from ..core.exceptions import ETLException
from ..loaders.etl_runs import close_run, open_run
from .pipeline import Pipeline

log = structlog.get_logger().bind(component="runner")


def run_pipeline_file(path: str | Path, *, period: dt.date | None = None) -> int:
    p = Path(path)
    cfg = load_pipeline_config(p)
    validate_pipeline(cfg, source=p.name)
    return run_pipeline(cfg, period=period, script=p.stem)


def run_pipeline(cfg: dict, *, period: dt.date | None = None, script: str | None = None) -> int:
    """Execute a single pipeline. Returns total rows loaded across steps."""
    pid = cfg["pipeline_id"]
    script = script or pid
    run_id = open_run(script, period_month=period)
    total = 0
    try:
        results = Pipeline(cfg, period=period).run()
        total = sum(r.rows_loaded for r in results)
        close_run(run_id, rows_loaded=total, status="Success")
        log.info("pipeline success", pipeline=pid, total=total, run_id=run_id)
    except Exception as exc:
        tb = traceback.format_exc()
        close_run(run_id, rows_loaded=total, status="Failed", error=tb)
        log.error("pipeline failed", pipeline=pid, run_id=run_id, error=repr(exc))
        raise
    return total
