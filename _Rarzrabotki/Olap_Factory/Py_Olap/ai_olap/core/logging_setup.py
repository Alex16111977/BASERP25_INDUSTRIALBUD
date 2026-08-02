"""structlog setup writing JSONL logs into ./logs/etl_YYYY-MM-DD.log."""
from __future__ import annotations

import datetime as dt
import logging
import os
from pathlib import Path

import structlog

_CONFIGURED = False


def setup_logging(log_dir: str | Path = "logs", level: str | None = None) -> structlog.BoundLogger:
    """Configure structlog once per process. Returns a logger bound with `app=ai_olap`."""
    global _CONFIGURED

    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"etl_{dt.date.today():%Y-%m-%d}.log"

    if not _CONFIGURED:
        level_name = (level or os.environ.get("LOG_LEVEL") or "INFO").upper()
        level_value = getattr(logging, level_name, logging.INFO)

        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setLevel(level_value)
        handler.setFormatter(logging.Formatter("%(message)s"))

        stream = logging.StreamHandler()
        stream.setLevel(level_value)
        stream.setFormatter(logging.Formatter("%(message)s"))

        root = logging.getLogger()
        root.handlers.clear()
        root.addHandler(handler)
        root.addHandler(stream)
        root.setLevel(level_value)

        structlog.configure(
            processors=[
                structlog.contextvars.merge_contextvars,
                structlog.processors.add_log_level,
                structlog.processors.TimeStamper(fmt="iso"),
                structlog.processors.StackInfoRenderer(),
                structlog.processors.format_exc_info,
                structlog.processors.JSONRenderer(),
            ],
            wrapper_class=structlog.make_filtering_bound_logger(level_value),
            logger_factory=structlog.stdlib.LoggerFactory(),
            cache_logger_on_first_use=True,
        )
        _CONFIGURED = True

    return structlog.get_logger().bind(app="ai_olap")
