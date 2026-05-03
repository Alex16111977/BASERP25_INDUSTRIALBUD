"""Ai_Olap CLI — entry point for ETL runs.

Usage:
    python main.py --validate
        Validate every pipelines/*.json file.

    python main.py --run-once <pipeline_id> [--period YYYY-MM]
        Run a single pipeline by id (filename stem). For Fact pipelines
        --period selects the month being loaded.

    python main.py --scheduled
        Start the APScheduler daemon (BlockingScheduler).

    python main.py --refresh-mapping
        Drop the mapping_resolver cache so the next call reads
        mapping/baserp_storage.json fresh.

Examples:
    python main.py --validate
    python main.py --run-once dim_catalogs
    python main.py --run-once fact_pnl --period 2026-02
    python main.py --scheduled
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from ai_olap.config.loader import discover_pipelines, load_pipeline_config  # noqa: E402
from ai_olap.config.validator import validate_pipeline  # noqa: E402
from ai_olap.core.exceptions import ETLException, ValidationError  # noqa: E402
from ai_olap.core.logging_setup import setup_logging  # noqa: E402
from ai_olap.orchestrator.runner import run_pipeline  # noqa: E402
from ai_olap.orchestrator.scheduler import EtlScheduler  # noqa: E402
from ai_olap.utils.mapping_resolver import reload_cache  # noqa: E402

PIPELINES_DIR = PROJECT_ROOT / "pipelines"


def _parse_period(s: str | None) -> dt.date | None:
    if not s:
        return None
    try:
        if len(s) == 7:  # YYYY-MM
            y, m = map(int, s.split("-"))
            return dt.date(y, m, 1)
        return dt.date.fromisoformat(s)
    except Exception as exc:
        raise ValidationError(f"Invalid --period {s!r}; use YYYY-MM or YYYY-MM-DD") from exc


def cmd_validate(_args) -> int:
    paths = discover_pipelines(PIPELINES_DIR)
    if not paths:
        print(f"No pipelines found in {PIPELINES_DIR}")
        return 1
    for p in paths:
        cfg = load_pipeline_config(p)
        validate_pipeline(cfg, source=p.name)
        print(f"OK  {p.name}  ({len(cfg['steps'])} steps, schedule={cfg.get('schedule')})")
    print(f"\nAll {len(paths)} pipelines valid.")
    return 0


def cmd_run_once(args) -> int:
    pid = args.run_once
    period = _parse_period(args.period)
    target = PIPELINES_DIR / f"{pid}.json"
    if not target.exists():
        print(f"Pipeline file {target} not found.", file=sys.stderr)
        print("Available:", ", ".join(p.stem for p in discover_pipelines(PIPELINES_DIR)))
        return 2
    cfg = load_pipeline_config(target)
    validate_pipeline(cfg, source=target.name)
    total = run_pipeline(cfg, period=period, script=pid)
    print(f"\n{pid}: {total} rows loaded.")
    return 0


def cmd_scheduled(_args) -> int:
    sched = EtlScheduler(pipelines_dir=PIPELINES_DIR)
    n = sched.register_jobs()
    if n == 0:
        print("No scheduled pipelines (configs without 'schedule' are skipped).")
        return 1
    print(f"Registered {n} job(s): {sched.list_jobs()}")
    print("Press Ctrl-C to stop.")
    try:
        sched.start()
    except (KeyboardInterrupt, SystemExit):
        sched.shutdown(wait=False)
    return 0


def cmd_refresh_mapping(_args) -> int:
    reload_cache()
    print("Mapping cache cleared. Next pipeline run will re-read mapping/baserp_storage.json.")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="ai_olap",
        description="ETL orchestrator: BaseERP -> OlapBASERP.",
    )
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--validate", action="store_true", help="Validate all pipelines/*.json")
    g.add_argument("--run-once", metavar="PIPELINE_ID", help="Run one pipeline by file stem")
    g.add_argument("--scheduled", action="store_true", help="Start APScheduler daemon")
    g.add_argument("--refresh-mapping", action="store_true", help="Drop mapping cache")
    parser.add_argument("--period", metavar="YYYY-MM", help="Period for Fact pipelines (1st of month)")
    args = parser.parse_args(argv)

    setup_logging()

    try:
        if args.validate:
            return cmd_validate(args)
        if args.run_once:
            return cmd_run_once(args)
        if args.scheduled:
            return cmd_scheduled(args)
        if args.refresh_mapping:
            return cmd_refresh_mapping(args)
    except ETLException as exc:
        print(f"ETL error: {exc}", file=sys.stderr)
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
