"""Ai_Olap CLI — entry point for ETL runs.

Usage:
    python main.py
        Default: повний прогон за ВСІ проведені документи (без period filter) —
        validate -> усі робочі Dim (dim_catalogs, dim_pap_articles,
        dim_denezhnye_sredstva) -> усі робочі Fact (fact_pnl, fact_cashflow,
        fact_balance). Усі Dim та Fact перезавантажуються повністю
        (TRUNCATE + INSERT, усі періоди). Зупиняється на першій помилці.
        Виключені (зламані/неповні, падали б весь ланцюг): dim_documents,
        fact_cf_balance — запускати/чинити окремо через --run-once.

    python main.py --period 2026-02
        Той самий ланцюжок, але Fact-pipelines обмежуються одним місяцем
        (idempotent: DELETE WHERE Period_Month=? + INSERT).

    python main.py --validate
        Валідує всі pipelines/*.json і виходить.

    python main.py --run-once <pipeline_id> [--period YYYY-MM]
        Запуск одного pipeline за file stem.

    python main.py --scheduled
        APScheduler daemon (BlockingScheduler) — для production (зараз НЕ
        використовуємо, поки йде ручне тестування).

    python main.py --refresh-mapping
        Скидає кеш mapping_resolver (НЕ регенерує JSON!).
        Регенерація mapping: `python mapping/refresh_mapping.py`.

Examples:
    python main.py                           # full reload всіх періодів
    python main.py --period 2026-02          # per-month reload лютого
    python main.py --validate
    python main.py --run-once dim_catalogs
    python main.py --run-once fact_pnl --period 2026-02
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

# Pipeline id-и які виконує `--all` режим, у порядку.
ALL_DEFAULT_PIPELINES: list[str] = [
    "dim_factory",                 # 7 справочников-измерений, full_reload
    "fact_planfact_proizvodstvo",  # факт из РС А_ПланФактПроизводство_Свод
]



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


def _run_one(pid: str, period: dt.date | None) -> int:
    target = PIPELINES_DIR / f"{pid}.json"
    if not target.exists():
        print(f"Pipeline file {target} not found.", file=sys.stderr)
        print("Available:", ", ".join(p.stem for p in discover_pipelines(PIPELINES_DIR)))
        return 2
    cfg = load_pipeline_config(target)
    validate_pipeline(cfg, source=target.name)
    needs_period = any(
        step.get("loader", {}).get("mode") == "idempotent_period"
        for step in cfg.get("steps", [])
    )
    total = run_pipeline(cfg, period=period if needs_period else None, script=pid)
    print(f"  {pid}: {total} rows loaded.")
    return 0


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
    period = _parse_period(args.period)
    rc = _run_one(args.run_once, period)
    return rc


def cmd_all(args) -> int:
    """Default mode: validate + dim_catalogs + fact_pnl + fact_cashflow.

    --period не передано → full reload (TRUNCATE + INSERT всі періоди).
    --period YYYY-MM   → per-month idempotent reload (DELETE WHERE + INSERT).
    """
    period = _parse_period(args.period)
    label = f"period={period:%Y-%m}" if period else "ALL periods (full reload)"
    print(f"=== Ai_Olap full update — {label} ===\n")

    print(f"[1/{1 + len(ALL_DEFAULT_PIPELINES)}] validate")
    rc = cmd_validate(args)
    if rc != 0:
        print("validate failed — aborting", file=sys.stderr)
        return rc

    for n, pid in enumerate(ALL_DEFAULT_PIPELINES, start=2):
        print(f"\n[{n}/{1 + len(ALL_DEFAULT_PIPELINES)}] {pid}")
        rc = _run_one(pid, period)
        if rc != 0:
            print(f"{pid} failed (rc={rc}) — aborting", file=sys.stderr)
            return rc

    print(f"\n=== Ai_Olap full update — DONE ({label}) ===")
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
        description="ETL orchestrator: BaseERP -> OlapFactory. "
                    "Default (no flags): full update for previous month.",
    )
    g = parser.add_mutually_exclusive_group(required=False)
    g.add_argument("--validate", action="store_true", help="Validate all pipelines/*.json")
    g.add_argument("--run-once", metavar="PIPELINE_ID", help="Run one pipeline by file stem")
    g.add_argument("--scheduled", action="store_true", help="Start APScheduler daemon")
    g.add_argument("--refresh-mapping", action="store_true", help="Drop mapping cache")
    parser.add_argument("--period", metavar="YYYY-MM", help="Period (defaults to previous month for --all)")
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
        return cmd_all(args)
    except ETLException as exc:
        print(f"ETL error: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
