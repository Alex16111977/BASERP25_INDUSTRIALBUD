"""Step 15: Export PL knowledge to _Rarzrabotki/notebook/knowledge_PL/.

Pipeline for NotebookLM INDUSTRIALBUD_PL (af143439-3c76-42f8-a410-6367b5fd609f).
Uses V83.COMConnector (via _erp_query).

Usage:
    python 15_export_to_knowledge_pl.py                    # all 3 months + statics + deltas (default)
    python 15_export_to_knowledge_pl.py --period 2026-02   # single month
    python 15_export_to_knowledge_pl.py --static-only      # only catalog + dds mapping
    python 15_export_to_knowledge_pl.py --skip-static      # skip statics (faster)
"""
import argparse
import calendar
import sys
from datetime import datetime
from pathlib import Path

# _Rarzrabotki/Python/PnL — джерело config.py (CONN_ERP, EXCEL_FILES, JSON_DIR).
_PNL_DIR = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL")
sys.path.insert(0, str(_PNL_DIR))
# Поточна папка — для імпорту _export_pl_knowledge_helpers (sibling).
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config  # noqa: E402
from _export_pl_knowledge_helpers import (  # noqa: E402
    fetch_catalog, fetch_mapping,
    fetch_plan, fetch_fact_expenses, fetch_fact_income, fetch_cash,
    render_articles_catalog, render_dds_mapping,
    render_month_split, render_delta,
    load_excel_comments,
    UA_MONTHS,
)

KNOWLEDGE_DIR = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_PL")
RAW_SHEETS = config.JSON_DIR / "01_raw_sheets.json"


def month_range(ym):
    y, m = map(int, ym.split("-"))
    last = calendar.monthrange(y, m)[1]
    return f"{y:04d}-{m:02d}-01", f"{y:04d}-{m:02d}-{last:02d}"


def label_for(ym):
    for f in config.EXCEL_FILES:
        if f["month_header"][:7] == ym:
            return f["label"]
    return ym


def month_filenames(ym):
    """Повертає dict bucket → filename для 6 split-файлів місяця."""
    y, m = ym.split("-")
    base = f"pl_dump_{y}_{m}_"
    return {
        "summary":         f"{base}01_summary.md",
        "income":          f"{base}02_income_revenue.md",
        "cost":            f"{base}03_cost_of_goods.md",
        "opex":            f"{base}04_opex_admin.md",
        "marketing_fin":   f"{base}05_marketing_fin.md",
        "cash_anomalies":  f"{base}06_cash_anomalies.md",
    }


def delta_filename(prev_ym, curr_ym):
    y1, m1 = prev_ym.split("-")
    y2, m2 = curr_ym.split("-")
    eng = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]
    return f"pl_dump_delta_{eng[int(m2)-1]}{y2}_vs_{eng[int(m1)-1]}{y1}.md"


def all_periods():
    return sorted({f["month_header"][:7] for f in config.EXCEL_FILES})


def export_static(log):
    log("  fetching catalog (А_Статьи_PL) ...")
    catalog = fetch_catalog()
    log(f"    {len(catalog)} articles")
    log("  rendering pl_articles_catalog.md ...")
    md = render_articles_catalog(catalog)
    out = KNOWLEDGE_DIR / "pl_articles_catalog.md"
    out.write_text(md, encoding="utf-8")
    log(f"    written: {out.name} ({len(md):,} chars)")

    log("  fetching mapping (ТЧ Статьи) ...")
    mapping = fetch_mapping()
    log(f"    {len(mapping)} rows")
    log("  rendering pl_dds_mapping.md ...")
    md = render_dds_mapping(mapping)
    out = KNOWLEDGE_DIR / "pl_dds_mapping.md"
    out.write_text(md, encoding="utf-8")
    log(f"    written: {out.name} ({len(md):,} chars)")
    return mapping


def fetch_month(ym, log):
    date_from, date_to = month_range(ym)
    log(f"  period: {date_from} … {date_to}")
    log("  fetch plan ...")
    plan = fetch_plan(date_from, date_to, include_children=False)
    log(f"    plan rows: {len(plan)}")
    log("  fetch fact expenses ...")
    fr = fetch_fact_expenses(date_from, date_to)
    log(f"    fact_rash rows: {len(fr)}")
    log("  fetch fact income ...")
    fd = fetch_fact_income(date_from, date_to)
    log(f"    fact_doh rows: {len(fd)}")
    log("  fetch cash ...")
    c = fetch_cash(date_from, date_to)
    log(f"    cash rows: {len(c)}")
    return plan, fr, fd, c


def export_month(ym, mapping, log):
    label = label_for(ym)
    log(f"  === {label} ({ym}) ===")
    data = fetch_month(ym, log)
    plan, fr, fd, c = data

    excel = load_excel_comments(RAW_SHEETS, label)
    log(f"  excel sheet/article keys: {len(excel)}")

    date_from, date_to = month_range(ym)
    # U1: render_month_split повертає dict з 6 buckets
    bucket_to_md = render_month_split(ym, date_from, date_to, plan, fr, fd, c, mapping, excel)
    filenames = month_filenames(ym)

    for bucket, md in bucket_to_md.items():
        out = KNOWLEDGE_DIR / filenames[bucket]
        out.write_text(md, encoding="utf-8")
        log(f"  written: {out.name} ({len(md):,} chars)")

    return data


def export_delta(prev_ym, curr_ym, prev_data, curr_data, mapping, log):
    p_label = label_for(prev_ym)
    c_label = label_for(curr_ym)
    log(f"  delta {c_label} vs {p_label} ...")
    md = render_delta(
        (*prev_data, mapping),
        (*curr_data, mapping),
        p_label, c_label,
    )
    out = KNOWLEDGE_DIR / delta_filename(prev_ym, curr_ym)
    out.write_text(md, encoding="utf-8")
    log(f"  written: {out.name} ({len(md):,} chars)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", help="YYYY-MM to export (default: all 3)")
    ap.add_argument("--skip-static", action="store_true")
    ap.add_argument("--static-only", action="store_true")
    args = ap.parse_args()

    KNOWLEDGE_DIR.mkdir(parents=True, exist_ok=True)
    LOG = Path(__file__).parent / "_export_run_log.txt"
    LOG.write_text("", encoding="utf-8")
    def log(msg):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        print(line, flush=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
            f.flush()

    log(f"Target: {KNOWLEDGE_DIR}")

    mapping = None
    if not args.skip_static:
        log("=== STATIC FILES ===")
        mapping = export_static(log)
    else:
        log("fetching mapping (needed for month rendering) ...")
        mapping = fetch_mapping()
        log(f"  {len(mapping)} rows")

    if args.static_only:
        log("static-only — done.")
        return

    periods = [args.period] if args.period else all_periods()
    log(f"=== MONTHLY ({len(periods)} periods) ===")

    data_cache = {}
    for ym in periods:
        d = export_month(ym, mapping, log)
        data_cache[ym] = d

    # Deltas between consecutive periods
    ordered = all_periods()
    if args.period and args.period in ordered:
        i = ordered.index(args.period)
        if i > 0:
            prev_ym = ordered[i - 1]
            if prev_ym not in data_cache:
                log(f"fetching prev month {prev_ym} for delta ...")
                data_cache[prev_ym] = fetch_month(prev_ym, log)
            log(f"=== DELTA ===")
            export_delta(prev_ym, args.period, data_cache[prev_ym], data_cache[args.period], mapping, log)
    else:
        log(f"=== DELTAS ===")
        for i in range(1, len(ordered)):
            prev_ym = ordered[i - 1]
            curr_ym = ordered[i]
            if prev_ym in data_cache and curr_ym in data_cache:
                export_delta(prev_ym, curr_ym, data_cache[prev_ym], data_cache[curr_ym], mapping, log)

    log("DONE")


if __name__ == "__main__":
    main()
