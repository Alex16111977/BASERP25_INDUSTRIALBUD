"""Compute monthly aggregates for pl_faq.md: top-N подрозділів, статтей, контрагентів.

Run: python _compute_pl_aggregates.py > _pl_aggregates.json
"""
import json
import sys
from pathlib import Path
from collections import defaultdict

# _Rarzrabotki/Python/PnL — джерело config.py (CONN_ERP, EXCEL_FILES).
_PNL_DIR = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL")
sys.path.insert(0, str(_PNL_DIR))
# Поточна папка — для імпорту _export_pl_knowledge_helpers (sibling).
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config  # noqa: E402
from _export_pl_knowledge_helpers import (  # noqa: E402
    fetch_plan, fetch_fact_expenses, fetch_fact_income, fetch_cash,
    fetch_mapping,
)

LOG = Path(__file__).parent / "_aggregates_log.txt"
def log(m):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(m + "\n"); f.flush()
LOG.write_text("", encoding="utf-8")

import calendar

def month_range(ym):
    y, m = map(int, ym.split("-"))
    last = calendar.monthrange(y, m)[1]
    return f"{y:04d}-{m:02d}-01", f"{y:04d}-{m:02d}-{last:02d}"


def aggregate_month(ym, mapping_dds_to_pl):
    date_from, date_to = month_range(ym)
    log(f"[{ym}] plan...")
    plan = fetch_plan(date_from, date_to, include_children=False)
    log(f"  plan rows: {len(plan)}")

    log(f"[{ym}] fact expenses...")
    fr = fetch_fact_expenses(date_from, date_to)
    log(f"  rash rows: {len(fr)}")

    log(f"[{ym}] cash...")
    c = fetch_cash(date_from, date_to)
    log(f"  cash rows: {len(c)}")

    # Total plan
    total_plan = sum(r["СуммаPL"] or 0.0 for r in plan)
    total_fact_rash = sum(r["Сумма"] or 0.0 for r in fr)
    total_cash_in = sum(r["Приход"] or 0.0 for r in c)
    total_cash_out = sum(r["Расход"] or 0.0 for r in c)

    # Top-N підрозділів за планом
    plan_per_podr = defaultdict(lambda: {"plan": 0.0, "docs": set()})
    for r in plan:
        podr = r["Подразделение"] or "(без)"
        plan_per_podr[podr]["plan"] += r["СуммаPL"] or 0.0
        if r["ДокНомер"]:
            plan_per_podr[podr]["docs"].add(r["ДокНомер"])
    top_podr_plan = sorted(
        ({"подр": p, "план": d["plan"], "n_docs": len(d["docs"])} for p, d in plan_per_podr.items()),
        key=lambda x: -x["план"]
    )[:15]

    # Top-N контрагентів за витратами (fact rash)
    kontr_rash = defaultdict(float)
    kontr_articles = defaultdict(set)  # контрагент → set of PL-статей
    for r in fr:
        k = r["Контрагент"] or "(без)"
        if k == "(без)":
            continue
        kontr_rash[k] += r["Сумма"] or 0.0
        # Resolve PL-стаття from ДДС
        dds = r["ДДСНаим"]
        pl = mapping_dds_to_pl.get(dds, "(без PL)")
        kontr_articles[k].add(pl)
    top_kontr_rash = sorted(
        ({"контрагент": k, "сума": v, "статті": sorted(kontr_articles[k])[:5]} for k, v in kontr_rash.items()),
        key=lambda x: -x["сума"]
    )[:15]

    # Top-N статей за планом
    art_plan = defaultdict(float)
    art_plan_comments = defaultdict(list)
    for r in plan:
        a = r["СтатьяНаим"] or "(без)"
        art_plan[a] += r["СуммаPL"] or 0.0
        if r["Комментарий"] and r["Комментарий"].strip():
            art_plan_comments[a].append(r["Комментарий"].strip()[:100])
    top_art_plan = sorted(
        ({"стаття": a, "план": v, "коменти": list(set(art_plan_comments[a]))[:3]} for a, v in art_plan.items()),
        key=lambda x: -x["план"]
    )[:15]

    # Top-N статей за фактом витрат
    art_fact = defaultdict(float)
    for r in fr:
        dds = r["ДДСНаим"]
        pl = mapping_dds_to_pl.get(dds, "(без PL)")
        art_fact[pl] += r["Сумма"] or 0.0
    top_art_fact = sorted(
        ({"стаття": a, "факт": v} for a, v in art_fact.items()),
        key=lambda x: -x["факт"]
    )[:15]

    # Top-N документів витрат
    top_docs_rash = sorted(fr, key=lambda r: -(r["Сумма"] or 0.0))[:10]
    top_docs_rash_out = [
        {
            "тип": (r["ТипДокумента"] or "").replace("Документ.", ""),
            "номер": (r["ДокНомер"] or "").strip(),
            "дата": str(r["ДокДата"])[:10] if r["ДокДата"] else "",
            "контрагент": r["Контрагент"] or "",
            "підр": r["Подразделение"] or "",
            "ддс": r["ДДСНаим"] or "",
            "сума": r["Сумма"] or 0.0,
        } for r in top_docs_rash
    ]

    # Top-N каса припливів
    top_cash_in = sorted(c, key=lambda r: -(r["Приход"] or 0.0))[:10]
    top_cash_in_out = [
        {
            "тип": (r["ТипДокумента"] or "").replace("Документ.", ""),
            "номер": (r["ДокНомер"] or "").strip(),
            "дата": str(r["ДокДата"])[:10] if r["ДокДата"] else "",
            "контрагент": r["Контрагент"] or "",
            "підр": r["Подразделение"] or "",
            "ддс": r["ДДСНаим"] or "",
            "сума": r["Приход"] or 0.0,
        } for r in top_cash_in if r["Приход"] and r["Приход"] > 0
    ]

    return {
        "period": ym,
        "totals": {
            "plan": total_plan,
            "fact_rash": total_fact_rash,
            "cash_in": total_cash_in,
            "cash_out": total_cash_out,
            "n_plan_docs": len(set(r["ДокНомер"] for r in plan if r["ДокНомер"])),
            "n_rash_rows": len(fr),
            "n_cash_rows": len(c),
        },
        "top_podr_plan": top_podr_plan,
        "top_kontr_rash": top_kontr_rash,
        "top_art_plan": top_art_plan,
        "top_art_fact": top_art_fact,
        "top_docs_rash": top_docs_rash_out,
        "top_cash_in": top_cash_in_out,
    }


def main():
    log("fetching mapping...")
    mapping = fetch_mapping()
    mapping_dds_to_pl = {}
    for r in mapping:
        tc = r["ДДСТЧНаим"]
        if tc and tc not in mapping_dds_to_pl:
            mapping_dds_to_pl[tc] = r["PLНаим"]
    log(f"  mapping has {len(mapping_dds_to_pl)} ДДС→PL pairs")

    # 5 актуальних місяців (грудень 2025 — квітень 2026). Розширено 2026-05-21
    # після завантаження 2024 і 2025 років. Для глибших ретроспектив — окремі monthly dumps.
    result = {}
    for ym in ["2025-12", "2026-01", "2026-02", "2026-03", "2026-04"]:
        result[ym] = aggregate_month(ym, mapping_dds_to_pl)
        log(f"[{ym}] done")

    out = Path(__file__).parent / "_pl_aggregates.json"
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    log(f"wrote {out}")


if __name__ == "__main__":
    main()
