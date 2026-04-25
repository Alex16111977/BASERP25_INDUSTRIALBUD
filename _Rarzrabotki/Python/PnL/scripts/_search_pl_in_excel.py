"""Ищет 3 PL-статьи без ДДС во ВСЕХ 3 Excel-файлах по всем листам.
Помогает финансисту решить — оставить ли эти статьи в PL-форме или убрать.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config

TARGET_ARTICLES = [
    "Доход от прочей операционной деятельности",
    "Расход от ивест. Деятельности (ОС)",
    "Финансовые расходы",
]


def main():
    raw = json.loads((config.JSON_DIR / "01_raw_sheets.json").read_text(encoding="utf-8"))

    hits = {name: [] for name in TARGET_ARTICLES}

    for period in raw:
        period_label = period.get("label", period.get("period", "?"))
        for sh in period["sheets"]:
            sname = sh["sheet_name"]
            for r in sh.get("rows", []):
                art = r.get("article", "").strip()
                for target in TARGET_ARTICLES:
                    # смотрим по схожести начала имени (финансист может писать кратко)
                    if art.lower().startswith(target.lower()[:25]) or target.lower().startswith(art.lower()[:25]):
                        tot = float(r.get("total", 0) or 0)
                        f1 = float(r.get("sum_f1", 0) or 0)
                        f2 = float(r.get("sum_f2", 0) or 0)
                        if tot != 0 or f1 != 0 or f2 != 0:
                            hits[target].append({
                                "период": period_label,
                                "лист": sname,
                                "статья_в_excel": art,
                                "Ф1": f1, "Ф2": f2, "Сумма": tot,
                                "комментарий": (r.get("comment") or "")[:80],
                            })

    print("\n=== Поиск 3 PL-статей БЕЗ ДДС в Excel ===\n")
    for target, found in hits.items():
        print(f"━━━ «{target}» ━━━")
        if not found:
            print("  НЕ НАЙДЕНО ни в одном листе (за 3 месяца)")
        else:
            print(f"  Найдено: {len(found)} вхождений")
            total_sum = sum(h["Сумма"] for h in found)
            print(f"  Суммарно: {total_sum:,.2f} ₴")
            # первые 10
            for h in found[:15]:
                print(f"    {h['период']:12} / {h['лист'][:20]:20} | Ф1={h['Ф1']:>10.0f} Ф2={h['Ф2']:>10.0f} Итого={h['Сумма']:>10.0f}")
            if len(found) > 15:
                print(f"    ... ещё {len(found)-15} строк")
        print()


if __name__ == "__main__":
    main()
