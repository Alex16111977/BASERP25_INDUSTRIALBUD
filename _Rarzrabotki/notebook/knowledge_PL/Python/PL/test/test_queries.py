"""Перевірити що всі 6 запитів з _export_pl_knowledge_helpers виконуються.

Коли запускати: коли у 1С-конфігурації зміняться назви реквізитів документів
(Контрагент, Дата, Номер, СтатьяРасходов, тощо) або структура регістрів накопичення.

Якщо тест FAIL — дивись у _export_pl_knowledge_helpers.py змінну відповідного
Q_* запиту і синхронізуй з реальною конфігурацією.

Usage:
    python test_queries.py
    python test_queries.py --period 2026-02   # тестувати з конкретним періодом
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

LOG = Path(__file__).parent / "_test_queries_log.txt"
LOG.write_text("", encoding="utf-8")


def log(msg):
    print(msg, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
        f.flush()


def run(name, fn, *args, **kwargs):
    log(f"[{name}] ...")
    try:
        r = fn(*args, **kwargs)
        log(f"  OK — {len(r)} рядків")
        if r:
            sample = {k: v for k, v in list(r[0].items())[:5]}
            log(f"  sample keys: {sample}")
        return r
    except Exception as e:
        log(f"  FAIL: {e}")
        import traceback
        log(traceback.format_exc())
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", default="2026-02", help="YYYY-MM для fetch_plan/fact/cash")
    args = ap.parse_args()

    log("=== test_queries.py — всі 6 запитів pipeline ===")
    log(f"Період для fetch_*: {args.period}")

    from _export_pl_knowledge_helpers import (
        fetch_catalog, fetch_mapping,
        fetch_plan, fetch_fact_expenses, fetch_fact_income, fetch_cash,
    )

    y, m = args.period.split("-")
    import calendar
    last = calendar.monthrange(int(y), int(m))[1]
    date_from = f"{y}-{m}-01"
    date_to = f"{y}-{m}-{last:02d}"

    ok = 0
    for name, fn, args_ in [
        ("fetch_catalog", fetch_catalog, ()),
        ("fetch_mapping", fetch_mapping, ()),
        ("fetch_plan", fetch_plan, (date_from, date_to, False)),
        ("fetch_fact_expenses", fetch_fact_expenses, (date_from, date_to)),
        ("fetch_fact_income", fetch_fact_income, (date_from, date_to)),
        ("fetch_cash", fetch_cash, (date_from, date_to)),
    ]:
        r = run(name, fn, *args_)
        if r is not None:
            ok += 1

    log("")
    log(f"=== РЕЗУЛЬТАТ: {ok}/6 запитів пройшли ===")
    if ok == 6:
        log("Усі запити OK — pipeline готовий до повного експорту.")
        log("Наступний крок: python ../15_export_to_knowledge_pl.py")
    else:
        log(f"{6 - ok} запитів провалилися. Дивись деталі вище.")
        log("Частини, які точно залежать від 1С-конфігурації:")
        log("  Q_CATALOG, Q_MAPPING — структура Справочник.А_Статьи_PL + ТЧ Статьи")
        log("  Q_PLAN — структура Документ.А_ОтчетPL.ДанныеОтчета")
        log("  Q_FACT_EXPENSES — реквізити документів-регістраторів ПрочиеРасходы")
        log("  Q_FACT_INCOME_FR/PD — реквізити ФинансовыеРезультаты + ПрочиеДоходы")
        log("  Q_CASH — реквізити А_ДвиженияДенегИзКазны + ДокументДвиженияКазны")
        sys.exit(1)


if __name__ == "__main__":
    main()
