"""
Тест фільтра розділів у А_ОтчетPL — 4 сценарії з різними значеннями параметра Розділи.

Запуск:
  C:\\Python313\\python.exe scripts/test/test_otchet_pl_sections.py --baseline   # без нового параметра (поточний стан)
  C:\\Python313\\python.exe scripts/test/test_otchet_pl_sections.py              # після правок (4 сценарії)
"""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import pywintypes  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

PERIOD_START = pywintypes.Time(datetime.datetime(2026, 2, 1, 0, 0, 0))
PERIOD_END = pywintypes.Time(datetime.datetime(2026, 2, 28, 23, 59, 59))


def fetch_table(report_func, *args):
    """Викликає функцію звіту з аргументами, повертає список рядків."""
    table = report_func(*args)
    rows = []
    n = table.Количество()
    for i in range(n):
        r = table.Получить(i)
        try:
            podr = str(r.Подразделение.Наименование) if r.Подразделение else ""
        except Exception:
            podr = ""
        rows.append({
            "Подразделение": podr,
            "СуммаPL": float(r.СуммаPL or 0),
            "СуммаЕРП": float(r.СуммаЕРП or 0),
            "СуммаКазнаПриход": float(r.СуммаДДСизКазныПриход or 0),
            "СуммаКазнаРасход": float(r.СуммаДДСизКазныРасход or 0),
        })
    return rows


def stats(rows, label=""):
    pl_rows = [r for r in rows if r["СуммаPL"] != 0]
    erp_rows = [r for r in rows if r["СуммаЕРП"] != 0]
    casa_rows = [r for r in rows if r["СуммаКазнаПриход"] != 0 or r["СуммаКазнаРасход"] != 0]
    sum_pl = sum(r["СуммаPL"] for r in rows)
    sum_erp = sum(r["СуммаЕРП"] for r in rows)
    sum_casa_p = sum(r["СуммаКазнаПриход"] for r in rows)
    sum_casa_r = sum(r["СуммаКазнаРасход"] for r in rows)
    print(f"  {label}:")
    print(f"    рядків усього: {len(rows)}  PL≠0: {len(pl_rows)}  ЕРП≠0: {len(erp_rows)}  Каса≠0: {len(casa_rows)}")
    print(f"    Σ PL: {sum_pl:>20,.2f}   Σ ЕРП: {sum_erp:>20,.2f}")
    print(f"    Σ Каса.Приход: {sum_casa_p:>14,.2f}   Σ Каса.Расход: {sum_casa_r:>14,.2f}")
    return {
        "rows": len(rows),
        "pl_rows": len(pl_rows), "erp_rows": len(erp_rows), "casa_rows": len(casa_rows),
        "sum_pl": sum_pl, "sum_erp": sum_erp,
        "sum_casa_p": sum_casa_p, "sum_casa_r": sum_casa_r,
    }


def main():
    baseline_mode = "--baseline" in sys.argv
    erp = connect_erp()
    report = erp.Отчеты.А_ОтчетPL.Создать()

    if baseline_mode:
        print("=== BASELINE: поточний .bsl без нового параметра ===")
        rows = fetch_table(report.ПолучитьОбъединенныеДанные, PERIOD_START, PERIOD_END, False)
        s = stats(rows, "default (3 параметри)")
        print()
        print("Зберегти ці цифри для порівняння після правок:")
        print(f"  rows={s['rows']}  pl≠0={s['pl_rows']}  erp≠0={s['erp_rows']}  casa≠0={s['casa_rows']}")
        return

    # === 4 сценарії після правок (.bsl з 6-ма параметрами) ===
    print("=== Сценарії після правок ===")
    print()

    print("[1] Розділи = (PL, ЕРП, Казна) — default")
    rows1 = fetch_table(report.ПолучитьОбъединенныеДанные, PERIOD_START, PERIOD_END, False, True, True, True)
    s1 = stats(rows1, "all-on")
    print()

    print("[2] Розділи = (PL) тільки")
    rows2 = fetch_table(report.ПолучитьОбъединенныеДанные, PERIOD_START, PERIOD_END, False, True, False, False)
    s2 = stats(rows2, "PL only")
    print()

    print("[3] Розділи = (PL, ЕРП) — без Кази (use case користувача)")
    rows3 = fetch_table(report.ПолучитьОбъединенныеДанные, PERIOD_START, PERIOD_END, False, True, True, False)
    s3 = stats(rows3, "PL+ЕРП без Кази")
    print()

    print("[4] Розділи = (Казна) тільки")
    rows4 = fetch_table(report.ПолучитьОбъединенныеДанные, PERIOD_START, PERIOD_END, False, False, False, True)
    s4 = stats(rows4, "Каса only")
    print()

    # === Перевірки ===
    print("=== Assertions ===")
    fails = []

    # [2] PL only — 0 ЕРП-рядків, 0 Каса-рядків
    if s2["erp_rows"] != 0:
        fails.append(f"[2] PL-only: ЕРП-рядків {s2['erp_rows']} замість 0")
    if s2["casa_rows"] != 0:
        fails.append(f"[2] PL-only: Каса-рядків {s2['casa_rows']} замість 0")

    # [3] PL+ЕРП без Кази — Каса має бути 0, PL/ЕРП НЕ змінюються
    if s3["casa_rows"] != 0:
        fails.append(f"[3] PL+ЕРП: Каса-рядків {s3['casa_rows']} замість 0")
    if abs(s3["sum_pl"] - s1["sum_pl"]) > 0.01:
        fails.append(f"[3] PL+ЕРП: sum_pl {s3['sum_pl']} ≠ baseline {s1['sum_pl']}")
    if abs(s3["sum_erp"] - s1["sum_erp"]) > 0.01:
        fails.append(f"[3] PL+ЕРП: sum_erp {s3['sum_erp']} ≠ baseline {s1['sum_erp']}")

    # [4] Каса only — 0 PL-сум, 0 ЕРП-сум
    if abs(s4["sum_pl"]) > 0.01:
        fails.append(f"[4] Каса-only: sum_pl {s4['sum_pl']} замість 0")
    if abs(s4["sum_erp"]) > 0.01:
        fails.append(f"[4] Каса-only: sum_erp {s4['sum_erp']} замість 0")

    if fails:
        print("FAIL:")
        for f in fails:
            print(f"  - {f}")
        sys.exit(1)
    print("ALL ASSERTIONS PASS")


if __name__ == "__main__":
    main()
