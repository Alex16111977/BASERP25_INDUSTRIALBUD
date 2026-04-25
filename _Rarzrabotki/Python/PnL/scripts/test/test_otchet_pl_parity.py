"""
Інтеграційний тест Reports.А_ОтчетPL за лютий 2026.

Перевіряє (до копійки):
  1. drill-down не містить РасчетСебестоимостиТоваров.
  2. (Глобино-2 / PL "Будівельні матеріали" / СтатьяPL.Группа) сума ЕРП
     містить наш CoGS = 13 687 175.04 ₴ (за умови що ДДС
     "Себестоимость реализации" має прапор А_ПриёмникСебестоимостиПродажPL
     і додана в ТЧ Статьи PL "Будівельні матеріали").
  3. Виручка по Глобино-2 = 38 432 968.66 ₴ (с НДС).

Запуск:
  C:\\Python313\\python.exe scripts/test/test_otchet_pl_parity.py
"""
import datetime
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pywintypes  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

PERIOD_START = pywintypes.Time(datetime.datetime(2026, 2, 1, 0, 0, 0))
PERIOD_END = pywintypes.Time(datetime.datetime(2026, 2, 28, 23, 59, 59))

GLOBINO_2 = "Глобино-2"
PL_BUILDING = "Строительные материалы"
# У ТЧ Статьи PL "Будівельні матеріали" дві ДДС:
#   - "Строительные материалы" (через ПрочиеРасходы) → +7 176 926.64
#   - "Себестоимость реализации" (наш CoGS, прапор) → +13 687 175.04
# Разом = 20 864 101.68 ₴ (за рішенням фінансиста — обидві суми реальні витрати).
ETALON_PL_BUILDING_GLOBINO = 20_864_101.68
ETALON_REVENUE = 38_432_968.66
FORBIDDEN_DOC_TYPE = "РасчетСебестоимостиТоваров"


def main():
    erp = connect_erp()
    report = erp.Отчеты.А_ОтчетPL.Создать()
    table = report.ПолучитьОбъединенныеДанные(PERIOD_START, PERIOD_END, False)

    rows = []
    n = table.Количество()
    for i in range(n):
        r = table.Получить(i)
        try:
            podr_name = str(r.Подразделение.Наименование) if r.Подразделение else ""
        except Exception:
            podr_name = ""
        try:
            statia_name = str(r.СтатьяPL.Наименование) if r.СтатьяPL else ""
        except Exception:
            statia_name = ""
        rows.append({
            "Подразделение": podr_name,
            "СтатьяPL": statia_name,
            "СуммаPL": float(r.СуммаPL or 0),
            "СуммаЕРП": float(r.СуммаЕРП or 0),
            "Документ": str(r.Документ),
        })

    print(f"Усього рядків у звіті: {len(rows)}")

    # === Assertion 1: drill-down не містить РасчетСеб ===
    bad = [r for r in rows if FORBIDDEN_DOC_TYPE in r["Документ"]]
    if bad:
        print(f"FAIL #1: знайдено {len(bad)} {FORBIDDEN_DOC_TYPE}")
        sys.exit(1)
    print(f"OK #1: drill-down чистий, немає {FORBIDDEN_DOC_TYPE}")

    # === Assertion 2: PL "Будівельні матеріали" Глобино-2 = 20 864 101.68 ===
    # (включає ПрочиеРасходы 7.2M + наш CoGS 13.7M через дві ДДС у ТЧ Статьи)
    pl_sum = sum(r["СуммаЕРП"] for r in rows
                 if r["Подразделение"] == GLOBINO_2
                 and r["СтатьяPL"] == PL_BUILDING)
    delta_pl = abs(pl_sum - ETALON_PL_BUILDING_GLOBINO)
    print(f"#2 PL {PL_BUILDING} Глобино-2: {pl_sum:,.2f} ₴ vs еталон {ETALON_PL_BUILDING_GLOBINO:,.2f} ₴ → Δ {delta_pl:,.2f}")
    if delta_pl > 0.01:
        print(f"FAIL #2: розбіжність {delta_pl:,.2f} ₴ > 0.01.")
        sys.exit(1)
    print("OK #2: PL «Будівельні матеріали» до копійки")

    # === Assertion 3: Виручка Глобино-2 = 38 432 968.66 ===
    # Виручка лежить у PL-статтях через СтатьяДоходов = "Выручка от продаж"
    revenue = sum(r["СуммаЕРП"] for r in rows
                  if r["Подразделение"] == GLOBINO_2
                  and "Выручка" in r["СтатьяPL"])
    delta_rev = abs(revenue - ETALON_REVENUE)
    print(f"#3 Виручка Глобино-2: {revenue:,.2f} ₴ vs еталон {ETALON_REVENUE:,.2f} ₴ → Δ {delta_rev:,.2f}")
    if delta_rev > 0.01:
        print(f"FAIL #3: розбіжність {delta_rev:,.2f} ₴ > 0.01.")
        sys.exit(1)
    print("OK #3: Виручка до копійки")

    print()
    print("ALL ASSERTIONS PASS (до копійки)")


if __name__ == "__main__":
    main()
