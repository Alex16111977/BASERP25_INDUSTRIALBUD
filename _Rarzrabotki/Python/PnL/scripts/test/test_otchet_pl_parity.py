"""
Запускає функцію Отчет.А_ОтчетPL.ПолучитьОбъединенныеДанные за лютий 2026
через COM-конект до ЕРП.

Перевіряє:
  1. Жоден Документ-Регістратор не = РасчетСебестоимостиТоваров
     (drill-down Виручки чистий від службових агрегацій).
  2. Друкує підсумок по підрозділах: PL-сума, ЕРП-сума, Δ.

CoGS перевірка: інформаційна. Якщо у системі є ДДС з прапором
А_ПриёмникСебестоимостиПродажPL і ця ДДС додана у ТЧ Статьи однієї з
PL-статей — CoGS з'явиться у звіті.
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

    print(f"Усього рядків: {len(rows)}")
    print()

    bad = [r for r in rows if FORBIDDEN_DOC_TYPE in r["Документ"]]
    if bad:
        print(f"FAIL: знайдено {len(bad)} документів {FORBIDDEN_DOC_TYPE}")
        for r in bad[:5]:
            print(f"  {r['Подразделение']} / {r['СтатьяPL']} / ЕРП {r['СуммаЕРП']:,.2f} / {r['Документ']}")
        sys.exit(1)
    print(f"OK: drill-down чистий, немає {FORBIDDEN_DOC_TYPE}")

    by_pod = defaultdict(lambda: {"PL": 0.0, "ЕРП": 0.0})
    for r in rows:
        by_pod[r["Подразделение"]]["PL"] += r["СуммаPL"]
        by_pod[r["Подразделение"]]["ЕРП"] += r["СуммаЕРП"]
    print()
    print(f"{'Подразделение':<40} {'PL':>15} {'ЕРП':>15} {'delta':>15}")
    for name in sorted(by_pod):
        v = by_pod[name]
        d = v["ЕРП"] - v["PL"]
        print(f"{name[:40]:<40} {v['PL']:>15,.2f} {v['ЕРП']:>15,.2f} {d:>15,.2f}")


if __name__ == "__main__":
    main()
