"""
Порівнюємо 3 суми Виручки за лютий 2026:
  1. ФР як є (поточний код А_ОтчетPL)
  2. ФР з фільтром "не РасчетСебестоимости"
  3. ВиР (ВыручкаИСебестоимостьПродаж.СуммаВыручкиБезНДС)

Якщо (2) ≈ (3) ≈ (1) — фільтр у ФР безпечний (РасчетСеб переписує).
Якщо (2) << (1) — фільтр губить реальні дані.
"""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pywintypes  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

PERIOD_START = pywintypes.Time(datetime.datetime(2026, 2, 1, 0, 0, 0))
PERIOD_END = pywintypes.Time(datetime.datetime(2026, 2, 28, 23, 59, 59))

QUERIES = {
    "1_FR_as_is": """
ВЫБРАТЬ СУММА(Ф.ДоходыОборот) КАК Сумма
ИЗ РегистрНакопления.ФинансовыеРезультаты.Обороты(&НачалоПериода, &КонецПериода, , ) КАК Ф
ГДЕ Ф.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
""",
    "2_FR_exclude_RaschetSeb": """
ВЫБРАТЬ СУММА(Ф.ДоходыОборот) КАК Сумма
ИЗ РегистрНакопления.ФинансовыеРезультаты.Обороты(&НачалоПериода, &КонецПериода, Регистратор, ) КАК Ф
ГДЕ Ф.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
    И НЕ (Ф.Регистратор ССЫЛКА Документ.РасчетСебестоимостиТоваров)
""",
    "3_VIS_revenue": """
ВЫБРАТЬ СУММА(Выр.СуммаВыручкиБезНДСОборот) КАК Сумма
ИЗ РегистрНакопления.ВыручкаИСебестоимостьПродаж.Обороты(&НачалоПериода, &КонецПериода, , ) КАК Выр
""",
}


def fetch(erp, text):
    q = erp.NewObject("Запрос")
    q.Текст = text
    q.УстановитьПараметр("НачалоПериода", PERIOD_START)
    q.УстановитьПараметр("КонецПериода", PERIOD_END)
    sel = q.Выполнить().Выбрать()
    sel.Следующий()
    return float(sel.Сумма or 0)


def main():
    erp = connect_erp()
    sums = {}
    for k, q in QUERIES.items():
        try:
            sums[k] = fetch(erp, q)
            print(f"  {k}: {sums[k]:>20,.2f}")
        except Exception as e:
            print(f"  {k}: ERROR {e}")
            sums[k] = None
    print()
    if all(v is not None for v in sums.values()):
        d_2_1 = sums["2_FR_exclude_RaschetSeb"] - sums["1_FR_as_is"]
        d_3_2 = sums["3_VIS_revenue"] - sums["2_FR_exclude_RaschetSeb"]
        print(f"  Δ(2-1) = {d_2_1:>20,.2f}  (втрата якщо фільтр)")
        print(f"  Δ(3-2) = {d_3_2:>20,.2f}  (різниця VIS vs ФР_filtered)")


if __name__ == "__main__":
    main()
