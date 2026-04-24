"""
Перевіряє: для (Підрозділ) сума з ВыручкаИСебестоимостьПродаж за лютий 2026
співпадає з сумою «Выручка от продаж» з ФинансовыеРезультаты.

Архітектура (узгоджена 2026-04-25):
- Джерело Виручки = ВыручкаИСебестоимостьПродаж (51M, без РасчетСеб)
- СтатьяДоходов hardcode = "Выручка от продаж" (UUID 77981290-6ea2-11f0-a2de-bccbaebe2890)
- Розклад по PL працює через JOIN А_Статьи_PL.СтатьяДоходов = "Выручка от продаж"

Тест: ВиР по підрозділах ≈ сума «Выручка от продаж» з ФР по тих самих підрозділах.
Допускається невелика різниця (РасчетСеб коригує по підрозділу).
"""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pywintypes  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

PERIOD_START = pywintypes.Time(datetime.datetime(2026, 2, 1, 0, 0, 0))
PERIOD_END = pywintypes.Time(datetime.datetime(2026, 2, 28, 23, 59, 59))

QUERY_VIS = """
ВЫБРАТЬ
    Выр.Подразделение.Наименование КАК ПодразНаим,
    СУММА(Выр.СуммаВыручкиБезНДСОборот) КАК Сумма
ИЗ РегистрНакопления.ВыручкаИСебестоимостьПродаж.Обороты(&НачалоПериода, &КонецПериода, , ) КАК Выр
СГРУППИРОВАТЬ ПО Выр.Подразделение.Наименование
"""

QUERY_FR_SALES_ONLY = """
ВЫБРАТЬ
    Ф.Подразделение.Наименование КАК ПодразНаим,
    СУММА(Ф.ДоходыОборот) КАК Сумма
ИЗ РегистрНакопления.ФинансовыеРезультаты.Обороты(&НачалоПериода, &КонецПериода, , ) КАК Ф
ГДЕ Ф.СтатьяДоходов.Наименование = "Выручка от продаж"
СГРУППИРОВАТЬ ПО Ф.Подразделение.Наименование
"""


def fetch(erp, text):
    q = erp.NewObject("Запрос")
    q.Текст = text
    q.УстановитьПараметр("НачалоПериода", PERIOD_START)
    q.УстановитьПараметр("КонецПериода", PERIOD_END)
    sel = q.Выполнить().Выбрать()
    out = {}
    while sel.Следующий():
        out[str(sel.ПодразНаим or "")] = float(sel.Сумма or 0)
    return out


def main():
    erp = connect_erp()
    vis = fetch(erp, QUERY_VIS)
    fr = fetch(erp, QUERY_FR_SALES_ONLY)

    sum_vis = sum(vis.values())
    sum_fr = sum(fr.values())
    print(f"VIS  total: {sum_vis:>20,.2f} ₴ ({len(vis)} підрозд.)")
    print(f"FR   total: {sum_fr:>20,.2f} ₴ ({len(fr)} підрозд.)")
    print(f"Δ FR-VIS  : {sum_fr - sum_vis:>20,.2f} ₴  (коригування РасчетСеб)")
    print()

    keys = sorted(set(vis) | set(fr))
    print(f"{'Подразделение':<50} {'VIS':>15} {'FR':>15} {'delta':>15}")
    for k in keys:
        a, b = vis.get(k, 0.0), fr.get(k, 0.0)
        d = b - a
        marker = "" if abs(d) < 100 else (" *" if abs(d) < 100000 else " **")
        print(f"{k[:50]:<50} {a:>15,.2f} {b:>15,.2f} {d:>15,.2f}{marker}")


if __name__ == "__main__":
    main()
