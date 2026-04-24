"""
Інформаційний тест: показує суму СтоимостьУпрБезНДС з ВыручкаИСебестоимостьПродаж
за лютий 2026 по підрозділах. Фінансист зіставляє з еталоном звіту "Доходы и расходы предприятия".
Без assertions — еталон у проєкті ще не підтверджений.
"""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pywintypes  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

PERIOD_START = pywintypes.Time(datetime.datetime(2026, 2, 1, 0, 0, 0))
PERIOD_END = pywintypes.Time(datetime.datetime(2026, 2, 28, 23, 59, 59))

GLOBINO_2_NAME = "Глобино-2"

QUERY = """
ВЫБРАТЬ
    Выр.Подразделение.Наименование КАК ПодразНаим,
    СУММА(Выр.СтоимостьУпрБезНДСОборот) КАК Сумма
ИЗ РегистрНакопления.ВыручкаИСебестоимостьПродаж.Обороты(&НачалоПериода, &КонецПериода, , ) КАК Выр
СГРУППИРОВАТЬ ПО Выр.Подразделение.Наименование
УПОРЯДОЧИТЬ ПО Сумма УБЫВ
"""


def main():
    erp = connect_erp()
    q = erp.NewObject("Запрос")
    q.Текст = QUERY
    q.УстановитьПараметр("НачалоПериода", PERIOD_START)
    q.УстановитьПараметр("КонецПериода", PERIOD_END)
    sel = q.Выполнить().Выбрать()
    rows = []
    while sel.Следующий():
        rows.append((str(sel.ПодразНаим or ""), float(sel.Сумма or 0)))

    total = sum(s for _, s in rows)
    print(f"Total CoGS February 2026: {total:>20,.2f} ₴ ({len(rows)} підрозд.)")
    print()
    print(f"{'Подразделение':<40} {'CoGS':>20}")
    for name, sum_ in rows:
        marker = " <-- ЕТАЛОН" if GLOBINO_2_NAME in name else ""
        print(f"{name[:40]:<40} {sum_:>20,.2f}{marker}")

    glob = next((s for n, s in rows if GLOBINO_2_NAME in n), None)
    print()
    print(f"Глобино-2 CoGS (ВиР.СтоимостьУпрБезНДС): {glob or 0:,.2f} ₴")
    print("INFO: фінансист зіставить з еталоном звіту 'Доходы и расходы'")


if __name__ == "__main__":
    main()
