"""
Точна перевірка CoGS Глобино-2 за лютий 2026.

Еталон зі стандартного звіту "Доходы и расходы предприятия" (упр.учёт с НДС):
  Глобино-2 / Себестоимость продаж = 13 687 175.04 ₴

Джерело: РегистрНакопления.ВыручкаИСебестоимостьПродаж.Стоимость (с НДС).
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
ETALON_COGS = 13_687_175.04

QUERY = """
ВЫБРАТЬ
    Выр.Подразделение.Наименование КАК ПодразНаим,
    СУММА(Выр.Стоимость) КАК Сумма
ИЗ РегистрНакопления.ВыручкаИСебестоимостьПродаж КАК Выр
ГДЕ Выр.Период МЕЖДУ &НачалоПериода И &КонецПериода
    И Выр.Активность
    И Выр.Стоимость <> 0
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
    print(f"CoGS лютий 2026 (с НДС): {total:>20,.2f} ₴ ({len(rows)} підрозд.)")
    print()
    print(f"{'Підрозділ':<40} {'CoGS':>20}")
    for name, sum_ in rows:
        marker = " <-- ЕТАЛОН" if GLOBINO_2_NAME == name else ""
        print(f"{name[:40]:<40} {sum_:>20,.2f}{marker}")

    glob = next((s for n, s in rows if GLOBINO_2_NAME == n), None)
    print()
    if glob is None:
        print(f"FAIL: не знайдено {GLOBINO_2_NAME}")
        sys.exit(1)
    delta = abs(glob - ETALON_COGS)
    print(f"Глобино-2: {glob:,.2f} ₴ vs еталон {ETALON_COGS:,.2f} ₴ → Δ {delta:,.2f}")
    if delta > 0.01:
        print(f"FAIL: розбіжність {delta:,.2f} ₴ перевищує 0.01")
        sys.exit(1)
    print("COGS SUM OK (до копійки)")


if __name__ == "__main__":
    main()
