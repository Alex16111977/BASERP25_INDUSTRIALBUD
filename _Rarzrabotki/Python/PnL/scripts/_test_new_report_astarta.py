"""Вызывает Отчеты.А_ОтчетPL.Создать().ПолучитьОбъединенныеДанные (теперь Экспорт)
и выводит строки по Астарта. Тищенки декабрь 2025.
"""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

conn = connect_erp()
rep = conn.Отчеты.А_ОтчетPL.Создать()

tz = rep.ПолучитьОбъединенныеДанные(
    datetime.datetime(2025, 12, 1),
    datetime.datetime(2025, 12, 31, 23, 59, 59)
)
n = tz.Количество()
print(f"Всего строк: {n}\n")

# Фильтруем по Астарте
print(f"{'Группа':45} {'СтатьяPL':45} {'СуммаPL':>12} {'СуммаЕРП':>12} {'Разница':>12}")
print("-" * 130)
totals = {"pl": 0, "erp": 0}
no_pl_rows = []
for i in range(n):
    r = tz.Получить(i)
    подр = str(r.ПодразделениеСтрока or r.Подразделение or "")
    if "Астарта" not in подр:
        continue
    pl_статья = str(r.СтатьяPL) if r.СтатьяPL else "(NULL)"
    группа = str(r.Группа) if r.Группа else "(NULL)"
    pl_сум = float(r.СуммаPL or 0)
    ерп = float(r.СуммаЕРП or 0)
    разн = float(r.Разница or 0)
    totals["pl"] += pl_сум
    totals["erp"] += ерп

    if pl_статья == "(NULL)" and ерп != 0:
        no_pl_rows.append((str(r.ДДС), ерп))
        continue

    mark = "★" if any(k in pl_статья for k in ("Зарплат", "ЗП ", "Удержан", "Начислен")) else " "
    print(f"{группа[:45]:45} {pl_статья[:45]:45} {pl_сум:>12.0f} {ерп:>12.0f} {разн:>12.0f} {mark}")

print("-" * 130)
print(f"{'ИТОГО Астарта. Тищенки':90}   {totals['pl']:>12.0f} {totals['erp']:>12.0f} {totals['pl']-totals['erp']:>12.0f}")

if no_pl_rows:
    print(f"\n⚠ Строки БЕЗ PL-статьи (сумма ЕРП без расшифровки):")
    for dds, summa in no_pl_rows:
        print(f"    ДДС={dds[:60]:60}  СуммаЕРП={summa:>12.0f}")
else:
    print("\n✅ Нет строк без PL-статьи (все ДДС привязаны)")
