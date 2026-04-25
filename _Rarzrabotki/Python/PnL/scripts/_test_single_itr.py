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
print(f"Всего строк: {n}")

# Показываем первые 5 строк со всеми полями
print("\n=== Колонки TЗ ===")
cols = tz.Колонки
for i in range(cols.Количество()):
    print(f"  {i}: {cols.Получить(i).Имя}")

print("\n=== Первые 3 строки ===")
for i in range(min(3, n)):
    r = tz.Получить(i)
    print(f"\nСтрока {i}:")
    for c in range(cols.Количество()):
        cn = cols.Получить(c).Имя
        val = getattr(r, cn, None)
        try:
            sv = str(val) if val is not None else "None"
        except Exception:
            sv = "(err)"
        print(f"  {cn[:25]:25} = {sv[:80]}")
print("\n=== Астарта ===")
found_astarta = 0
total_pl = 0
total_erp = 0
for i in range(n):
    r = tz.Получить(i)
    # ПодразделениеСтрока — это настоящая строка (str работает)
    подр_str = str(r.ПодразделениеСтрока or "")
    if "Астарта" not in подр_str:
        continue
    found_astarta += 1
    # Для ссылочных полей — используем conn.XMLСтрока или .Наименование
    pl_name = ""
    if r.СтатьяPL:
        try:
            pl_name = str(r.СтатьяPL.Наименование)
        except Exception:
            pl_name = "(err)"
    else:
        pl_name = "(NULL)"
    pl_sum = float(r.СуммаPL or 0)
    erp = float(r.СуммаЕРП or 0)
    total_pl += pl_sum
    total_erp += erp
    mark = "★" if any(k in pl_name for k in ("ЗП", "Зарплат", "Начислен", "Удержан")) else " "
    print(f"  [{i:3}] {pl_name[:45]:45} PL={pl_sum:>10.0f}  ЕРП={erp:>10.0f}  Δ={pl_sum-erp:>10.0f} {mark}")

print(f"\nИТОГО Астарта: строк {found_astarta}  PL={total_pl:.0f}  ЕРП={total_erp:.0f}  Δ={total_pl-total_erp:.0f}")
