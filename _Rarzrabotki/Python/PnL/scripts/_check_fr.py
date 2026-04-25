"""Check РегистрНакопления.ФинансовыеРезультаты structure."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

c = connect_erp()
md = c.Метаданные
r = md.РегистрыНакопления.Найти("ФинансовыеРезультаты")
print("=== Измерения ===")
for x in r.Измерения:
    print(f"  - {x.Имя}")
print("=== Ресурсы ===")
for x in r.Ресурсы:
    print(f"  - {x.Имя}")
print("=== Реквизиты ===")
for x in r.Реквизиты:
    print(f"  - {x.Имя}")
