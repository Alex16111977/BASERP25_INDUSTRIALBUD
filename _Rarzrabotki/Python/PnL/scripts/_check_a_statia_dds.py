"""Find which catalogs have А_СтатьяДвиженияДенежныхСредств attribute."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

c = connect_erp()
md = c.Метаданные

pvh_to_check = ["СтатьиРасходов", "СтатьиДоходов"]
for name in pvh_to_check:
    o = md.ПланыВидовХарактеристик.Найти(name)
    if not o:
        print(f"PVH {name}: NOT FOUND")
        continue
    print(f"\n=== ПВХ: {name} ===")
    for r in o.Реквизиты:
        nm = str(r.Имя)
        print(f"  - {nm}")
