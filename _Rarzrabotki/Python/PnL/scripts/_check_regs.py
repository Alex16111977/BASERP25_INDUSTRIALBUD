"""Check structure of 3 candidate registers."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

c = connect_erp()
md = c.Метаданные
for name in ["ВыручкаИСебестоимостьПродаж", "ПрочиеРасходы",
             "ДвиженияНоменклатураДоходыРасходы", "ДвиженияКонтрагентДоходыРасходы",
             "ДвиженияДенежныеСредстваДоходыРасходы"]:
    r = md.РегистрыНакопления.Найти(name)
    if not r:
        print(f"\n{name}: NOT FOUND")
        continue
    print(f"\n=== {name} ({r.ВидРегистра}) ===")
    print("Изм:", ", ".join(str(x.Имя) for x in r.Измерения))
    print("Рес:", ", ".join(str(x.Имя) for x in r.Ресурсы))
