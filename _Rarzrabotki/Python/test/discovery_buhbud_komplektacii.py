# -*- coding: utf-8 -*-
"""Discovery метаданных BuhBud для обработки 'Создания комплектаций для бух учета'."""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
MD = buh.Метаданные

def dump_obj(coll, name, label):
    try:
        o = getattr(coll, name)
    except Exception:
        o = coll.Найти(name)
    if o is None:
        print(f"  [{label}] '{name}' НЕ НАЙДЕН")
        return None
    print(f"\n=== {label}: {o.Имя} (синоним: {o.Синоним}) ===")
    try:
        print("  Реквизиты:")
        for r in o.Реквизиты:
            print(f"    {r.Имя}: {r.Тип}")
    except Exception as e:
        print("   реквизиты:", e)
    try:
        for tc in o.ТабличныеЧасти:
            print(f"  ТЧ '{tc.Имя}':")
            for c in tc.Реквизиты:
                print(f"      {c.Имя}: {c.Тип}")
    except Exception as e:
        print("   ТЧ:", e)
    return o

# 1) Документ комплектации
dump_obj(MD.Документы, "КомплектацияНоменклатуры", "Документ")

# 2) Справочник норм (спецификация)
dump_obj(MD.Справочники, "СтруктураСебестоимости", "Справочник")

# 3) Склады
try:
    print("\n=== Справочник.Склады:", MD.Справочники.Склады.Имя, "===")
except Exception as e:
    print("Склады:", e)

# 4) Регистры накопления — кандидаты на остатки товаров
print("\n=== РегистрыНакопления (Товар/Склад/Остат/Запас) ===")
for rn in MD.РегистрыНакопления:
    im = rn.Имя
    if any(k in im for k in ("Товар","Склад","Остат","Запас","Партии","Номенклатур")):
        try:
            izм = ",".join([d.Имя for d in rn.Измерения])
            res = ",".join([r.Имя for r in rn.Ресурсы])
            print(f"  {im} [{rn.ВидРегистра}] изм:({izм}) рес:({res})")
        except Exception as e:
            print(f"  {im}: {e}")

# 5) Хозрасчетный — есть ли субконто Склад/Номенклатура
print("\n=== РегистрыБухгалтерии ===")
for rb in MD.РегистрыБухгалтерии:
    print(f"  {rb.Имя}")
