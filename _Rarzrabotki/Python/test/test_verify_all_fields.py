# -*- coding: utf-8 -*-
"""Перевірка: чи всі поля з нашого запиту існують в ВтПрочиеРасходыНаЗатраты.
Імітуємо запит який виконує ВыполнитьЗапросПоОписаниюПолей.
"""
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Наш запит бере поля з ВтПрочиеРасходыНаЗатраты (= регістр ПрочиеРасходы після маппінгу)
# Перевіримо кожне поле окремо

# Список полів які ми використовуємо (з КопируемыеПоля + модифіковані)
fields_to_check = [
    "Период",
    "Организация",
    "Подразделение",
    "АналитикаРасходов",
    "ХозяйственнаяОперация",
    "АналитикаУчетаНоменклатуры",
    "ГруппаПродукции",
    "ДокументДвижения",
    "НалоговоеНазначение",
    "Сумма",
    "СтатьяРасходов",
    "НаправлениеДеятельности",
    # Службові
    "СлужебноеВидДвиженияПриход",
]

# Поля регістру ПрочиеРасходы
meta = conn.Metadata.AccumulationRegisters.ПрочиеРасходы
reg_fields = set()
# Стандартні
reg_fields.add("Период")
reg_fields.add("Регистратор")
reg_fields.add("ВидДвижения")
reg_fields.add("Активность")
# Виміри
for i in range(meta.Dimensions.Count()):
    reg_fields.add(meta.Dimensions.Get(i).Name)
# Ресурси
for i in range(meta.Resources.Count()):
    reg_fields.add(meta.Resources.Get(i).Name)
# Реквізити
for i in range(meta.Attributes.Count()):
    reg_fields.add(meta.Attributes.Get(i).Name)

print("=== Поля регістру ПрочиеРасходы ===")
for f in sorted(reg_fields):
    print(f"  {f}")

print(f"\nВсього: {len(reg_fields)}")

# Перевірка наших полів
print("\n=== Перевірка наших полів ===")
missing = []
for field in fields_to_check:
    if field in reg_fields:
        print(f"  ✅ {field}")
    else:
        print(f"  ❌ {field} — НЕМАЄ!")
        missing.append(field)

# Поля які використовуємо в маппінгу але їх може не бути
extra_fields = ["Знак", "КорНаправлениеДеятельности", "СтатьяРасходовСписания", "ВариантРаспределенияРасходовРегл", "ЭтоВыпускПродукции", "ИсходнаяСтоимость", "СуммаСНДС"]
print("\n=== Поля з ПРИХОД які НЕ мають бути в нашому РАСХОД ===")
for field in extra_fields:
    if field in reg_fields:
        print(f"  ⚠️ {field} — є в регістрі (ок)")
    else:
        print(f"  ❌ {field} — НЕМАЄ в регістрі (не можна використовувати!)")

if missing:
    print(f"\n⚠️ ПРОБЛЕМА: поля {missing} не знайдені!")
else:
    print("\n✅ Всі наші поля існують!")

# Додатково: перевірити ГруппаПродукции
print("\n=== Перевірка ГруппаПродукции ===")
if "ГруппаПродукции" in reg_fields:
    print("  ✅ ГруппаПродукции є в регістрі")
else:
    print("  ❌ ГруппаПродукции НЕМАЄ — це реквізит іншого регістру!")
    # Перевірити чи це не від СебестоимостьТоваров
    meta2 = conn.Metadata.AccumulationRegisters.СебестоимостьТоваров
    for i in range(meta2.Attributes.Count()):
        if meta2.Attributes.Get(i).Name == "ГруппаПродукции":
            print("  ⚠️ ГруппаПродукции є в СебестоимостьТоваров!")

print("\n=== DONE ===")
