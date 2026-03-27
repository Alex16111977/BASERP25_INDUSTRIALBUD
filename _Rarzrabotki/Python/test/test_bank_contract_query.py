# -*- coding: utf-8 -*-
"""Тест запиту для обробки 'Оновити в договорах підрозділ з банку'
Перевіряємо що запит працює коректно в BaseERP напряму.
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK")

query_text = """ВЫБРАТЬ
    Док.Договор КАК Договор,
    МАКСИМУМ(Док.А_Подразделение) КАК Подразделение,
    МАКСИМУМ(Док.А_НаправлениеДеятельности) КАК НаправлениеДеятельности
ИЗ
    Документ.СписаниеБезналичныхДенежныхСредств КАК Док
ГДЕ
    Док.Проведен
    И Док.Договор <> ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка)
    И (Док.А_Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ИЛИ Док.А_НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка))
    И Док.Дата МЕЖДУ &Дата1 И &Дата2

СГРУППИРОВАТЬ ПО
    Док.Договор

ОБЪЕДИНИТЬ ВСЕ

ВЫБРАТЬ
    Док.Договор,
    МАКСИМУМ(Док.А_Подразделение),
    МАКСИМУМ(Док.А_НаправлениеДеятельности)
ИЗ
    Документ.ПоступлениеБезналичныхДенежныхСредств КАК Док
ГДЕ
    Док.Проведен
    И Док.Договор <> ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка)
    И (Док.А_Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ИЛИ Док.А_НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка))
    И Док.Дата МЕЖДУ &Дата1 И &Дата2

СГРУППИРОВАТЬ ПО
    Док.Договор"""

print("\nExecuting query...")
q = conn.NewObject("Запрос")
q.Text = query_text
import datetime
q.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
q.SetParameter("Дата2", datetime.datetime(2026, 12, 31, 23, 59, 59))

result = q.Execute()
table = result.Unload()

count = table.Count()
print(f"Query OK! Rows returned: {count}")

# Show first 20 rows
print("\n--- First 20 rows ---")
for i in range(min(count, 20)):
    row = table.Get(i)
    dogovor = str(row.Договор) if conn.ValueIsFilled(row.Договор) else "<empty>"
    podrazd = str(row.Подразделение) if conn.ValueIsFilled(row.Подразделение) else "<empty>"
    napr = str(row.НаправлениеДеятельности) if conn.ValueIsFilled(row.НаправлениеДеятельности) else "<empty>"
    print(f"  {i+1}. Договор: {dogovor}")
    print(f"     Подразделение: {podrazd}")
    print(f"     Направление: {napr}")

# Count how many have both fields filled
both_filled = 0
only_podrazd = 0
only_napr = 0
for i in range(count):
    row = table.Get(i)
    has_p = conn.ValueIsFilled(row.Подразделение)
    has_n = conn.ValueIsFilled(row.НаправлениеДеятельности)
    if has_p and has_n:
        both_filled += 1
    elif has_p:
        only_podrazd += 1
    elif has_n:
        only_napr += 1

print(f"\n--- Statistics ---")
print(f"Total unique contracts: {count}")
print(f"Both filled: {both_filled}")
print(f"Only Подразделение: {only_podrazd}")
print(f"Only Направление: {only_napr}")

# Check that А_ПодразделениеБанк and А_НаправлениеДеятельностиБанк exist in contract
print("\n--- Checking contract attributes ---")
meta = conn.Metadata.Catalogs.ДоговорыКонтрагентов
found_podrazd_bank = False
found_napr_bank = False
for i in range(meta.Attributes.Count()):
    attr = meta.Attributes.Get(i)
    if attr.Name == "А_ПодразделениеБанк":
        found_podrazd_bank = True
        print(f"  Found: А_ПодразделениеБанк (Type: {attr.Type})")
    if attr.Name == "А_НаправлениеДеятельностиБанк":
        found_napr_bank = True
        print(f"  Found: А_НаправлениеДеятельностиБанк (Type: {attr.Type})")

if not found_podrazd_bank:
    print("  WARNING: А_ПодразделениеБанк NOT FOUND in ДоговорыКонтрагентов!")
if not found_napr_bank:
    print("  WARNING: А_НаправлениеДеятельностиБанк NOT FOUND in ДоговорыКонтрагентов!")

print("\n=== TEST PASSED ===" if count > 0 else "\n=== WARNING: No data found ===")
