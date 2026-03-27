# -*- coding: utf-8 -*-
"""Тест виправленого запиту: беремо договір з РасшифровкаПлатежа.А_Договор,
а підрозділ та напрямок з шапки документа.
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

# Виправлений запит
query_text = """ВЫБРАТЬ
    РП.А_Договор КАК Договор,
    МАКСИМУМ(Док.А_Подразделение) КАК Подразделение,
    МАКСИМУМ(Док.А_НаправлениеДеятельности) КАК НаправлениеДеятельности
ИЗ
    Документ.СписаниеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК РП
    ЛЕВОЕ СОЕДИНЕНИЕ Документ.СписаниеБезналичныхДенежныхСредств КАК Док
    ПО РП.Ссылка = Док.Ссылка
ГДЕ
    Док.Проведен
    И РП.А_Договор <> ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка)
    И (Док.А_Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ИЛИ Док.А_НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка))
    И Док.Дата МЕЖДУ &Дата1 И &Дата2

СГРУППИРОВАТЬ ПО
    РП.А_Договор

ОБЪЕДИНИТЬ ВСЕ

ВЫБРАТЬ
    РП.А_Договор,
    МАКСИМУМ(Док.А_Подразделение),
    МАКСИМУМ(Док.А_НаправлениеДеятельности)
ИЗ
    Документ.ПоступлениеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК РП
    ЛЕВОЕ СОЕДИНЕНИЕ Документ.ПоступлениеБезналичныхДенежныхСредств КАК Док
    ПО РП.Ссылка = Док.Ссылка
ГДЕ
    Док.Проведен
    И РП.А_Договор <> ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка)
    И (Док.А_Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ИЛИ Док.А_НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка))
    И Док.Дата МЕЖДУ &Дата1 И &Дата2

СГРУППИРОВАТЬ ПО
    РП.А_Договор"""

print("Executing FIXED query...")
q = conn.NewObject("Запрос")
q.Text = query_text
q.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
q.SetParameter("Дата2", datetime.datetime(2026, 12, 31, 23, 59, 59))

result = q.Execute()
table = result.Unload()

count = table.Count()
print(f"Query OK! Rows returned: {count}")

# Шукаємо наш договір
print("\n--- Пошук 'Договір 18/СК від 13.01.2025' ---")
found = False
for i in range(count):
    row = table.Get(i)
    if conn.ValueIsFilled(row.Договор):
        name = str(row.Договор)
        if "18" in name and "СК" in name:
            found = True
            print(f"  FOUND! Row {i}: Договор={name}")
            print(f"    Подразделение={row.Подразделение} filled={conn.ValueIsFilled(row.Подразделение)}")
            print(f"    НаправлениеДеятельности={row.НаправлениеДеятельности} filled={conn.ValueIsFilled(row.НаправлениеДеятельности)}")

if not found:
    print("  NOT FOUND - showing first 10 rows:")
    for i in range(min(count, 10)):
        row = table.Get(i)
        print(f"  Row {i}: Договор filled={conn.ValueIsFilled(row.Договор)}, Подр filled={conn.ValueIsFilled(row.Подразделение)}, Напр filled={conn.ValueIsFilled(row.НаправлениеДеятельности)}")

# Статистика
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

# Тест запису — dry run
print("\n--- DRY RUN: перевірка оновлення договору ---")
q_test = conn.NewObject("Запрос")
q_test.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Спр.Ссылка ИЗ Справочник.ДоговорыКонтрагентов КАК Спр
    ГДЕ Спр.Наименование ПОДОБНО &Наим"""
q_test.SetParameter("Наим", "%18/СК%13.01.2025%")
res_test = q_test.Execute()
sel_test = res_test.Select()
if sel_test.Next():
    cref = sel_test.Ссылка
    print(f"  Договір: {cref}")
    print(f"  А_ПодразделениеБанк BEFORE: filled={conn.ValueIsFilled(cref.А_ПодразделениеБанк)}")
    print(f"  А_НаправлениеДеятельностиБанк BEFORE: filled={conn.ValueIsFilled(cref.А_НаправлениеДеятельностиБанк)}")

print("\n=== TEST PASSED ===" if count > 0 else "\n=== WARNING: No data ===")
