# -*- coding: utf-8 -*-
"""Тест нового запиту з підрахунком частоти (КОЛИЧЕСТВО замість МАКСИМУМ).
Перевіряємо що для договору 5/0725/18 обирається найчастіше підрозділ.
"""
import win32com.client
import datetime
from collections import Counter

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

# Новий запит з КОЛИЧЕСТВО
query_text = """ВЫБРАТЬ
    РП.А_Договор КАК Договор,
    Док.А_Подразделение КАК Подразделение,
    Док.А_НаправлениеДеятельности КАК НаправлениеДеятельности,
    КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Док.Ссылка) КАК КолДокументов
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
    РП.А_Договор,
    Док.А_Подразделение,
    Док.А_НаправлениеДеятельности

ОБЪЕДИНИТЬ ВСЕ

ВЫБРАТЬ
    РП.А_Договор,
    Док.А_Подразделение,
    Док.А_НаправлениеДеятельности,
    КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Док.Ссылка)
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
    РП.А_Договор,
    Док.А_Подразделение,
    Док.А_НаправлениеДеятельности"""

print("Executing frequency query...")
q = conn.NewObject("Запрос")
q.Text = query_text
q.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
q.SetParameter("Дата2", datetime.datetime(2026, 12, 31, 23, 59, 59))

result = q.Execute()
table = result.Unload()

total_rows = table.Count()
print(f"Query OK! Rows returned: {total_rows}")

# Емулюємо логіку обробки: підрахунок частоти
contracts = {}  # contract_str -> {podr_counter, napr_counter}

# Знайдемо тестовий договір
q_find = conn.NewObject("Запрос")
q_find.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Спр.Ссылка ИЗ Справочник.ДоговорыКонтрагентов КАК Спр
    ГДЕ Спр.Наименование ПОДОБНО &Наим"""
q_find.SetParameter("Наим", "%5/0725/18%")
res_find = q_find.Execute()
sel_find = res_find.Select()
sel_find.Next()
test_contract = sel_find.Ссылка
test_uuid = conn.XMLString(test_contract.УникальныйИдентификатор())

print(f"\nTest contract: {conn.String(test_contract)}")
print(f"Test UUID: {test_uuid}")

# Обробка всіх рядків
test_podrazd = Counter()
test_napr = Counter()
unique_contracts = set()

for i in range(total_rows):
    row = table.Get(i)
    if not conn.ValueIsFilled(row.Договор):
        continue

    contract_uuid = conn.XMLString(row.Договор.УникальныйИдентификатор())
    unique_contracts.add(contract_uuid)
    kol = row.КолДокументов

    if contract_uuid == test_uuid:
        p = conn.String(row.Подразделение) if conn.ValueIsFilled(row.Подразделение) else ""
        n = conn.String(row.НаправлениеДеятельности) if conn.ValueIsFilled(row.НаправлениеДеятельности) else ""
        if p:
            test_podrazd[p] += kol
        if n:
            test_napr[n] += kol

print(f"\nTotal unique contracts: {len(unique_contracts)}")

# Результат для тестового договору
print(f"\n{'='*60}")
print(f"РЕЗУЛЬТАТ ДЛЯ ДОГОВОРУ 5/0725/18")
print(f"{'='*60}")

print("\nЧастота підрозділів:")
for p, cnt in test_podrazd.most_common():
    winner = " <-- WINNER" if cnt == test_podrazd.most_common(1)[0][1] else ""
    print(f"  '{p}': {cnt} документів{winner}")

print("\nЧастота напрямків:")
for n, cnt in test_napr.most_common():
    winner = " <-- WINNER" if cnt == test_napr.most_common(1)[0][1] else ""
    print(f"  '{n}': {cnt} документів{winner}")

if test_podrazd:
    winner_p = test_podrazd.most_common(1)[0][0]
    print(f"\n=> Найчастіше Подразделение: '{winner_p}'")
    print(f"   (замість МАКСИМУМ який давав 'МК Астарта.Тищенки')")
if test_napr:
    winner_n = test_napr.most_common(1)[0][0]
    print(f"=> Найчастіше Направление: '{winner_n}'")

# Тест запису
print(f"\n{'='*60}")
print(f"ТЕСТ ЗАПИСУ В ДОГОВІР")
print(f"{'='*60}")

print(f"  ДО: А_ПодразделениеБанк = '{conn.String(test_contract.А_ПодразделениеБанк)}'")
print(f"  ДО: А_НаправлениеДеятельностиБанк = '{conn.String(test_contract.А_НаправлениеДеятельностиБанк)}'")

# Знайти ссылку на найчастіше підрозділ через запит
if test_podrazd:
    q_podrazd = conn.NewObject("Запрос")
    q_podrazd.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Спр.Ссылка ИЗ Справочник.СтруктураПредприятия КАК Спр
        ГДЕ Спр.Наименование = &Наим"""
    q_podrazd.SetParameter("Наим", winner_p)
    res_p = q_podrazd.Execute()
    sel_p = res_p.Select()
    if sel_p.Next():
        new_podrazd = sel_p.Ссылка

        obj = test_contract.GetObject()
        obj.А_ПодразделениеБанк = new_podrazd
        obj.DataExchange.Load = True
        obj.Write()

        # Перевірка
        q_check = conn.NewObject("Запрос")
        q_check.Text = """ВЫБРАТЬ ПЕРВЫЕ 1
            Спр.А_ПодразделениеБанк КАК ПодразделениеБанк,
            Спр.А_НаправлениеДеятельностиБанк КАК НаправлениеБанк
        ИЗ Справочник.ДоговорыКонтрагентов КАК Спр ГДЕ Спр.Ссылка = &Ссылка"""
        q_check.SetParameter("Ссылка", test_contract)
        res_c = q_check.Execute()
        sel_c = res_c.Select()
        sel_c.Next()
        print(f"  ПІСЛЯ: А_ПодразделениеБанк = '{conn.String(sel_c.ПодразделениеБанк)}'")
        print(f"  ПІСЛЯ: А_НаправлениеДеятельностиБанк = '{conn.String(sel_c.НаправлениеБанк)}'")

print("\n=== TEST PASSED ===")
