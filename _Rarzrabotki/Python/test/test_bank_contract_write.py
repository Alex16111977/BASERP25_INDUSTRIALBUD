# -*- coding: utf-8 -*-
"""Тест запису: оновлюємо один конкретний договір 'Договір 18/СК від 13.01.2025'
через COM, аналогічно до обробки.
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

# 1. Виправлений запит — для одного договору
print("=" * 60)
print("1. ВИПРАВЛЕНИЙ ЗАПИТ ДЛЯ ДОГОВОРУ 18/СК")
print("=" * 60)

# Знайти договір
q_find = conn.NewObject("Запрос")
q_find.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Спр.Ссылка ИЗ Справочник.ДоговорыКонтрагентов КАК Спр
    ГДЕ Спр.Наименование ПОДОБНО &Наим"""
q_find.SetParameter("Наим", "%18/СК%13.01.2025%")
res_find = q_find.Execute()
sel_find = res_find.Select()
sel_find.Next()
contract_ref = sel_find.Ссылка
print(f"  Договір знайдено: {conn.String(contract_ref)}")

# Стан ДО
print(f"\n  СТАН ДО:")
print(f"    Подразделение: {conn.String(contract_ref.Подразделение)} (filled={conn.ValueIsFilled(contract_ref.Подразделение)})")
print(f"    А_ПодразделениеБанк: {conn.String(contract_ref.А_ПодразделениеБанк)} (filled={conn.ValueIsFilled(contract_ref.А_ПодразделениеБанк)})")
print(f"    А_НаправлениеДеятельностиБанк: {conn.String(contract_ref.А_НаправлениеДеятельностиБанк)} (filled={conn.ValueIsFilled(contract_ref.А_НаправлениеДеятельностиБанк)})")

# 2. Запит обробки для цього договору
print("\n" + "=" * 60)
print("2. РЕЗУЛЬТАТ ЗАПИТУ ОБРОБКИ ДЛЯ ЦЬОГО ДОГОВОРУ")
print("=" * 60)

q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ
    РП.А_Договор КАК Договор,
    МАКСИМУМ(Док.А_Подразделение) КАК Подразделение,
    МАКСИМУМ(Док.А_НаправлениеДеятельности) КАК НаправлениеДеятельности
ИЗ
    Документ.СписаниеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК РП
    ЛЕВОЕ СОЕДИНЕНИЕ Документ.СписаниеБезналичныхДенежныхСредств КАК Док
    ПО РП.Ссылка = Док.Ссылка
ГДЕ
    Док.Проведен
    И РП.А_Договор = &Договор
    И (Док.А_Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ИЛИ Док.А_НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка))
    И Док.Дата МЕЖДУ &Дата1 И &Дата2
СГРУППИРОВАТЬ ПО
    РП.А_Договор"""
q.SetParameter("Договор", contract_ref)
q.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
q.SetParameter("Дата2", datetime.datetime(2026, 12, 31, 23, 59, 59))

res = q.Execute()
tbl = res.Unload()
print(f"  Rows: {tbl.Count()}")

if tbl.Count() > 0:
    row = tbl.Get(0)
    podrazd = row.Подразделение
    napr = row.НаправлениеДеятельности
    print(f"  Договор: {conn.String(row.Договор)}")
    print(f"  Подразделение: {conn.String(podrazd)} (filled={conn.ValueIsFilled(podrazd)})")
    print(f"  НаправлениеДеятельности: {conn.String(napr)} (filled={conn.ValueIsFilled(napr)})")

    # 3. Запис
    print("\n" + "=" * 60)
    print("3. ЗАПИС В ДОГОВІР")
    print("=" * 60)

    obj = contract_ref.GetObject()

    if conn.ValueIsFilled(podrazd):
        obj.А_ПодразделениеБанк = podrazd
        print(f"  Встановлено А_ПодразделениеБанк = {conn.String(podrazd)}")

    if conn.ValueIsFilled(napr):
        obj.А_НаправлениеДеятельностиБанк = napr
        print(f"  Встановлено А_НаправлениеДеятельностиБанк = {conn.String(napr)}")

    obj.DataExchange.Load = True
    obj.Write()
    print("  Записано успішно!")

    # 4. Перевірка ПІСЛЯ
    print("\n" + "=" * 60)
    print("4. СТАН ПІСЛЯ ЗАПИСУ")
    print("=" * 60)

    # Перечитати
    q_check = conn.NewObject("Запрос")
    q_check.Text = """ВЫБРАТЬ ПЕРВЫЕ 1
        Спр.Подразделение КАК Подразделение,
        Спр.А_ПодразделениеБанк КАК ПодразделениеБанк,
        Спр.А_НаправлениеДеятельностиБанк КАК НаправлениеБанк
    ИЗ
        Справочник.ДоговорыКонтрагентов КАК Спр
    ГДЕ
        Спр.Ссылка = &Ссылка"""
    q_check.SetParameter("Ссылка", contract_ref)
    res_check = q_check.Execute()
    sel_check = res_check.Select()
    sel_check.Next()
    print(f"  Подразделение: {conn.String(sel_check.Подразделение)}")
    print(f"  А_ПодразделениеБанк: {conn.String(sel_check.ПодразделениеБанк)} (filled={conn.ValueIsFilled(sel_check.ПодразделениеБанк)})")
    print(f"  А_НаправлениеДеятельностиБанк: {conn.String(sel_check.НаправлениеБанк)} (filled={conn.ValueIsFilled(sel_check.НаправлениеБанк)})")

else:
    print("  EMPTY - no data for this contract!")

print("\n=== DONE ===")
