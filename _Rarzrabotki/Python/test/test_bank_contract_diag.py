# -*- coding: utf-8 -*-
"""Діагностика: чому обробка не оновлює договори.
Перевіряємо типи реквізитів та конкретний договір 'Договір 18/СК від 13.01.2025'
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

# === 1. Перевірка типів реквізитів в договорі ===
print("=" * 60)
print("1. ТИПИ РЕКВІЗИТІВ В ДОВІДНИКУ ДоговорыКонтрагентов")
print("=" * 60)
meta_contract = conn.Metadata.Catalogs.ДоговорыКонтрагентов
target_attrs = [
    "Подразделение", "А_Подразделение", "А_ПодразделениеБанк", "А_ПодразделениеКазна",
    "НаправлениеДеятельности", "А_НаправлениеДеятельности", "А_НаправлениеДеятельностиБанк", "А_НаправлениеДеятельностиКазна"
]
contract_attrs = {}
for i in range(meta_contract.Attributes.Count()):
    attr = meta_contract.Attributes.Get(i)
    if attr.Name in target_attrs:
        type_str = str(attr.Type)
        contract_attrs[attr.Name] = type_str
        print(f"  {attr.Name}: {type_str}")

# === 2. Перевірка типів в документах банку ===
print("\n" + "=" * 60)
print("2. ТИПИ РЕКВІЗИТІВ В СписаниеБезналичныхДенежныхСредств")
print("=" * 60)
meta_spis = conn.Metadata.Documents.СписаниеБезналичныхДенежныхСредств
doc_attrs = {}
for name in ["А_Подразделение", "А_НаправлениеДеятельности", "Договор", "Подразделение", "НаправлениеДеятельности"]:
    for i in range(meta_spis.Attributes.Count()):
        attr = meta_spis.Attributes.Get(i)
        if attr.Name == name:
            type_str = str(attr.Type)
            doc_attrs[name] = type_str
            print(f"  {attr.Name}: {type_str}")

# === 3. Знайти конкретний договір ===
print("\n" + "=" * 60)
print("3. ПОШУК ДОГОВОРУ 'Договір 18/СК від 13.01.2025'")
print("=" * 60)

q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 5
    Спр.Ссылка КАК Ссылка,
    Спр.Наименование КАК Наименование,
    Спр.Подразделение КАК Подразделение,
    Спр.А_ПодразделениеБанк КАК ПодразделениеБанк,
    Спр.А_НаправлениеДеятельностиБанк КАК НаправлениеБанк
ИЗ
    Справочник.ДоговорыКонтрагентов КАК Спр
ГДЕ
    Спр.Наименование ПОДОБНО &Наим"""
q.SetParameter("Наим", "%18/СК%13.01.2025%")

res = q.Execute()
sel = res.Select()
contract_ref = None
while sel.Next():
    print(f"  Найдено: {sel.Наименование}")
    print(f"    Подразделение: {sel.Подразделение}")
    print(f"    А_ПодразделениеБанк: {sel.ПодразделениеБанк}")
    print(f"    А_НаправлениеДеятельностиБанк: {sel.НаправлениеБанк}")
    print(f"    Подразделение filled: {conn.ValueIsFilled(sel.Подразделение)}")
    print(f"    ПодразделениеБанк filled: {conn.ValueIsFilled(sel.ПодразделениеБанк)}")
    print(f"    НаправлениеБанк filled: {conn.ValueIsFilled(sel.НаправлениеБанк)}")
    contract_ref = sel.Ссылка

# === 4. Знайти документи Списание для цього договору ===
print("\n" + "=" * 60)
print("4. ДОКУМЕНТИ СписаниеБезналичныхДС ДЛЯ ЦЬОГО ДОГОВОРУ")
print("=" * 60)

if contract_ref:
    q2 = conn.NewObject("Запрос")
    q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 10
        Док.Ссылка КАК Ссылка,
        Док.Номер КАК Номер,
        Док.Дата КАК ДатаДок,
        Док.Договор КАК Договор,
        Док.А_Подразделение КАК А_Подразделение,
        Док.А_НаправлениеДеятельности КАК А_НаправлениеДеятельности,
        Док.Подразделение КАК Подразделение,
        Док.НаправлениеДеятельности КАК НаправлениеДеятельности
    ИЗ
        Документ.СписаниеБезналичныхДенежныхСредств КАК Док
    ГДЕ
        Док.Проведен
        И Док.Договор = &Договор
        И Док.Дата >= &Дата1
    УПОРЯДОЧИТЬ ПО
        Док.Дата"""
    q2.SetParameter("Договор", contract_ref)
    q2.SetParameter("Дата1", datetime.datetime(2025, 12, 1))

    res2 = q2.Execute()
    sel2 = res2.Select()
    while sel2.Next():
        print(f"  Док #{sel2.Номер} от {sel2.ДатаДок}")
        print(f"    А_Подразделение: {sel2.А_Подразделение} (filled={conn.ValueIsFilled(sel2.А_Подразделение)})")
        print(f"    А_НаправлениеДеятельности: {sel2.А_НаправлениеДеятельности} (filled={conn.ValueIsFilled(sel2.А_НаправлениеДеятельности)})")
        print(f"    Подразделение (стандарт): {sel2.Подразделение} (filled={conn.ValueIsFilled(sel2.Подразделение)})")
        print(f"    НаправлениеДеятельности (стандарт): {sel2.НаправлениеДеятельности} (filled={conn.ValueIsFilled(sel2.НаправлениеДеятельности)})")

# === 5. Перевірка запиту обробки - що повертає для цього договору ===
print("\n" + "=" * 60)
print("5. ЩО ПОВЕРТАЄ ЗАПИТ ОБРОБКИ ДЛЯ ЦЬОГО ДОГОВОРУ")
print("=" * 60)

if contract_ref:
    q3 = conn.NewObject("Запрос")
    q3.Text = """ВЫБРАТЬ
        Док.Договор КАК Договор,
        МАКСИМУМ(Док.А_Подразделение) КАК Подразделение,
        МАКСИМУМ(Док.А_НаправлениеДеятельности) КАК НаправлениеДеятельности
    ИЗ
        Документ.СписаниеБезналичныхДенежныхСредств КАК Док
    ГДЕ
        Док.Проведен
        И Док.Договор = &Договор
        И Док.Дата МЕЖДУ &Дата1 И &Дата2
    СГРУППИРОВАТЬ ПО
        Док.Договор"""
    q3.SetParameter("Договор", contract_ref)
    q3.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
    q3.SetParameter("Дата2", datetime.datetime(2026, 12, 31, 23, 59, 59))

    res3 = q3.Execute()
    tbl = res3.Unload()
    print(f"  Rows: {tbl.Count()}")
    for i in range(tbl.Count()):
        row = tbl.Get(i)
        print(f"  Row {i}: Договор={row.Договор}")
        print(f"    Подразделение={row.Подразделение} (filled={conn.ValueIsFilled(row.Подразделение)})")
        print(f"    НаправлениеДеятельности={row.НаправлениеДеятельности} (filled={conn.ValueIsFilled(row.НаправлениеДеятельности)})")

# === 6. Перевірка сумісності типів ===
print("\n" + "=" * 60)
print("6. ПЕРЕВІРКА СУМІСНОСТІ ТИПІВ")
print("=" * 60)

# А_Подразделение (в документі) -> А_ПодразделениеБанк (в договорі)
if "А_Подразделение" in doc_attrs and "А_ПодразделениеБанк" in contract_attrs:
    doc_type = doc_attrs["А_Подразделение"]
    contract_type = contract_attrs["А_ПодразделениеБанк"]
    match = "OK" if doc_type == contract_type else "MISMATCH!"
    print(f"  Док.А_Подразделение: {doc_type}")
    print(f"  Дог.А_ПодразделениеБанк: {contract_type}")
    print(f"  => {match}")

if "А_НаправлениеДеятельности" in doc_attrs and "А_НаправлениеДеятельностиБанк" in contract_attrs:
    doc_type = doc_attrs["А_НаправлениеДеятельности"]
    contract_type = contract_attrs["А_НаправлениеДеятельностиБанк"]
    match = "OK" if doc_type == contract_type else "MISMATCH!"
    print(f"  Док.А_НаправлениеДеятельности: {doc_type}")
    print(f"  Дог.А_НаправлениеДеятельностиБанк: {contract_type}")
    print(f"  => {match}")

print("\n=== DONE ===")
