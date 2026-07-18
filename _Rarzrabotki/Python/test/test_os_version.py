# -*- coding: utf-8 -*-
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

print("=== Functional Options ===")
for opt in ["ИспользоватьВнеоборотныеАктивы", "ИспользоватьВнеоборотныеАктивы2_2",
            "ИспользоватьВнеоборотныеАктивы2_4", "ИспользоватьВнеоборотныеАктивы2_5"]:
    try:
        val = conn.GetFunctionalOption(opt)
        print(f"  {opt} = {val}")
    except:
        print(f"  {opt} - not found")

print("\n=== ВнеоборотныеАктивы.ИспользуетсяУправлениеВНА_2_4 ===")
try:
    q = conn.NewObject("Запрос")
    q.Text = """ВЫБРАТЬ
        РС.ВерсияУчетаНеоборотныхАктивов
    ИЗ
        РегистрСведений.ВерсииУчетаНеоборотныхАктивов КАК РС"""
    r = q.Execute().Select()
    while r.Next():
        print(f"  Версія: {S(r.ВерсияУчетаНеоборотныхАктивов)}")
except Exception as e:
    print(f"  Помилка: {str(e)[:100]}")

# Check all constants with "Необорот" or "Внеоборот"
print("\n=== Constants search ===")
q2 = conn.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Истина ИЗ Константа.ИспользоватьНеоборотныеАктивыВерсия2_5"""
try:
    q2.Execute()
    print("  ИспользоватьНеоборотныеАктивыВерсия2_5 exists")
except:
    print("  ИспользоватьНеоборотныеАктивыВерсия2_5 NOT FOUND")

print("\n=== Doc ПринятиеКУчетуОС (first 3) ===")
q3 = conn.NewObject("Запрос")
q3.Text = """ВЫБРАТЬ ПЕРВЫЕ 3 Д.Ссылка, Д.Номер, Д.Дата, Д.Проведен
    ИЗ Документ.ПринятиеКУчетуОС КАК Д УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"""
r3 = q3.Execute().Select()
while r3.Next():
    ref = r3.Ссылка
    # Count movements
    q4 = conn.NewObject("Запрос")
    q4.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.СтоимостьОС ГДЕ Регистратор = &С"
    q4.SetParameter("С", ref)
    r4 = q4.Execute().Select()
    r4.Next()
    print(f"  #{r3.Номер} от {r3.Дата} Пров={r3.Проведен} СтоимостьОС={r4.К} рухів")

# Check if there's ПринятиеКУчетуОС2_4 document type
print("\n=== Doc types check ===")
for dt in ["ПринятиеКУчетуОС", "ПринятиеКУчетуОС2_4", "ПринятиеКУчетуВНА"]:
    try:
        meta = conn.Metadata.Documents[dt]
        print(f"  {dt} — EXISTS")
    except:
        print(f"  {dt} — NOT FOUND")

print("\n=== DONE ===")
