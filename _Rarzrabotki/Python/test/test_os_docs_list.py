# -*- coding: utf-8 -*-
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

print("=== Документи з ОС/ВНА/Внеоборот ===")
for i in range(conn.Metadata.Documents.Count()):
    doc = conn.Metadata.Documents.Get(i)
    name = str(doc.Name)
    try:
        syn = str(doc.Synonym)
    except:
        syn = ""
    if any(k in name for k in ["ОС", "ВНА", "Внеоборот", "Эксплуатац", "Амортизац"]):
        marker = " [НЕ ВИКОРИСТОВУЄТЬСЯ]" if "не используется" in syn.lower() else ""
        print(f"  {name} | {syn}{marker}")

print("\n=== Регістри СтоимостьОС — хто пише ===")
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    ТИПЗНАЧЕНИЯ(Д.Регистратор) КАК ТипРегистратора
ИЗ РегистрНакопления.СтоимостьОС КАК Д"""
try:
    r = q.Execute().Select()
    while r.Next():
        print(f"  {r.ТипРегистратора}")
except Exception as e:
    print(f"  Помилка: {str(e)[:100]}")

# Альтернатива — подивитися хто реально пише в СтоимостьОС
print("\n=== Останні 5 записів СтоимостьОС ===")
q2 = conn.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 5
    Д.Период, Д.Регистратор, Д.ОсновноеСредство, Д.Стоимость
ИЗ РегистрНакопления.СтоимостьОС КАК Д
УПОРЯДОЧИТЬ ПО Д.Период УБЫВ"""
try:
    r2 = q2.Execute().Select()
    cnt = 0
    while r2.Next():
        cnt += 1
        print(f"  {r2.Период} | {conn.String(r2.Регистратор)} | {conn.String(r2.ОсновноеСредство)} | {r2.Стоимость}")
    if cnt == 0:
        print("  ПУСТО — регістр СтоимостьОС не має жодних записів!")
except Exception as e:
    print(f"  Помилка: {str(e)[:100]}")

print("\n=== DONE ===")
