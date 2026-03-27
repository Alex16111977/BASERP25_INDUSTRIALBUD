# -*- coding: utf-8 -*-
"""Тест: правильний спосіб встановити RecordType = Расход"""
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Отримати значення Расход через запит
q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Знач"
r = q.Execute().Select()
r.Next()
RASHOD = r.Знач
print(f"РАСХОД XML: {conn.XMLString(RASHOD)}")
print(f"РАСХОД str: {S(RASHOD)}")
print(f"РАСХОД type: {type(RASHOD)}")

# Спроба встановити
nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
rec = nabor.Add()
print(f"\nДо: {conn.XMLString(rec.RecordType)}")

rec.RecordType = RASHOD
print(f"Після RASHOD: {conn.XMLString(rec.RecordType)}")

# Альтернатива: ВидДвиженияНакопления через глобальний контекст
try:
    rashod2 = conn.Eval("ВидДвиженияНакопления.Расход")
    print(f"\nEval: {conn.XMLString(rashod2)}")
    rec2 = nabor.Add()
    rec2.RecordType = rashod2
    print(f"Після Eval: {conn.XMLString(rec2.RecordType)}")
except Exception as e:
    print(f"Eval error: {e}")

# Ще альтернатива
try:
    rashod3 = conn.AccumulationRecordType.Expense
    print(f"\nAccumulationRecordType: {rashod3}")
except Exception as e:
    print(f"AccumulationRecordType error: {e}")

print(f"\nРезультат:")
for i in range(nabor.Count()):
    r = nabor.Get(i)
    print(f"  #{i+1} RT={conn.XMLString(r.RecordType)}")
