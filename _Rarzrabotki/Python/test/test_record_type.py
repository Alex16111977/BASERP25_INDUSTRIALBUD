# -*- coding: utf-8 -*-
"""Тест: як правильно встановити RecordType = Расход через COM"""
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
rec = nabor.Add()

print(f"Default RecordType XML: {conn.XMLString(rec.RecordType)}")
print(f"Default RecordType str: {S(rec.RecordType)}")

# Спроба через число
try:
    rec.RecordType = 1
    print(f"After =1: XML={conn.XMLString(rec.RecordType)} str={S(rec.RecordType)}")
except Exception as e:
    print(f"=1 error: {e}")

# Спроба через Перечисление
try:
    q = conn.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК В"
    r = q.Execute().Select()
    r.Next()
    rashod_val = r.В
    print(f"Enum value XML: {conn.XMLString(rashod_val)}")
    print(f"Enum value str: {S(rashod_val)}")

    rec2 = nabor.Add()
    rec2.RecordType = rashod_val
    print(f"After enum: XML={conn.XMLString(rec2.RecordType)} str={S(rec2.RecordType)}")
except Exception as e:
    print(f"Enum error: {e}")

# Спроба через V83.COMConnector
try:
    rec3 = nabor.Add()
    rec3.RecordType = v8.AccumulationRecordType.Expense
    print(f"After v8: XML={conn.XMLString(rec3.RecordType)} str={S(rec3.RecordType)}")
except Exception as e:
    print(f"v8 error: {e}")

print(f"\nTotal records: {nabor.Count()}")
for i in range(nabor.Count()):
    r = nabor.Get(i)
    print(f"  #{i+1} RT XML={conn.XMLString(r.RecordType)}")
