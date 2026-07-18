# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn_kazn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
S = conn_kazn.String

q = conn_kazn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ
    СУММА(ВЫБОР КОГДА С.ИНН = "" ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК БезИНН,
    СУММА(ВЫБОР КОГДА С.ИНН <> "" ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СИНН,
    КОЛИЧЕСТВО(*) КАК Всего
ИЗ Справочник.Сотрудники КАК С
ГДЕ НЕ С.ПометкаУдаления"""
r = q.Execute().Choose()
r.Next()
print(f"Казна Сотрудники: всего={r.Всего}, с ИНН={r.СИНН}, без ИНН={r.БезИНН}")

# Show first 5 without INN
q2 = conn_kazn.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 5 С.Наименование КАК ФИО, С.Код КАК Код
ИЗ Справочник.Сотрудники КАК С ГДЕ С.ИНН = "" И НЕ С.ПометкаУдаления"""
r2 = q2.Execute().Choose()
print("Без ИНН:")
while r2.Next():
    print(f"  {S(r2.ФИО)}, Код={S(r2.Код)}")
