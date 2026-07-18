# -*- coding: utf-8 -*-
"""
Тест: проверка полей виртуальных таблиц регистра ДенежныеСредстваФинАгенты в Казне.
Остаточный регистр -> .Обороты() дает только СуммаОборот, НЕ СуммаПриходОборот/СуммаРасходОборот.
"""
import sys
import pythoncom
import win32com.client

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_KAZNA = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'
conn = v8.Connect(CONN_KAZNA)

print("=== Test 1: .Обороты() field names ===")
q = conn.NewObject("Query")
q.Text = (
    "SELECT TOP 1 * "
    "FROM AccumulationRegister.ДенежныеСредстваФинАгенты.Turnovers(&D1, &D2, , ) AS T"
)
q.SetParameter("D1", win32com.client.pywintypes.TimeType(2026, 1, 1, 0, 0, 0))
q.SetParameter("D2", win32com.client.pywintypes.TimeType(2026, 1, 31, 23, 59, 59))
res = q.Execute()
cols = res.Columns
print(f"Columns count: {cols.Count()}")
for i in range(cols.Count()):
    col = cols.Get(i)
    print(f"  [{i}] {col.Name}")

print()
print("=== Test 2: .ОстаткиИОбороты() field names ===")
q2 = conn.NewObject("Query")
q2.Text = (
    "SELECT TOP 1 * "
    "FROM AccumulationRegister.ДенежныеСредстваФинАгенты.BalanceAndTurnovers(&D1, &D2, , , ) AS T"
)
q2.SetParameter("D1", win32com.client.pywintypes.TimeType(2026, 1, 1, 0, 0, 0))
q2.SetParameter("D2", win32com.client.pywintypes.TimeType(2026, 1, 31, 23, 59, 59))
res2 = q2.Execute()
cols2 = res2.Columns
print(f"Columns count: {cols2.Count()}")
for i in range(cols2.Count()):
    col2 = cols2.Get(i)
    print(f"  [{i}] {col2.Name}")

print()
print("=== Test 3: Raw movements field names ===")
q3 = conn.NewObject("Query")
q3.Text = (
    "SELECT TOP 1 * "
    "FROM AccumulationRegister.ДенежныеСредстваФинАгенты AS T "
    "WHERE T.Period BETWEEN &D1 AND &D2"
)
q3.SetParameter("D1", win32com.client.pywintypes.TimeType(2026, 1, 1, 0, 0, 0))
q3.SetParameter("D2", win32com.client.pywintypes.TimeType(2026, 1, 31, 23, 59, 59))
res3 = q3.Execute()
cols3 = res3.Columns
print(f"Columns count: {cols3.Count()}")
for i in range(cols3.Count()):
    col3 = cols3.Get(i)
    print(f"  [{i}] {col3.Name}")

print()
print("DONE")
