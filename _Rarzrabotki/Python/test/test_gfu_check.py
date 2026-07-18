# -*- coding: utf-8 -*-
"""Перевірка: чи увімкнена константа ИспользоватьГруппыФинансовогоУчета
та чи є записи в довідниках ГФУ.
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

S = conn.String
VF = conn.ValueIsFilled

# 1. Перевірка констант
print("=" * 60)
print("КОНСТАНТИ")
print("=" * 60)

for name in ["ИспользоватьГруппыФинансовогоУчета",
             "ФормироватьУправленческийБаланс",
             "УчитыватьПрочиеАктивыПассивы",
             "ИспользоватьУправленческийУчет"]:
    try:
        val = eval(f"conn.Constants.{name}.Get()")
        print(f"  {name} = {val}")
    except:
        try:
            val = conn.GetFunctionalOption(name)
            print(f"  {name} (FO) = {val}")
        except:
            print(f"  {name} — не знайдено")

# 2. ГруппыФинансовогоУчетаДоходовРасходов
print("\n" + "=" * 60)
print("ГруппыФинансовогоУчетаДоходовРасходов")
print("=" * 60)
q1 = conn.NewObject("Запрос")
q1.Text = "ВЫБРАТЬ Г.Ссылка, Г.Наименование ИЗ Справочник.ГруппыФинансовогоУчетаДоходовРасходов КАК Г"
res1 = q1.Execute()
sel1 = res1.Select()
cnt = 0
while sel1.Next():
    cnt += 1
    print(f"  {cnt}. {S(sel1.Наименование)}")
print(f"  ВСЬОГО: {cnt}" + (" ← ПУСТО!" if cnt == 0 else ""))

# 3. ГруппыФинансовогоУчетаВнеоборотныхАктивов
print("\n" + "=" * 60)
print("ГруппыФинансовогоУчетаВнеоборотныхАктивов")
print("=" * 60)
q2 = conn.NewObject("Запрос")
q2.Text = "ВЫБРАТЬ Г.Ссылка, Г.Наименование ИЗ Справочник.ГруппыФинансовогоУчетаВнеоборотныхАктивов КАК Г"
res2 = q2.Execute()
sel2 = res2.Select()
cnt2 = 0
while sel2.Next():
    cnt2 += 1
    print(f"  {cnt2}. {S(sel2.Наименование)}")
print(f"  ВСЬОГО: {cnt2}" + (" ← ПУСТО!" if cnt2 == 0 else ""))

# 4. ГруппыФинансовогоУчетаНоменклатуры
print("\n" + "=" * 60)
print("ГруппыФинансовогоУчетаНоменклатуры")
print("=" * 60)
try:
    q3 = conn.NewObject("Запрос")
    q3.Text = "ВЫБРАТЬ Г.Ссылка, Г.Наименование ИЗ Справочник.ГруппыФинансовогоУчетаНоменклатуры КАК Г"
    res3 = q3.Execute()
    sel3 = res3.Select()
    cnt3 = 0
    while sel3.Next():
        cnt3 += 1
        print(f"  {cnt3}. {S(sel3.Наименование)}")
    print(f"  ВСЬОГО: {cnt3}" + (" ← ПУСТО!" if cnt3 == 0 else ""))
except Exception as e:
    print(f"  Помилка: {e}")

print("\n=== DONE ===")
