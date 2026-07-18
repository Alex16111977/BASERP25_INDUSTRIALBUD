# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")

# BuhBud
conn_buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = conn_buh.String

# Казна
conn_kazn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
Sk = conn_kazn.String

names = ["Анурін", "Безсмертн", "Билім", "Гусаков", "Камінськ", "Киченко",
         "Коваленко Олександр", "Липа Вадим", "Манжела", "Мудрагель", "Мудрий Роман",
         "Наконечний", "Омельченко Сергій", "Самогородськ", "Стаднік",
         "Строкач Андрій", "Хоменко Віталій", "Чернобров"]

print("=== Check 18 missing employees across 3 bases ===")
print(f"{'Name':<30} {'BuhBud FL':<12} {'Kazna Sotr':<12} {'ERP Sotr':<12}")
print("-" * 70)

# ERP
conn_erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
Se = conn_erp.String

for name in names:
    # BuhBud
    q1 = conn_buh.NewObject("Запрос")
    q1.Text = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ф.Наименование КАК Н ИЗ Справочник.ФизическиеЛица КАК Ф ГДЕ Ф.Наименование ПОДОБНО "%' + name + '%"'
    r1 = q1.Execute()
    buh = "YES" if not r1.IsEmpty() else "NO"

    # Казна
    q2 = conn_kazn.NewObject("Запрос")
    q2.Text = 'ВЫБРАТЬ ПЕРВЫЕ 1 С.Наименование КАК Н ИЗ Справочник.Сотрудники КАК С ГДЕ С.Наименование ПОДОБНО "%' + name + '%"'
    r2 = q2.Execute()
    kazn = "YES" if not r2.IsEmpty() else "NO"

    # ERP
    q3 = conn_erp.NewObject("Запрос")
    q3.Text = 'ВЫБРАТЬ ПЕРВЫЕ 1 С.Наименование КАК Н ИЗ Справочник.Сотрудники КАК С ГДЕ С.Наименование ПОДОБНО "%' + name + '%"'
    r3 = q3.Execute()
    erp = "YES" if not r3.IsEmpty() else "NO"

    print(f"  {name:<28} {buh:<12} {kazn:<12} {erp:<12}")

print("\nDone")
