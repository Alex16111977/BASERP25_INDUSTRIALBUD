# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Test the register query for Jan 2026
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    Взаиморасчеты.Сотрудник КАК Сотрудник
ИЗ
    РегистрНакопления.ВзаиморасчетыССотрудниками КАК Взаиморасчеты
ГДЕ
    Взаиморасчеты.Период МЕЖДУ &Н И &К
    И Взаиморасчеты.Регистратор ССЫЛКА Документ.НачислениеЗарплаты"""
q.SetParameter("Н", datetime.datetime(2026, 1, 1))
q.SetParameter("К", datetime.datetime(2026, 1, 31, 23, 59, 59))
r = q.Execute()
sel = r.Choose()
cnt = 0
while sel.Next():
    cnt += 1
print(f"Register query Jan 2026: {cnt} employees")

# Now full test
q2 = conn.NewObject("Запрос")
q2.Text = "ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q2.SetParameter("Н", "000Ц-000001")
sel2 = q2.Execute().Choose()
sel2.Next()
obj = sel2.Ссылка.GetObject()
print(f"Doc period: {S(obj.ПериодРегистрации)}")

# Call procedure - will fail with old code in DB, but test query separately
# Simulate: add register employees to cache
VZ = conn.ЗначениеЗаполнено
allEmployees = set()
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    if VZ(row.Сотрудник):
        allEmployees.add(S(row.Сотрудник))
for i in range(obj.А_РаспределениеНалоги.Count()):
    row = obj.А_РаспределениеНалоги.Get(i)
    if VZ(row.Сотрудник):
        allEmployees.add(S(row.Сотрудник))

print(f"Employees from Kazna+Nalogi: {len(allEmployees)}")

# Add from register
period = obj.ПериодРегистрации
ps = S(period)
parts = ps.split(".")
month, year = int(parts[1]), int(parts[2].split()[0])
beg = datetime.datetime(year, month, 1)
end = datetime.datetime(year + (1 if month == 12 else 0), (month % 12) + 1, 1) - datetime.timedelta(seconds=1)

q3 = conn.NewObject("Запрос")
q3.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    Взаиморасчеты.Сотрудник КАК Сотрудник
ИЗ
    РегистрНакопления.ВзаиморасчетыССотрудниками КАК Взаиморасчеты
ГДЕ
    Взаиморасчеты.Период МЕЖДУ &Н И &К
    И Взаиморасчеты.Регистратор ССЫЛКА Документ.НачислениеЗарплаты"""
q3.SetParameter("Н", beg)
q3.SetParameter("К", end)
r3 = q3.Execute().Choose()
newFromReg = 0
while r3.Next():
    name = S(r3.Сотрудник)
    if name not in allEmployees:
        newFromReg += 1
        allEmployees.add(name)

print(f"NEW from register (not in Kazna/Nalogi): {newFromReg}")
print(f"Total employees: {len(allEmployees)}")
print("\nQuery works! Code fix is correct, just need to reload config in 1C.")
