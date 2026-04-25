# -*- coding: utf-8 -*-
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Find latest document
q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка, "
    "Док.Номер КАК Номер, "
    "Док.ПериодРегистрации КАК Период "
    "ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док "
    "ГДЕ НЕ Док.ПометкаУдаления "
    "УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
)
res = q.Execute()
sel = res.Choose()
sel.Next()

print("Document:", S(sel.Номер), "Period:", S(sel.Период))
ref = sel.Ссылка
obj = ref.GetObject()

# Before
kaz = obj.А_РаспределениеКазна.Count()
nal = obj.А_РаспределениеНалоги.Count()
print(f"A_Kazna: {kaz}, A_Nalogi: {nal}")

# Count empty employees BEFORE
empty_before = 0
for i in range(kaz):
    row = obj.А_РаспределениеКазна.Get(i)
    if not conn.ЗначениеЗаполнено(row.Сотрудник):
        empty_before += 1
print(f"Empty Sotrudnik in A_Kazna: {empty_before}")

bfiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
print(f"BEFORE: NachFiz={bfiz}")

# Call the procedure
print("\nCalling procedure...")
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print("CALL: OK")
except Exception as e:
    print(f"ERROR: {e}")
    sys.exit(1)

afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
anach = obj.НачисленнаяЗарплатаИВзносы.Count()
andfl = obj.НачисленныйНДФЛ.Count()
print(f"AFTER: NachFiz={afiz}, Nach={anach}, NDFL={andfl}")

# Check for empty departments and empty FizLitso
empty_dept = 0
empty_fl = 0
for i in range(anach):
    r = obj.НачисленнаяЗарплатаИВзносы.Get(i)
    if not conn.ЗначениеЗаполнено(r.ПодразделениеПредприятия):
        empty_dept += 1

for i in range(afiz):
    r = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i)
    if not conn.ЗначениеЗаполнено(r.ФизическоеЛицо):
        empty_fl += 1

print(f"\nEmpty Подразделение in Nach: {empty_dept} (was ~1)")
print(f"Empty ФизЛицо in NachFiz: {empty_fl} (was ~many)")

if empty_dept == 0 and empty_fl == 0:
    print("\nALL CHECKS PASSED - no empty departments or persons!")
else:
    print(f"\nWARN: still have issues (expected with 124 unfixed rows in A_Kazna)")
    print("These will be fixed after re-running Kazna loading")

print("\nDone")
