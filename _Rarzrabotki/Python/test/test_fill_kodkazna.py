# -*- coding: utf-8 -*-
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

# Doc 000Ц-000001
q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Док.Ссылка КАК Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q.SetParameter("Н", "000Ц-000001")
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

kaz = obj.А_РаспределениеКазна.Count()

# Before
sotr_empty = 0
test_people = {}
for i in range(kaz):
    row = obj.А_РаспределениеКазна.Get(i)
    fio = S(row.ФИО).strip()
    has_sotr = VZ(row.Сотрудник)
    inn = S(row.ИНН).strip()
    kod = S(row.КодСотрудникаКазна).strip()
    if not has_sotr:
        sotr_empty += 1
    if fio in ("Безсмертний Анатолій Семенович", "Стаднік Анатолій Петрович",
               "Омельченко Сергій Григорович", "Мудрагель Ігор Валентинович"):
        if fio not in test_people:
            test_people[fio] = {"sotr": S(row.Сотрудник) if has_sotr else "EMPTY", "inn": inn, "kod": kod}

print(f"Doc 000Ц-000001, A_Kazna: {kaz} rows")
print(f"Empty Сотрудник BEFORE: {sotr_empty}")
print(f"\nTest people BEFORE:")
for fio, data in sorted(test_people.items()):
    print(f"  {fio}: Сотр={data['sotr'][:20]}, ИНН={data['inn']}, КодКазна={data['kod']}")

# Call
print("\nCalling ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()...")
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print("CALL: OK")
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# After
sotr_empty_after = 0
test_people_after = {}
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    fio = S(row.ФИО).strip()
    has_sotr = VZ(row.Сотрудник)
    if not has_sotr:
        sotr_empty_after += 1
    if fio in ("Безсмертний Анатолій Семенович", "Стаднік Анатолій Петрович",
               "Омельченко Сергій Григорович", "Мудрагель Ігор Валентинович"):
        if fio not in test_people_after:
            test_people_after[fio] = {
                "sotr": S(row.Сотрудник)[:30] if has_sotr else "EMPTY",
                "inn": S(row.ИНН).strip(),
                "kod": S(row.КодСотрудникаКазна).strip()
            }

afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
anach = obj.НачисленнаяЗарплатаИВзносы.Count()
andfl = obj.НачисленныйНДФЛ.Count()

empty_dept = 0
empty_fl = 0
for i in range(anach):
    if not VZ(obj.НачисленнаяЗарплатаИВзносы.Get(i).ПодразделениеПредприятия):
        empty_dept += 1
for i in range(afiz):
    if not VZ(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i).ФизическоеЛицо):
        empty_fl += 1

print(f"\n=== RESULTS ===")
print(f"Empty Сотрудник: {sotr_empty} → {sotr_empty_after}")
print(f"NachFiz={afiz}, Nach={anach}, NDFL={andfl}")
print(f"Empty Подразделение: {empty_dept}")
print(f"Empty ФизЛицо: {empty_fl}")

print(f"\nTest people AFTER:")
for fio, data in sorted(test_people_after.items()):
    print(f"  {fio}: Сотр={data['sotr']}, ИНН={data['inn']}, КодКазна={data['kod']}")

# Check ФизЛица created from Kazna
q2 = conn.NewObject("Запрос")
q2.Text = "ВЫБРАТЬ Ф.Наименование КАК ФИО, Ф.КодПоДРФО КАК ДРФО, Ф.А_КодСотрудникаКазна КАК КодК, Ф.А_ИННКазна КАК ИННКазна ИЗ Справочник.ФизическиеЛица КАК Ф ГДЕ Ф.А_КазнаСоздан = ИСТИНА"
r2 = q2.Execute().Choose()
print(f"\nФизЛица created from Kazna (А_КазнаСоздан=True):")
while r2.Next():
    print(f"  {S(r2.ФИО)}: ДРФО={S(r2.ДРФО)}, КодКазна={S(r2.КодК)}, ИННКазна={S(r2.ИННКазна)}")
