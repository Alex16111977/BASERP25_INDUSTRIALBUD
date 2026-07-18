# -*- coding: utf-8 -*-
import win32com.client, sys, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

# === Берём документ 000Ц-000001 ===
q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка КАК Ссылка "
    "ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док "
    "ГДЕ Док.Номер = &Номер И НЕ Док.ПометкаУдаления"
)
q.SetParameter("Номер", "000Ц-000001")
sel = q.Execute().Choose()
if not sel.Next():
    print("Doc 000Ц-000001 not found!")
    sys.exit(1)

obj = sel.Ссылка.GetObject()
kaz = obj.А_РаспределениеКазна.Count()
print(f"Doc 000Ц-000001, A_Kazna: {kaz} rows")

# Считаем пустых Сотрудников
sotr_empty_before = 0
for i in range(kaz):
    if not VZ(obj.А_РаспределениеКазна.Get(i).Сотрудник):
        sotr_empty_before += 1
print(f"Empty Сотрудник BEFORE: {sotr_empty_before}")

# === Заполняем ИНН из Казны ===
conn_kazn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
Sk = conn_kazn.String

# Get period
period = obj.ПериодРегистрации
ps = S(period)
parts = ps.split(".")
month, year = int(parts[1]), int(parts[2].split()[0])
beg = datetime.datetime(year, month, 1)
end = datetime.datetime(year + (1 if month == 12 else 0), (month % 12) + 1, 1) - datetime.timedelta(seconds=1)

qk = conn_kazn.NewObject("Запрос")
qk.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    БДДС.Сотрудник.ИНН КАК ИНН,
    БДДС.Сотрудник.Наименование КАК ФИО
ИЗ
    РегистрНакопления.БДДС КАК БДДС
ГДЕ
    БДДС.Период МЕЖДУ &Н И &К
    И БДДС.Регистратор.Организация.А_ПереносЕРП
    И (БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
        ИЛИ БДДС.Регистратор ССЫЛКА Документ.РаспределениеФ2)"""
qk.SetParameter("Н", beg)
qk.SetParameter("К", end)
selk = qk.Execute().Choose()

kazna_inn = {}
while selk.Next():
    fio = Sk(selk.ФИО).strip()
    inn = Sk(selk.ИНН).strip()
    if fio and inn:
        kazna_inn[fio] = inn
        # Normalized versions
        kazna_inn[" ".join(fio.split())] = inn

print(f"Kazna INN: {len(kazna_inn)} entries")

# Fill ИНН in ALL rows
filled = 0
not_found = set()
for i in range(kaz):
    row = obj.А_РаспределениеКазна.Get(i)
    if S(row.ИНН).strip():
        continue
    fio = S(row.ФИО).strip()
    fio_norm = " ".join(fio.split())
    inn = kazna_inn.get(fio) or kazna_inn.get(fio_norm)
    if inn:
        row.ИНН = inn
        filled += 1
    else:
        not_found.add(fio)

print(f"Filled ИНН: {filled} rows")
if not_found:
    print(f"ФИО NOT found in Kazna ({len(not_found)}):")
    for nf in sorted(not_found):
        print(f"  [{nf}]")

# === Вызываем процедуру ===
print("\nCalling ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()...")
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print("CALL: OK")
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# === Результаты ===
afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
anach = obj.НачисленнаяЗарплатаИВзносы.Count()
andfl = obj.НачисленныйНДФЛ.Count()
print(f"\nAFTER: NachFiz={afiz}, Nach={anach}, NDFL={andfl}")

sotr_empty_after = 0
for i in range(obj.А_РаспределениеКазна.Count()):
    if not VZ(obj.А_РаспределениеКазна.Get(i).Сотрудник):
        sotr_empty_after += 1

empty_dept = 0
empty_fl = 0
for i in range(anach):
    if not VZ(obj.НачисленнаяЗарплатаИВзносы.Get(i).ПодразделениеПредприятия):
        empty_dept += 1
for i in range(afiz):
    if not VZ(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i).ФизическоеЛицо):
        empty_fl += 1

print(f"\nEmpty Сотрудник in A_Kazna: {sotr_empty_after} (was {sotr_empty_before})")
print(f"Empty Подразделение in Nach: {empty_dept}")
print(f"Empty ФизЛицо in NachFiz: {empty_fl}")

if sotr_empty_after == 0:
    print("\n*** ALL EMPLOYEES FIXED ***")
if empty_dept == 0 and empty_fl == 0:
    print("*** ALL DEPARTMENTS AND FIZLITSO FILLED ***")
