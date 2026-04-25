# -*- coding: utf-8 -*-
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка "
    "ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док "
    "ГДЕ НЕ Док.ПометкаУдаления УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
)
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

kaz = obj.А_РаспределениеКазна.Count()
print(f"Doc found. A_Kazna: {kaz} rows")

# Check ИНН field - is it filled?
inn_filled = 0
inn_empty = 0
sotr_empty = 0
for i in range(kaz):
    row = obj.А_РаспределениеКазна.Get(i)
    inn = S(row.ИНН).strip()
    has_sotr = VZ(row.Сотрудник)
    if inn:
        inn_filled += 1
    else:
        inn_empty += 1
    if not has_sotr:
        sotr_empty += 1

print(f"ИНН filled: {inn_filled}, ИНН empty: {inn_empty}")
print(f"Сотрудник empty: {sotr_empty}")

if inn_empty > 0 and sotr_empty > 0:
    print("\nИНН не заполнен! Нужно перезагрузить из Казны.")
    print("Заполняю ИНН вручную из Казны через COM...")

    # Connect to Kazna and get INN mapping for doc 267
    conn_kazn = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
    Sk = conn_kazn.String

    import datetime
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

    period = obj.ПериодРегистрации
    # Вычисляем даты через Python
    import datetime
    period_str = S(period)  # "01.01.2026 0:00:00"
    # Parse period to get year/month
    parts = period_str.split(".")
    day = int(parts[0])
    month = int(parts[1])
    year = int(parts[2].split()[0])
    beg = datetime.datetime(year, month, 1)
    if month == 12:
        end = datetime.datetime(year + 1, 1, 1) - datetime.timedelta(seconds=1)
    else:
        end = datetime.datetime(year, month + 1, 1) - datetime.timedelta(seconds=1)
    qk.SetParameter("Н", beg)
    qk.SetParameter("К", end)
    rk = qk.Execute()
    selk = rk.Choose()

    # Build mapping: (ФИО, ПодрКод, СтатьяКод) -> ИНН
    kazna_inn = {}
    kazna_inn_norm = {}  # normalized key
    while selk.Next():
        fio = Sk(selk.ФИО).strip()
        inn = Sk(selk.ИНН).strip()
        if fio and inn:
            kazna_inn[fio] = inn
            # Normalized: strip all extra spaces, lowercase
            norm = " ".join(fio.split()).lower()
            kazna_inn_norm[norm] = inn

    print(f"Kazna INN mapping: {len(kazna_inn)} exact, {len(kazna_inn_norm)} normalized")

    # Fill ИНН in А_РаспределениеКазна
    filled_count = 0
    for i in range(kaz):
        row = obj.А_РаспределениеКазна.Get(i)
        if S(row.ИНН).strip():
            continue  # already filled
        fio = S(row.ФИО).strip()
        if not fio:
            continue
        # Try exact match
        inn = kazna_inn.get(fio)
        if not inn:
            # Try normalized
            norm = " ".join(fio.split()).lower()
            inn = kazna_inn_norm.get(norm)
        if inn:
            row.ИНН = inn
            filled_count += 1

    print(f"Filled ИНН for {filled_count} rows")

# Now call the procedure
print("\nCalling ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()...")
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print("CALL: OK")
except Exception as e:
    print(f"ERROR: {e}")
    sys.exit(1)

# Check results
afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
anach = obj.НачисленнаяЗарплатаИВзносы.Count()
andfl = obj.НачисленныйНДФЛ.Count()
print(f"\nAFTER: NachFiz={afiz}, Nach={anach}, NDFL={andfl}")

# Count empty
empty_dept = 0
empty_fl = 0
for i in range(anach):
    r = obj.НачисленнаяЗарплатаИВзносы.Get(i)
    if not VZ(r.ПодразделениеПредприятия):
        empty_dept += 1

for i in range(afiz):
    r = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i)
    if not VZ(r.ФизическоеЛицо):
        empty_fl += 1

# Count empty Sotr in Kazna after fix
sotr_empty_after = 0
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    if not VZ(row.Сотрудник):
        sotr_empty_after += 1

print(f"Empty Сотрудник in A_Kazna AFTER: {sotr_empty_after} (was {sotr_empty})")
print(f"Empty Подразделение in Nach: {empty_dept}")
print(f"Empty ФизЛицо in NachFiz: {empty_fl}")

if empty_dept == 0 and empty_fl == 0:
    print("\n*** ALL CHECKS PASSED ***")
elif sotr_empty_after < sotr_empty:
    print(f"\n*** IMPROVED: Сотрудник fixed {sotr_empty - sotr_empty_after} rows ***")
