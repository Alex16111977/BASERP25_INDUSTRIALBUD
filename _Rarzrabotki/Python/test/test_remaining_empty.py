# -*- coding: utf-8 -*-
import win32com.client, sys, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

# Get doc
q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка "
    "ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док "
    "ГДЕ НЕ Док.ПометкаУдаления УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
)
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

# Get Kazna INN mapping
conn_kazn = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
Sk = conn_kazn.String

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
    БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
    И БДДС.Период МЕЖДУ &Н И &К
    И БДДС.Регистратор.Организация.А_ПереносЕРП"""
qk.SetParameter("Н", beg)
qk.SetParameter("К", end)
selk = qk.Execute().Choose()
kazna_inn = {}
while selk.Next():
    fio = Sk(selk.ФИО).strip()
    inn = Sk(selk.ИНН).strip()
    if fio and inn:
        kazna_inn[fio] = inn

# Check remaining empty after ИНН fill
print("=== Remaining empty Сотрудник rows after ИНН fill ===")
empty_fios = {}
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    if VZ(row.Сотрудник):
        continue
    fio = S(row.ФИО).strip()
    inn = S(row.ИНН).strip()
    if fio not in empty_fios:
        empty_fios[fio] = inn

for fio, inn in sorted(empty_fios.items()):
    # Check if in kazna
    kazna_inn_val = kazna_inn.get(fio, "NOT IN KAZNA")
    # Check with trimmed trailing space
    kazna_inn_val2 = "N/A"
    for kf, ki in kazna_inn.items():
        if kf.rstrip() == fio.rstrip():
            kazna_inn_val2 = ki
            break
    print(f"  ФИО=[{fio}] len={len(fio)} ИНН=[{inn}] Kazna_exact=[{kazna_inn_val}] Kazna_trim=[{kazna_inn_val2}]")

print(f"\nTotal unique empty: {len(empty_fios)}")
