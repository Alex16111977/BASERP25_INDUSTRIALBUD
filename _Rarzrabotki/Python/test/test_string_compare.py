# -*- coding: utf-8 -*-
import win32com.client, sys, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q.SetParameter("Н", "000Ц-000001")
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

# Get first empty-sotr row ФИО
fio_doc = None
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    if not conn.ЗначениеЗаполнено(row.Сотрудник):
        fio_doc = row.ФИО  # raw COM value, NOT converted
        fio_doc_str = S(row.ФИО)
        fio_doc_name = S(row.ФИО).strip()
        print(f"Doc ФИО raw type: {type(fio_doc)}")
        print(f"Doc ФИО S():     [{fio_doc_str}] len={len(fio_doc_str)}")
        print(f"Doc ФИО strip(): [{fio_doc_name}] len={len(fio_doc_name)}")
        print(f"Doc ФИО hex:     {fio_doc_str.encode('utf-8')[:50].hex()}")
        break

# Get same person from Kazna
conn_kazn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
Sk = conn_kazn.String

period = obj.ПериодРегистрации
ps = S(period)
parts = ps.split(".")
month, year = int(parts[1]), int(parts[2].split()[0])
beg = datetime.datetime(year, month, 1)
end = datetime.datetime(year + (1 if month == 12 else 0), (month % 12) + 1, 1) - datetime.timedelta(seconds=1)

# Search in Kazna by name fragment
search_name = fio_doc_name[:10] if fio_doc_name else "Безсмертн"
qk = conn_kazn.NewObject("Запрос")
qk.Text = f"""ВЫБРАТЬ ПЕРВЫЕ 1
    БДДС.Сотрудник.Наименование КАК ФИО,
    БДДС.Сотрудник.ИНН КАК ИНН,
    БДДС.Сотрудник.Код КАК КодС
ИЗ РегистрНакопления.БДДС КАК БДДС
ГДЕ БДДС.Период МЕЖДУ &Н И &К
    И БДДС.Регистратор.Организация.А_ПереносЕРП
    И БДДС.Сотрудник.Наименование ПОДОБНО "%{search_name}%"
"""
qk.SetParameter("Н", beg)
qk.SetParameter("К", end)
selk = qk.Execute().Choose()
if selk.Next():
    fio_kazna = selk.ФИО  # raw COM
    fio_kazna_str = Sk(selk.ФИО)
    print(f"\nKazna ФИО raw type: {type(fio_kazna)}")
    print(f"Kazna ФИО Sk():    [{fio_kazna_str}] len={len(fio_kazna_str)}")
    print(f"Kazna ФИО strip(): [{fio_kazna_str.strip()}] len={len(fio_kazna_str.strip())}")
    print(f"Kazna ФИО hex:     {fio_kazna_str.encode('utf-8')[:50].hex()}")

    # Direct comparison
    print(f"\n=== Comparison ===")
    print(f"S(doc) == Sk(kazna): {fio_doc_str == fio_kazna_str}")
    print(f"strip(doc) == strip(kazna): {fio_doc_str.strip() == fio_kazna_str.strip()}")

    # NOW: simulate what 1C Строка() does
    # In 1C, Строка(row.ФИО) for a fixed-length string field returns padded value
    # But Строка(Выборка.ФИО) from query returns the query result value

    # The REAL question: does the 1С code see different values?
    # Let's check what 1C String() returns by using Вычислить

print("\n=== What does 1C see? ===")
# Can't use Вычислить on COM connection, but we can check field metadata
print("Doc А_РаспределениеКазна.ФИО is a Строка(П150) - fixed 150 chars")
print("Kazna query returns Строка from Справочник.Сотрудники.Наименование")
print("")
print("CONCLUSION: Doc ФИО is padded with spaces to 150 chars!")
print(f"Doc ФИО actual len with spaces: checking...")

# Check actual COM string length
fio_com_val = str(fio_doc) if fio_doc else ""
print(f"Doc ФИО via str(): [{fio_com_val}] len={len(fio_com_val)}")

# Check if it's padded
if fio_doc_str and len(fio_doc_str) > len(fio_doc_str.rstrip()):
    print(f"TRAILING SPACES: {len(fio_doc_str) - len(fio_doc_str.rstrip())} spaces")
    print(f"Doc rstrip: [{fio_doc_str.rstrip()}] len={len(fio_doc_str.rstrip())}")
