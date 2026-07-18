# -*- coding: utf-8 -*-
"""
Test: check contract mismatch in vzaimozachet 000Ц-000252
for ПеремещениеТоваров СТ00-000192 от 08.12.2025
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    """Safe string from COM object"""
    try:
        if obj is None:
            return "<None>"
        if not conn.ValueIsFilled(obj):
            return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

print("=" * 70)
print("1. Vzaimozachet 000C-000252")
print("=" * 70)

q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка, Док.Номер, Док.СуммаРегл, "
    "   Док.КонтрагентДебитор КАК КДеб, Док.КонтрагентКредитор КАК ККред "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.Дата МЕЖДУ &Н И &К И Док.Номер ПОДОБНО \"%000252\""
)
q.SetParameter("Н", datetime.datetime(2025, 12, 8))
q.SetParameter("К", datetime.datetime(2025, 12, 8, 23, 59, 59))

res = q.Execute().Select()
vz_ref = None
while res.Next():
    vz_ref = res.Ссылка
    print(f"  Num: {s(res.Номер).strip()}")
    print(f"  Sum: {res.СуммаРегл}")
    print(f"  Debtor contragent: {s(res.КДеб)}")
    print(f"  Creditor contragent: {s(res.ККред)}")

if not vz_ref:
    print("NOT FOUND!"); exit()

# 2. Debtor tab
print("\n--- Debtor tab ---")
q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ ДЗ.Партнер, ДЗ.ОбъектРасчетов КАК ОР "
    "ИЗ Документ.ВзаимозачетЗадолженности.ДебиторскаяЗадолженность КАК ДЗ "
    "ГДЕ ДЗ.Ссылка = &Вз"
)
q2.SetParameter("Вз", vz_ref)
r2 = q2.Execute().Select()
dog_deb_name = ""
dog_deb_podr = ""
while r2.Next():
    print(f"  Partner: {s(r2.Партнер)}")
    or_ref = r2.ОР
    if conn.ValueIsFilled(or_ref):
        dog = or_ref.Договор
        if conn.ValueIsFilled(dog):
            dog_deb_name = s(dog.Наименование)
            podr = dog.А_ПодразделениеОказаниеУслугПодразделению
            dog_deb_podr = s(podr)
            print(f"  Contract: {dog_deb_name}")
            print(f"  А_ПодрОказ: {dog_deb_podr}")
        else:
            print(f"  Contract: <EMPTY>")
    else:
        print(f"  ОбъектРасчетов: <EMPTY>")

# 3. Creditor tab
print("\n--- Creditor tab ---")
q3 = conn.NewObject("Query")
q3.Text = (
    "ВЫБРАТЬ КЗ.Партнер, КЗ.ОбъектРасчетов КАК ОР "
    "ИЗ Документ.ВзаимозачетЗадолженности.КредиторскаяЗадолженность КАК КЗ "
    "ГДЕ КЗ.Ссылка = &Вз"
)
q3.SetParameter("Вз", vz_ref)
r3 = q3.Execute().Select()
dog_kred_name = ""
dog_kred_podr = ""
while r3.Next():
    print(f"  Partner: {s(r3.Партнер)}")
    or_ref = r3.ОР
    if conn.ValueIsFilled(or_ref):
        dog = or_ref.Договор
        if conn.ValueIsFilled(dog):
            dog_kred_name = s(dog.Наименование)
            podr = dog.А_ПодразделениеОказаниеУслугПодразделению
            dog_kred_podr = s(podr)
            print(f"  Contract: {dog_kred_name}")
            print(f"  А_ПодрОказ: {dog_kred_podr}")
        else:
            print(f"  Contract: <EMPTY>")
    else:
        print(f"  ОбъектРасчетов: <EMPTY>")

# 4. Source document disbalance
print("\n" + "=" * 70)
print("4. Source document departments")
print("=" * 70)

doc_osn = vz_ref.А_ДокументОснованиеДляБаланса
if not conn.ValueIsFilled(doc_osn):
    # Find by number
    q_find = conn.NewObject("Query")
    q_find.Text = (
        "ВЫБРАТЬ Д.Ссылка ИЗ Документ.ПеремещениеТоваров КАК Д "
        "ГДЕ Д.Номер ПОДОБНО \"%000192\""
    )
    r_find = q_find.Execute().Select()
    if r_find.Next():
        doc_osn = r_find.Ссылка

print(f"  Source doc: {s(doc_osn)}")

q4 = conn.NewObject("Query")
q4.Text = (
    "ВЫБРАТЬ Об.Подразделение КАК П, "
    "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
    "ГДЕ Об.Регистратор = &Р "
    "   И НЕ Об.Статья В ("
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),"
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"
)
q4.SetParameter("Р", doc_osn)
r4 = q4.Execute().Select()

dept_debtor_name = ""
dept_creditor_name = ""
while r4.Next():
    dept = s(r4.П)
    ctrl = round(float(r4.Контроль), 2)
    role = "DEBTOR" if ctrl > 0 else "CREDITOR"
    print(f"  {dept}: Ctrl={ctrl} ({role})")
    if ctrl > 0:
        dept_debtor_name = dept
    elif ctrl < 0:
        dept_creditor_name = dept

# 5. Verdict
print("\n" + "=" * 70)
print("5. VERDICT: contract match check")
print("=" * 70)

print(f"\n  Debtor dept: {dept_debtor_name}")
print(f"  Creditor dept: {dept_creditor_name}")
print(f"\n  Debtor contract: {dog_deb_name}")
print(f"    А_ПодрОказ: {dog_deb_podr}")
print(f"    Expected А_ПодрОказ: {dept_creditor_name} (creditor = opposite)")
deb_match = (dog_deb_podr == dept_creditor_name)
print(f"    MATCH: {deb_match}")

print(f"\n  Creditor contract: {dog_kred_name}")
print(f"    А_ПодрОказ: {dog_kred_podr}")
print(f"    Expected А_ПодрОказ: {dept_debtor_name} (debtor = opposite)")
kred_match = (dog_kred_podr == dept_debtor_name)
print(f"    MATCH: {kred_match}")

mismatch = not deb_match or not kred_match
print(f"\n  >>> РасхождениеДоговор = {mismatch} <<<")
if mismatch:
    print(f"  *** CONTRACTS ARE WRONG - need to recreate vzaimozachet! ***")
else:
    print(f"  Contracts are correct.")
