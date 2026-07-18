# -*- coding: utf-8 -*-
"""
Test: contract creation with А_ПодразделениеОказаниеУслугПодразделению
Document: Передача в эксплуатацию ІБ00-000641 от 16.12.2025
Etalon: Взаимозачет задолженности 000Ц-000433
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

print("=" * 70)
print("1. DISBALANCE for IB00-000641")
print("=" * 70)

# Find doc by query
q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Об.Подразделение КАК П, Об.Организация КАК О, "
    "   Об.СуммаПриход КАК Д, Об.СуммаРасход КАК К, "
    "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
    "ГДЕ Об.Регистратор = &Р "
    "   И НЕ Об.Статья В ("
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),"
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"
)

# Find the document
qd = conn.NewObject("Query")
qd.Text = (
    "ВЫБРАТЬ Д.Ссылка ИЗ Документ.ПередачаВЭксплуатацию КАК Д "
    "ГДЕ Д.Номер ПОДОБНО \"%000641\""
)
rd = qd.Execute().Select()
if not rd.Next():
    print("Document IB00-000641 not found!")
    exit()

doc_ref = rd.Ссылка
print(f"  Document: {doc_ref}")

q.SetParameter("Р", doc_ref)
res = q.Execute().Select()

debtor = None
creditor = None
org = None
sum_vzaim = 0

while res.Next():
    dept = str(res.П)
    ctrl = round(float(res.Контроль), 2)
    print(f"  {dept}: Dt={res.Д:.2f} Kt={res.К:.2f} Ctrl={ctrl}")
    org = res.О
    if ctrl > 0:
        debtor = res.П
        sum_vzaim = ctrl
    elif ctrl < 0:
        creditor = res.П

print(f"\n  Debtor: {debtor} (+{sum_vzaim})")
print(f"  Creditor: {creditor}")
print(f"  Org: {org}")

# 2. Find Partners
print("\n" + "=" * 70)
print("2. PARTNERS & COUNTERPARTIES")
print("=" * 70)

def find_partner(dept):
    qp = conn.NewObject("Query")
    qp.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка КАК Партнер "
        "ИЗ Справочник.Партнеры КАК П "
        "ГДЕ П.А_Подразделение = &Подр И НЕ П.ПометкаУдаления"
    )
    qp.SetParameter("Подр", dept)
    rp = qp.Execute().Select()
    if rp.Next():
        return rp.Партнер
    return None

def find_contragent(partner):
    qc = conn.NewObject("Query")
    qc.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 К.Ссылка КАК Контрагент "
        "ИЗ Справочник.Контрагенты КАК К "
        "ГДЕ К.Партнер = &П И НЕ К.ПометкаУдаления"
    )
    qc.SetParameter("П", partner)
    rc = qc.Execute().Select()
    if rc.Next():
        return rc.Контрагент
    return None

partner_deb = find_partner(debtor)
partner_cred = find_partner(creditor)
contragent_deb = find_contragent(partner_deb) if partner_deb else None
contragent_cred = find_contragent(partner_cred) if partner_cred else None

print(f"  Debtor partner: {partner_deb}")
print(f"  Debtor contragent: {contragent_deb}")
print(f"  Creditor partner: {partner_cred}")
print(f"  Creditor contragent: {contragent_cred}")

# 3. Search contracts with new logic
print("\n" + "=" * 70)
print("3. CONTRACT SEARCH (new logic)")
print("=" * 70)

def find_contract(contragent, dept, typ_dogovor, org, podr_okazania):
    qc = conn.NewObject("Query")
    qc.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка, Д.Наименование КАК Н, "
        "   Д.А_ПодразделениеОказаниеУслугПодразделению КАК ПодрОказ "
        "ИЗ Справочник.ДоговорыКонтрагентов КАК Д "
        "ГДЕ Д.Контрагент = &К И Д.ТипДоговора = &Т "
        "   И Д.Подразделение = &П И Д.Организация = &О "
        "   И Д.А_ПодразделениеОказаниеУслугПодразделению = &ПО "
        "   И НЕ Д.ПометкаУдаления"
    )
    qc.SetParameter("К", contragent)
    qc.SetParameter("Т", typ_dogovor)
    qc.SetParameter("П", dept)
    qc.SetParameter("О", org)
    qc.SetParameter("ПО", podr_okazania)
    rc = qc.Execute().Select()
    if rc.Next():
        return rc.Ссылка, str(rc.Н).strip(), rc.ПодрОказ
    return None, None, None

enum_supplier = conn.Перечисления.ТипыДоговоров.СПоставщиком
enum_buyer = conn.Перечисления.ТипыДоговоров.СПокупателем

# Debtor contract: А_ПодрОказанияУслуг = Creditor (opposite)
ref_d, name_d, podr_d = find_contract(contragent_deb, debtor, enum_supplier, org, creditor)
print(f"  Debtor contract: {name_d}")
print(f"    А_ПодрОказанияУслуг: {podr_d}")
if ref_d:
    print(f"    FOUND!")
else:
    print(f"    NOT FOUND - will be created")

# Creditor contract: А_ПодрОказанияУслуг = Debtor (opposite)
ref_c, name_c, podr_c = find_contract(contragent_cred, creditor, enum_buyer, org, debtor)
print(f"\n  Creditor contract: {name_c}")
print(f"    А_ПодрОказанияУслуг: {podr_c}")
if ref_c:
    print(f"    FOUND!")
else:
    print(f"    NOT FOUND - will be created")

# 4. Compare with etalon 000Ц-000433
print("\n" + "=" * 70)
print("4. ETALON 000Ц-000433")
print("=" * 70)

qe = conn.NewObject("Query")
qe.Text = (
    "ВЫБРАТЬ ДЗ.ОбъектРасчетов КАК ОР "
    "ИЗ Документ.ВзаимозачетЗадолженности.ДебиторскаяЗадолженность КАК ДЗ "
    "ГДЕ ДЗ.Ссылка.СуммаРегл = 18553.31 И ДЗ.Ссылка.Дата МЕЖДУ &Н И &К"
)
qe.SetParameter("Н", datetime.datetime(2025, 12, 16))
qe.SetParameter("К", datetime.datetime(2025, 12, 16, 23, 59, 59))
re = qe.Execute().Select()
if re.Next():
    or_ref = re.ОР
    if conn.ValueIsFilled(or_ref):
        dogovor_et = or_ref.Договор
        print(f"  Etalon debtor ОР: {or_ref}")
        print(f"  Etalon debtor contract: {dogovor_et}")
        if conn.ValueIsFilled(dogovor_et):
            print(f"    Name: {dogovor_et.Наименование}")
            podr_et = dogovor_et.А_ПодразделениеОказаниеУслугПодразделению
            print(f"    А_ПодрОказанияУслуг: {podr_et}")

qe2 = conn.NewObject("Query")
qe2.Text = (
    "ВЫБРАТЬ КЗ.ОбъектРасчетов КАК ОР "
    "ИЗ Документ.ВзаимозачетЗадолженности.КредиторскаяЗадолженность КАК КЗ "
    "ГДЕ КЗ.Ссылка.СуммаРегл = 18553.31 И КЗ.Ссылка.Дата МЕЖДУ &Н И &К"
)
qe2.SetParameter("Н", datetime.datetime(2025, 12, 16))
qe2.SetParameter("К", datetime.datetime(2025, 12, 16, 23, 59, 59))
re2 = qe2.Execute().Select()
if re2.Next():
    or_ref2 = re2.ОР
    if conn.ValueIsFilled(or_ref2):
        dogovor_et2 = or_ref2.Договор
        print(f"\n  Etalon creditor ОР: {or_ref2}")
        print(f"  Etalon creditor contract: {dogovor_et2}")
        if conn.ValueIsFilled(dogovor_et2):
            print(f"    Name: {dogovor_et2.Наименование}")
            podr_et2 = dogovor_et2.А_ПодразделениеОказаниеУслугПодразделению
            print(f"    А_ПодрОказанияУслуг: {podr_et2}")

print("\n" + "=" * 70)
print("DONE")
