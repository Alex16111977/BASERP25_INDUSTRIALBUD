# -*- coding: utf-8 -*-
"""
Find vzaimozachety with WRONG contracts (without А_ПодрОказанияУслуг)
in December 2025 to test РасхождениеДоговор detection
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    try:
        if obj is None: return "<None>"
        if not conn.ValueIsFilled(obj): return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

print("Searching vzaimozachety with А_ДокОснование filled (our auto-created)...")
print("Checking contracts for А_ПодрОказанияУслуг...\n")

q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка, Док.Номер КАК Н, Док.СуммаРегл КАК С, "
    "   Док.А_ДокументОснованиеДляБаланса КАК Осн "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.Дата МЕЖДУ &Нач И &Кон "
    "   И Док.А_ДокументОснованиеДляБаланса <> НЕОПРЕДЕЛЕНО "
    "   И Док.Проведен И НЕ Док.ПометкаУдаления "
    "УПОРЯДОЧИТЬ ПО Док.Дата"
)
q.SetParameter("Нач", datetime.datetime(2025, 12, 1))
q.SetParameter("Кон", datetime.datetime(2025, 12, 31, 23, 59, 59))

res = q.Execute().Select()

ok_count = 0
bad_count = 0

while res.Next():
    vz_ref = res.Ссылка
    num = s(res.Н).strip()

    # Get debtor OR
    q2 = conn.NewObject("Query")
    q2.Text = (
        "ВЫБРАТЬ ДЗ.ОбъектРасчетов КАК ОР "
        "ИЗ Документ.ВзаимозачетЗадолженности.ДебиторскаяЗадолженность КАК ДЗ "
        "ГДЕ ДЗ.Ссылка = &Вз"
    )
    q2.SetParameter("Вз", vz_ref)
    r2 = q2.Execute().Select()

    deb_podr_ok = False
    kred_podr_ok = False
    deb_contract = ""
    kred_contract = ""

    if r2.Next():
        or_ref = r2.ОР
        if conn.ValueIsFilled(or_ref):
            dog = or_ref.Договор
            if conn.ValueIsFilled(dog):
                deb_contract = s(dog.Наименование)
                podr = dog.А_ПодразделениеОказаниеУслугПодразделению
                deb_podr_ok = conn.ValueIsFilled(podr)

    # Get creditor OR
    q3 = conn.NewObject("Query")
    q3.Text = (
        "ВЫБРАТЬ КЗ.ОбъектРасчетов КАК ОР "
        "ИЗ Документ.ВзаимозачетЗадолженности.КредиторскаяЗадолженность КАК КЗ "
        "ГДЕ КЗ.Ссылка = &Вз"
    )
    q3.SetParameter("Вз", vz_ref)
    r3 = q3.Execute().Select()

    if r3.Next():
        or_ref = r3.ОР
        if conn.ValueIsFilled(or_ref):
            dog = or_ref.Договор
            if conn.ValueIsFilled(dog):
                kred_contract = s(dog.Наименование)
                podr = dog.А_ПодразделениеОказаниеУслугПодразделению
                kred_podr_ok = conn.ValueIsFilled(podr)

    if deb_podr_ok and kred_podr_ok:
        ok_count += 1
    else:
        bad_count += 1
        source = s(res.Осн)
        print(f"BAD: {num} sum={res.С}")
        print(f"  Source: {source}")
        print(f"  Deb contract: {deb_contract} (А_Подр filled={deb_podr_ok})")
        print(f"  Kred contract: {kred_contract} (А_Подр filled={kred_podr_ok})")
        print()

print(f"\nTotal: OK={ok_count}, BAD={bad_count}")
