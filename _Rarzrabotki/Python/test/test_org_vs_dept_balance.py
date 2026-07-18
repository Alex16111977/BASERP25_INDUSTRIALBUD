# -*- coding: utf-8 -*-
"""
Proverka: balans po organizacii shoditsya, a po podrazdeleniyam — net.
Eto uslovie dlya sozdaniya Vzaimozacheta.
"""
import win32com.client, pythoncom
from datetime import datetime

CONN_STR = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK\n")

date_from = datetime(2025, 12, 1)
date_to = datetime(2025, 12, 31, 23, 59, 59)

# Zapros balansa po organizacii (bez razreza podrazdelenij)
QUERY_ORG = """ВЫБРАТЬ
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Дебет,
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
        ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Кредит,
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
ГДЕ АП.Регистратор = &Рег
    И НЕ АП.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"""

# Zapros balansa po podrazdeleniyam
QUERY_DEPT = """ВЫБРАТЬ
    АП.Подразделение КАК Подразделение,
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
ГДЕ АП.Регистратор = &Рег
    И НЕ АП.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
СГРУППИРОВАТЬ ПО АП.Подразделение"""


def find_doc(conn, doc_type, number):
    q = conn.NewObject("Query")
    q.Text = (f"ВЫБРАТЬ Док.Ссылка ИЗ Документ.{doc_type} КАК Док"
              f" ГДЕ Док.Номер ПОДОБНО &Н И Док.Дата МЕЖДУ &Нач И &Кон")
    q.SetParameter("Н", f"%{number}%")
    q.SetParameter("Нач", date_from)
    q.SetParameter("Кон", date_to)
    sel = q.Execute().Choose()
    return sel.Ссылка if sel.Next() else None


def check_doc(conn, label, doc_type, number):
    print(f"  {label}")
    dok = find_doc(conn, doc_type, number)
    if not dok:
        print(f"    NE NAJDEN!\n")
        return

    # 1) Balans po organizacii
    q1 = conn.NewObject("Query")
    q1.Text = QUERY_ORG
    q1.SetParameter("Рег", dok)
    sel1 = q1.Execute().Choose()
    sel1.Next()
    org_db = float(sel1.Дебет)
    org_kr = float(sel1.Кредит)
    org_k = float(sel1.Контроль)
    org_ok = abs(org_k) < 0.01

    print(f"    Org:  Db={org_db:12.2f}  Kr={org_kr:12.2f}  K={org_k:+.2f}  {'SHODITSYA' if org_ok else 'NE SHODITSYA!'}")

    # 2) Balans po podrazdeleniyam
    q2 = conn.NewObject("Query")
    q2.Text = QUERY_DEPT
    q2.SetParameter("Рег", dok)
    sel2 = q2.Execute().Choose()
    dept_disbalance = False
    while sel2.Next():
        dn = sel2.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel2.Подразделение) else "?"
        k = float(sel2.Контроль)
        if abs(k) > 0.01:
            dept_disbalance = True
            print(f"    Dept: {dn:25s}  K={k:+.2f}")

    if org_ok and dept_disbalance:
        print(f"    --> VZAIMOZACHET NUZHEN (org OK, dept NE OK)")
    elif not org_ok:
        print(f"    --> VZAIMOZACHET NE NUZHEN (org NE shoditsya!)")
    else:
        print(f"    --> VZAIMOZACHET NE NUZHEN (dept OK)")
    print()


print("=" * 70)
print("  ORG vs DEPT BALANCE CHECK")
print("=" * 70 + "\n")

check_doc(conn, "Kejs 1: PeremTovarov IB00-001591", "ПеремещениеТоваров", "001591")
check_doc(conn, "Kejs 2: SpisNedostach IB00-000752", "СписаниеНедостачТоваров", "000752")
check_doc(conn, "Kejs 3: PriobretenieTU IB00-007779", "ПриобретениеТоваровУслуг", "007779")
check_doc(conn, "Kejs 4: SpisBeznal 00000019249", "СписаниеБезналичныхДенежныхСредств", "019249")
check_doc(conn, "Kejs 5: PeremTovarov IB00-001666", "ПеремещениеТоваров", "001666")
check_doc(conn, "Kejs 6: PriobretenieTU IB00-007107", "ПриобретениеТоваровУслуг", "007107")
check_doc(conn, "Kejs 7: SpisNedostach IB00-000755", "СписаниеНедостачТоваров", "000755")

print("=== GOTOVO ===")
