# -*- coding: utf-8 -*-
"""
Analiz vsekh keysov disbalansa:
1. PeremeschenieTovarov (2 podrazd) - etalon
2. SpisanieNedostachTovarov (2 podrazd: dept + CO)
3. PriobretenieTU (Logistika - otritsatelnyj kontrol)
4. SpisanieBeznalichnyhDS (Logistika - polozhitelnyj kontrol)

Dlya kazhdogo: kakie stat'i, kakie podrazdeleniya, znaki
"""
import win32com.client, pythoncom
from datetime import datetime

CONN_STR = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK\n")

date_from = datetime(2025, 12, 1)
date_to = datetime(2025, 12, 31, 23, 59, 59)


def analyze_doc(conn, doc_type, number, label):
    print("=" * 70)
    print(f"  {label}")
    print("=" * 70)

    q = conn.NewObject("Query")
    q.Text = (f"ВЫБРАТЬ Док.Ссылка ИЗ Документ.{doc_type} КАК Док"
              f" ГДЕ Док.Номер ПОДОБНО &Н И Док.Дата МЕЖДУ &Нач И &Кон")
    q.SetParameter("Н", f"%{number}%")
    q.SetParameter("Нач", date_from)
    q.SetParameter("Кон", date_to)
    sel = q.Execute().Choose()
    if not sel.Next():
        print("  NE NAJDEN!\n")
        return
    dok = sel.Ссылка
    print(f"  Dok: {dok}\n")

    # Disbalans po podrazdeleniyam
    q2 = conn.NewObject("Query")
    q2.Text = """ВЫБРАТЬ
        АП.Подразделение КАК Подразделение,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
    ГДЕ АП.Регистратор = &Рег
        И НЕ АП.Статья В (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
    СГРУППИРОВАТЬ ПО АП.Подразделение
    ИМЕЮЩИЕ СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ)
        <> СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ)"""
    q2.SetParameter("Рег", dok)
    sel2 = q2.Execute().Choose()

    print("  Disbalans po podrazdeleniyam:")
    while sel2.Next():
        dept = sel2.Подразделение
        dn = dept.Наименование if conn.ЗначениеЗаполнено(dept) else "?"
        k = float(sel2.Контроль)
        role = "DEBITOR (+)" if k > 0 else "KREDITOR (-)"
        print(f"    {dn}: {k:+.2f}  [{role}]")

    # Detali po stat'yam i podrazdeleniyam
    q3 = conn.NewObject("Query")
    q3.Text = """ВЫБРАТЬ
        АП.Подразделение КАК Подразделение,
        АП.Статья КАК Статья,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Дебет,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Кредит
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
    ГДЕ АП.Регистратор = &Рег
        И НЕ АП.Статья В (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
    СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Статья"""
    q3.SetParameter("Рег", dok)
    sel3 = q3.Execute().Choose()

    print("\n  Detali po stat'yam:")
    while sel3.Next():
        dept = sel3.Подразделение
        dn = dept.Наименование if conn.ЗначениеЗаполнено(dept) else "?"
        sn = sel3.Статья.Наименование if conn.ЗначениеЗаполнено(sel3.Статья) else "?"
        db = float(sel3.Дебет)
        kr = float(sel3.Кредит)
        print(f"    {dn:25s} | {sn:40s} | Db={db:12.2f} | Kr={kr:12.2f}")

    print()


# === Keysy ===

# 1. PeremeschenieTovarov - etalon
analyze_doc(conn, "ПеремещениеТоваров", "001591",
    "KEYS 1: PeremeschenieTovarov IB00-001591 (KMD-2 -> Globino-2)")

# 2. SpisanieNedostach
analyze_doc(conn, "СписаниеНедостачТоваров", "000752",
    "KEYS 2: SpisanieNedostach IB00-000752 (Astarta.Tishchenki)")

# 3. PriobretenieTU s otricatelnym kontrolem (Logistika)
analyze_doc(conn, "ПриобретениеТоваровУслуг", "007779",
    "KEYS 3: PriobretenieTU IB00-007779 (Logistika, Kontrol<0)")

# 4. SpisanieBeznalichnyhDS (Logistika)
analyze_doc(conn, "СписаниеБезналичныхДенежныхСредств", "019249",
    "KEYS 4: SpisanieBeznalichnyhDS 00000019249 (Logistika)")

# 5. Esche odin PeremeschenieTovarov
analyze_doc(conn, "ПеремещениеТоваров", "001666",
    "KEYS 5: PeremeschenieTovarov IB00-001666 (Globino-2)")

# 6. PriobretenieTU s polozh. kontrolem (MD VOOZ EMS)
analyze_doc(conn, "ПриобретениеТоваровУслуг", "007107",
    "KEYS 6: PriobretenieTU IB00-007107 (MD VOOZ EMS, Kontrol>0)")

print("=== GOTOVO ===")
