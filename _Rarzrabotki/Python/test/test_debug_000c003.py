# -*- coding: utf-8 -*-
"""Debug: pochemu 000C-000003 poluchil nepravilnuyu statyu"""
import win32com.client, pythoncom
from datetime import datetime

CONN_STR = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK\n")

# 1. Najti dokument-osnovu SpisanieNedostach IB00-000755
q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ Док.Ссылка ИЗ Документ.СписаниеНедостачТоваров КАК Док
ГДЕ Док.Номер ПОДОБНО &Н И Док.Дата МЕЖДУ &Нач И &Кон"""
q.SetParameter("Н", "%000755%")
q.SetParameter("Нач", datetime(2025, 12, 1))
q.SetParameter("Кон", datetime(2025, 12, 31, 23, 59, 59))
sel = q.Execute().Choose()
sel.Next()
dok = sel.Ссылка
print(f"Dok-osnova: {dok}")

# 2. VSE stat'i iz registra po etomu dokumentu
print("\n=== VSE stat'i iz registra ===")
q2 = conn.NewObject("Query")
q2.Text = """ВЫБРАТЬ
    АП.Подразделение КАК Подразделение,
    АП.Статья КАК Статья,
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
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Статья"""
q2.SetParameter("Рег", dok)
sel2 = q2.Execute().Choose()

while sel2.Next():
    dn = sel2.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel2.Подразделение) else "?"
    sn = sel2.Статья.Наименование if conn.ЗначениеЗаполнено(sel2.Статья) else "?"
    db = float(sel2.Дебет)
    kr = float(sel2.Кредит)
    k = float(sel2.Контроль)
    print(f"  {dn:25s} | {sn:40s} | Db={db:12.2f} | Kr={kr:12.2f} | K={k:+12.2f}")

# 3. Disbalans po podrazdeleniyam (krok 1 algoritma)
print("\n=== Disbalans po podrazdeleniyam ===")
q3 = conn.NewObject("Query")
q3.Text = """ВЫБРАТЬ
    АП.Подразделение КАК Подразделение,
    АП.Организация КАК Организация,
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
ГДЕ АП.Регистратор = &Рег
    И НЕ АП.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Организация"""
q3.SetParameter("Рег", dok)
sel3 = q3.Execute().Choose()

podrazd_deb = podrazd_kred = None
while sel3.Next():
    dn = sel3.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel3.Подразделение) else "?"
    k = float(sel3.Контроль)
    role = "DEBITOR" if k > 0 else ("KREDITOR" if k < 0 else "ZERO")
    print(f"  {dn}: {k:+.2f} [{role}]")
    if k > 0:
        podrazd_deb = sel3.Подразделение
    elif k < 0:
        podrazd_kred = sel3.Подразделение

if podrazd_deb and podrazd_kred:
    print(f"\n  Debitor: {podrazd_deb.Наименование}")
    print(f"  Kreditor: {podrazd_kred.Наименование}")

    # 4. Emulyaciya kroka 5.5 — opredelenie statej
    print("\n=== Krok 5.5: opredelenie statej (kak v ObjectModule) ===")
    q4 = conn.NewObject("Query")
    q4.Text = """ВЫБРАТЬ
        АП.Подразделение КАК Подразделение,
        АП.Статья КАК Статья,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
    ГДЕ АП.Регистратор = &Регистратор
        И НЕ АП.Статья В (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
        И АП.Подразделение В (&ПодразделениеДебитор, &ПодразделениеКредитор)
    СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Статья"""
    q4.SetParameter("Регистратор", dok)
    q4.SetParameter("ПодразделениеДебитор", podrazd_deb)
    q4.SetParameter("ПодразделениеКредитор", podrazd_kred)
    sel4 = q4.Execute().Choose()

    stat_deb = None
    stat_kred = None
    max_k_deb = 0
    max_k_kred = 0

    while sel4.Next():
        dn = sel4.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel4.Подразделение) else "?"
        sn = sel4.Статья.Наименование if conn.ЗначениеЗаполнено(sel4.Статья) else "?"
        k = float(sel4.Контроль)
        is_deb = (sel4.Подразделение == podrazd_deb)
        is_kred = (sel4.Подразделение == podrazd_kred)
        marker = ""
        if is_deb and k > max_k_deb:
            stat_deb = sel4.Статья
            max_k_deb = k
            marker = " <-- MAX DEB"
        if is_kred and (-k) > max_k_kred:
            stat_kred = sel4.Статья
            max_k_kred = -k
            marker = " <-- MAX KRED"
        print(f"  {dn:25s} | {sn:40s} | K={k:+12.2f}{marker}")

    print(f"\n  REZULTAT:")
    if stat_deb:
        print(f"    StatyaDebitora = {stat_deb.Наименование} (max kontrol = {max_k_deb:.2f})")
    else:
        print(f"    StatyaDebitora = NE OPREDELENA!")
    if stat_kred:
        print(f"    StatyaKreditora = {stat_kred.Наименование} (max |kontrol| = {max_k_kred:.2f})")
    else:
        print(f"    StatyaKreditora = NE OPREDELENA!")

    if stat_deb and stat_kred:
        expected = stat_deb.Наименование + " - " + stat_kred.Наименование
        print(f"\n    Ozhidaemoe nazvanie: '{expected}'")
        print(f"    Fakticheskoe:        'Tovary na optovyh skladah - Tovary na optovyh skladah'")
        if "Raskhody" in expected or "asxod" in expected:
            print("    BUG: statya debitora dolzhna byt' Raskhody, no algoritm vybral Tovary!")

print("\nGotovo")
