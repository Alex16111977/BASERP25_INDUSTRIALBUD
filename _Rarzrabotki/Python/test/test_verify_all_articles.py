# -*- coding: utf-8 -*-
"""
Proverka opredeleniya statej disbalansa dlya vsekh 6 keysov.
Emuliruem logiku ObjectModule.bsl: 2 otdelnyh zaprosa (deb/kred).
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

QUERY_DISBALANS = """ВЫБРАТЬ
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

QUERY_ARTICLE = """ВЫБРАТЬ ПЕРВЫЕ 1
    АП.Статья КАК Статья,
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
ГДЕ АП.Регистратор = &Регистратор
    И АП.Подразделение = &Подразделение
    И НЕ АП.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
СГРУППИРОВАТЬ ПО АП.Статья
УПОРЯДОЧИТЬ ПО СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
    ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) {ORDER}"""


def find_doc(conn, doc_type, number):
    q = conn.NewObject("Query")
    q.Text = (f"ВЫБРАТЬ Док.Ссылка ИЗ Документ.{doc_type} КАК Док"
              f" ГДЕ Док.Номер ПОДОБНО &Н И Док.Дата МЕЖДУ &Нач И &Кон")
    q.SetParameter("Н", f"%{number}%")
    q.SetParameter("Нач", date_from)
    q.SetParameter("Кон", date_to)
    sel = q.Execute().Choose()
    return sel.Ссылка if sel.Next() else None


def get_article(conn, dok_ref, podrazd, order):
    q = conn.NewObject("Query")
    q.Text = QUERY_ARTICLE.replace("{ORDER}", order)
    q.SetParameter("Регистратор", dok_ref)
    q.SetParameter("Подразделение", podrazd)
    sel = q.Execute().Choose()
    if sel.Next():
        sn = sel.Статья.Наименование if conn.ЗначениеЗаполнено(sel.Статья) else "?"
        return sn, float(sel.Контроль)
    return "NE OPREDELENA", 0


def test_case(conn, label, doc_type, number, expected_deb_art, expected_kred_art):
    print(f"  {label}")
    dok = find_doc(conn, doc_type, number)
    if not dok:
        print(f"    NE NAJDEN!\n")
        return

    q = conn.NewObject("Query")
    q.Text = QUERY_DISBALANS
    q.SetParameter("Рег", dok)
    sel = q.Execute().Choose()

    podrazd_deb = podrazd_kred = None
    while sel.Next():
        k = float(sel.Контроль)
        if k > 0:
            podrazd_deb = sel.Подразделение
        elif k < 0:
            podrazd_kred = sel.Подразделение

    if not podrazd_deb or not podrazd_kred:
        print(f"    < 2 podrazd\n")
        return

    deb_name = podrazd_deb.Наименование
    kred_name = podrazd_kred.Наименование

    art_deb, k_deb = get_article(conn, dok, podrazd_deb, "УБЫВ")
    art_kred, k_kred = get_article(conn, dok, podrazd_kred, "ВОЗР")

    result = f"{art_deb} - {art_kred}"
    ok_deb = expected_deb_art in art_deb
    ok_kred = expected_kred_art in art_kred
    status = "OK" if (ok_deb and ok_kred) else "FAIL"

    print(f"    Deb: {deb_name:20s} art={art_deb:40s} K={k_deb:+.2f}")
    print(f"    Kred: {kred_name:20s} art={art_kred:40s} K={k_kred:+.2f}")
    print(f"    Rezultat: {result}")
    print(f"    [{status}] Ozhidali: {expected_deb_art} - {expected_kred_art}\n")


print("=" * 70)
print("  PROVERKA VSEKH 6 KEYSOV")
print("=" * 70 + "\n")

test_case(conn, "Kejs 1: PeremTovarov IB00-001591",
          "ПеремещениеТоваров", "001591",
          "Товары на оптовых складах", "Товары на оптовых складах")

test_case(conn, "Kejs 2: SpisNedostach IB00-000752",
          "СписаниеНедостачТоваров", "000752",
          "Расходы текущего периода", "Товары на оптовых складах")

test_case(conn, "Kejs 3: PriobretenieTU IB00-007779",
          "ПриобретениеТоваровУслуг", "007779",
          "Товары на оптовых складах", "Выданные авансы")

test_case(conn, "Kejs 4: SpisBeznal 00000019249",
          "СписаниеБезналичныхДенежныхСредств", "019249",
          "Выданные авансы", "Денежные средства")

test_case(conn, "Kejs 5: PeremTovarov IB00-001666",
          "ПеремещениеТоваров", "001666",
          "Товары на оптовых складах", "Товары на оптовых складах")

test_case(conn, "Kejs 6: PriobretenieTU IB00-007107",
          "ПриобретениеТоваровУслуг", "007107",
          "Товары на оптовых складах", "Задолженность перед поставщиками")

# Dopolnitelno: kejs 000C-000003 (SpisNedostach IB00-000755)
test_case(conn, "Kejs 7: SpisNedostach IB00-000755 (bug 000C-003)",
          "СписаниеНедостачТоваров", "000755",
          "Расходы текущего периода", "Товары на оптовых складах")

print("=== GOTOVO ===")
