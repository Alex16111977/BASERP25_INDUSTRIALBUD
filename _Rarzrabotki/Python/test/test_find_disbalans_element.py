# -*- coding: utf-8 -*-
"""
Test: poisk elementa A_DisbalansPoPodr po ТЧ СтатьиБаланса
Kejs 1: PeremeschenieTovarov IB00-001666 -> stati = [Tovary, Tovary] (2 odinakov.)
Kejs 2: SpisanieNedostach IB00-000752 -> stati = [Raskhody, Tovary] (2 raznyh)
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


def get_articles_by_dept(conn, dok_ref):
    """Poluchit' stat'i po podrazdeleniyam (ne RAZLICHNYE — kazhdyj podrazd otdel'no)"""
    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ
        АП.Подразделение КАК Подразделение,
        АП.Статья КАК Статья,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
    ГДЕ АП.Регистратор = &Рег
        И НЕ АП.Статья В (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
    СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Статья
    УПОРЯДОЧИТЬ ПО АП.Подразделение"""
    q.SetParameter("Рег", dok_ref)
    sel = q.Execute().Choose()
    results = []
    while sel.Next():
        dn = sel.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel.Подразделение) else "?"
        sn = sel.Статья.Наименование if conn.ЗначениеЗаполнено(sel.Статья) else "?"
        k = float(sel.Контроль)
        results.append({'dept': dn, 'dept_ref': sel.Подразделение,
                        'article': sn, 'article_ref': sel.Статья, 'kontrol': k})
    return results


def find_disbalans_element(conn, article_refs):
    """Najti element spravochnika A_DisbalansPoPodr po naboru statej v TCh"""
    # Sozdaem SpisokZnachenij iz statej
    val_list = conn.NewObject("СписокЗначений")
    for art in article_refs:
        val_list.Добавить(art)

    kol_statej = len(article_refs)

    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ
        ТЧ.Ссылка КАК Элемент,
        ТЧ.Ссылка.Наименование КАК Наименование,
        КОЛИЧЕСТВО(ТЧ.НомерСтроки) КАК КолСтрок,
        СУММА(ВЫБОР КОГДА ТЧ.Статья В (&СписокСтатей) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Совпадений
    ИЗ Справочник.А_ДисбалансПоПодразделению.СтатьиБаланса КАК ТЧ
    ГДЕ НЕ ТЧ.Ссылка.ПометкаУдаления
    СГРУППИРОВАТЬ ПО ТЧ.Ссылка"""
    q.SetParameter("СписокСтатей", val_list)

    sel = q.Execute().Choose()
    while sel.Next():
        kol_strok = int(sel.КолСтрок)
        sovpadenij = int(sel.Совпадений)
        # Tochnoe sovpadenie: kol strok = kol nashikh statej I vse sovpali
        if kol_strok == kol_statej and sovpadenij == kol_statej:
            return sel.Элемент
    return None


def find_doc(conn, doc_type, number):
    q = conn.NewObject("Query")
    q.Text = (f"ВЫБРАТЬ Док.Ссылка ИЗ Документ.{doc_type} КАК Док"
              f" ГДЕ Док.Номер ПОДОБНО &Н И Док.Дата МЕЖДУ &Нач И &Кон")
    q.SetParameter("Н", f"%{number}%")
    q.SetParameter("Нач", date_from)
    q.SetParameter("Кон", date_to)
    sel = q.Execute().Choose()
    return sel.Ссылка if sel.Next() else None


# === KEJS 1: PeremeschenieTovarov IB00-001666 ===
print("=" * 60)
print("KEJS 1: PeremeschenieTovarov IB00-001666")
print("  Ozhidaem: 'Tovary na optovyh skladah - Tovary na optovyh skladah'")
print("=" * 60)

dok1 = find_doc(conn, "ПеремещениеТоваров", "001666")
arts1 = get_articles_by_dept(conn, dok1)
print(f"\n  Stat'i iz registra:")
for a in arts1:
    print(f"    {a['dept']:25s} | {a['article']:40s} | {a['kontrol']:+.2f}")

# Sobirayem stat'i: po odnoy ot kazhdogo podrazdeleniya (glavnaya statya)
# Debitor = max kontrol, Kreditor = min kontrol
debitor_arts = [a for a in arts1 if a['kontrol'] > 0]
kreditor_arts = [a for a in arts1 if a['kontrol'] < 0]

# Berem stat'u s max |kontrol| dlya kazhdoj storony
deb_art = max(debitor_arts, key=lambda x: abs(x['kontrol']))['article_ref'] if debitor_arts else None
kred_art = max(kreditor_arts, key=lambda x: abs(x['kontrol']))['article_ref'] if kreditor_arts else None

if deb_art and kred_art:
    article_refs = [deb_art, kred_art]
    print(f"\n  Stat'i dlya poiska: [{deb_art.Наименование}, {kred_art.Наименование}]")

    found = find_disbalans_element(conn, article_refs)
    if conn.ЗначениеЗаполнено(found):
        print(f"  NAJDEN: {found.Наименование}")
    else:
        print("  NE NAJDEN — nuzhno sozdat'")
else:
    print("  Nevozmozhno opredelit' storony")


# === KEJS 2: SpisanieNedostach IB00-000752 ===
print("\n" + "=" * 60)
print("KEJS 2: SpisanieNedostach IB00-000752")
print("  Ozhidaem: 'Raskhody tekushchego perioda - Tovary na optovyh skladah'")
print("=" * 60)

dok2 = find_doc(conn, "СписаниеНедостачТоваров", "000752")
arts2 = get_articles_by_dept(conn, dok2)
print(f"\n  Stat'i iz registra:")
for a in arts2:
    print(f"    {a['dept']:25s} | {a['article']:40s} | {a['kontrol']:+.2f}")

debitor_arts2 = [a for a in arts2 if a['kontrol'] > 0]
kreditor_arts2 = [a for a in arts2 if a['kontrol'] < 0]

deb_art2 = max(debitor_arts2, key=lambda x: abs(x['kontrol']))['article_ref'] if debitor_arts2 else None
kred_art2 = max(kreditor_arts2, key=lambda x: abs(x['kontrol']))['article_ref'] if kreditor_arts2 else None

if deb_art2 and kred_art2:
    article_refs2 = [deb_art2, kred_art2]
    print(f"\n  Stat'i dlya poiska: [{deb_art2.Наименование}, {kred_art2.Наименование}]")

    found2 = find_disbalans_element(conn, article_refs2)
    if conn.ЗначениеЗаполнено(found2):
        print(f"  NAJDEN: {found2.Наименование}")
    else:
        print("  NE NAJDEN — nuzhno sozdat'")


# === KEJS 3: PriobretenieTU IB00-007779 (3 stat'i) ===
print("\n" + "=" * 60)
print("KEJS 3: PriobretenieTU IB00-007779 (vozmozhno 3 stat'i)")
print("=" * 60)

dok3 = find_doc(conn, "ПриобретениеТоваровУслуг", "007779")
if dok3:
    arts3 = get_articles_by_dept(conn, dok3)
    print(f"\n  Stat'i iz registra:")
    for a in arts3:
        print(f"    {a['dept']:25s} | {a['article']:40s} | {a['kontrol']:+.2f}")

    # Vse unikalnye stat'i
    unique_arts = []
    seen = set()
    for a in arts3:
        art_name = a['article']
        if art_name not in seen:
            seen.add(art_name)
            unique_arts.append(a['article_ref'])
    print(f"\n  Unikalnye stat'i: {[a.Наименование for a in unique_arts]}")

    # Po podrazdeleniyam
    debitor_arts3 = [a for a in arts3 if a['kontrol'] > 0]
    kreditor_arts3 = [a for a in arts3 if a['kontrol'] < 0]
    if debitor_arts3:
        deb3 = max(debitor_arts3, key=lambda x: abs(x['kontrol']))
        print(f"  Debitor stat'ya: {deb3['article']} ({deb3['kontrol']:+.2f})")
    if kreditor_arts3:
        kred3 = max(kreditor_arts3, key=lambda x: abs(x['kontrol']))
        print(f"  Kreditor stat'ya: {kred3['article']} ({kred3['kontrol']:+.2f})")


print("\n=== GOTOVO ===")
