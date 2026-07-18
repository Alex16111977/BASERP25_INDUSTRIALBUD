# -*- coding: utf-8 -*-
"""
Test v3: Sozdanie VzaimozachetZadolzhennosti s pravilnym ObjektomRaschetov
Dokument-osnova: SpisanieNedostachTovarov IB00-000752 ot 31.12.2025
"""
import sys
import win32com.client
import pythoncom
from datetime import datetime

CONN_STR = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK")

date_from = datetime(2025, 12, 1)
date_to = datetime(2025, 12, 31, 23, 59, 59)

# 1. Dokument-osnova
print("\n=== 1. Dokument ===")
q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ Док.Ссылка ИЗ Документ.СписаниеНедостачТоваров КАК Док
ГДЕ Док.Номер ПОДОБНО &Н И Док.Дата МЕЖДУ &Нач И &Кон"""
q.SetParameter("Н", "%000752%")
q.SetParameter("Нач", date_from)
q.SetParameter("Кон", date_to)
sel = q.Execute().Choose()
sel.Next()
dok_osnov = sel.Ссылка
print(f"  {dok_osnov}")

# 2. Disbalans
print("\n=== 2. Disbalans ===")
q2 = conn.NewObject("Query")
q2.Text = """ВЫБРАТЬ
    АП.Подразделение, АП.Организация,
    СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
ГДЕ АП.Регистратор = &Рег
    И НЕ АП.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Организация
ИМЕЮЩИЕ СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ)
    <> СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
        ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ)"""
q2.SetParameter("Рег", dok_osnov)
sel2 = q2.Execute().Choose()

podrazd_deb = podrazd_kred = org_dok = None
summa = 0
while sel2.Next():
    k = float(sel2.Контроль)
    dept = sel2.Подразделение
    dn = dept.Наименование if conn.ЗначениеЗаполнено(dept) else "?"
    print(f"  {dn}: {k:.2f}")
    if k > 0:
        podrazd_deb = dept; summa = k; org_dok = sel2.Организация
    elif k < 0:
        podrazd_kred = dept

if not podrazd_deb or not podrazd_kred:
    print("  < 2 podrazdelenij"); sys.exit()

# 3. Partnery / Kontragenty
print("\n=== 3. Partnery ===")
def find_partner(conn, podrazd):
    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка ИЗ Справочник.Партнеры КАК П
    ГДЕ П.А_Подразделение = &П И НЕ П.ПометкаУдаления"""
    q.SetParameter("П", podrazd)
    sel = q.Execute().Choose()
    if sel.Next(): return sel.Ссылка
    q2 = conn.NewObject("Query")
    q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка ИЗ Справочник.Партнеры КАК П
    ГДЕ П.Наименование = &Н И НЕ П.ПометкаУдаления"""
    q2.SetParameter("Н", str(podrazd))
    sel2 = q2.Execute().Choose()
    return sel2.Ссылка if sel2.Next() else None

def find_kontr(conn, partner):
    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 К.Ссылка ИЗ Справочник.Контрагенты КАК К
    ГДЕ К.Партнер = &П И НЕ К.ПометкаУдаления"""
    q.SetParameter("П", partner)
    sel = q.Execute().Choose()
    return sel.Ссылка if sel.Next() else None

partner_deb = find_partner(conn, podrazd_deb)
kontr_deb = find_kontr(conn, partner_deb)
partner_kred = find_partner(conn, podrazd_kred)
kontr_kred = find_kontr(conn, partner_kred)
print(f"  Deb: {kontr_deb.Наименование}, Kred: {kontr_kred.Наименование}")

# 4. Dogovory
print("\n=== 4. Dogovory ===")
def find_dogovor(conn, kontr, tip_name):
    q = conn.NewObject("Query")
    q.Text = ("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Справочник.ДоговорыКонтрагентов КАК Д"
        " ГДЕ Д.Контрагент = &К И Д.ТипДоговора = ЗНАЧЕНИЕ(Перечисление.ТипыДоговоров." + tip_name + ")"
        " И НЕ Д.ПометкаУдаления")
    q.SetParameter("К", kontr)
    sel = q.Execute().Choose()
    return sel.Ссылка if sel.Next() else None

dog_deb = find_dogovor(conn, kontr_deb, "СПоставщиком")
dog_kred = find_dogovor(conn, kontr_kred, "СПокупателем")
print(f"  Dog deb: {dog_deb.Наименование if conn.ЗначениеЗаполнено(dog_deb) else 'NE NAJDEN'}")
print(f"  Dog kred: {dog_kred.Наименование if conn.ЗначениеЗаполнено(dog_kred) else 'NE NAJDEN'}")

# 5. ObjektyRaschetov — KLYUCHEVOJ MOMENT
print("\n=== 5. ObjektyRaschetov ===")
def find_obj_raschetov(conn, kontr, tip_rasch_name, dog):
    q = conn.NewObject("Query")
    q.Text = ("ВЫБРАТЬ ПЕРВЫЕ 1 ОР.Ссылка ИЗ Справочник.ОбъектыРасчетов КАК ОР"
        " ГДЕ ОР.Контрагент = &К"
        " И ОР.ТипРасчетов = ЗНАЧЕНИЕ(Перечисление.ТипыРасчетовСПартнерами." + tip_rasch_name + ")"
        " И ОР.Договор = &Д"
        " И НЕ ОР.ПометкаУдаления")
    q.SetParameter("К", kontr)
    q.SetParameter("Д", dog)
    sel = q.Execute().Choose()
    return sel.Ссылка if sel.Next() else None

obj_raschetov_deb = find_obj_raschetov(conn, kontr_deb, "РасчетыСПоставщиком", dog_deb)
obj_raschetov_kred = find_obj_raschetov(conn, kontr_kred, "РасчетыСКлиентом", dog_kred)

if conn.ЗначениеЗаполнено(obj_raschetov_deb):
    print(f"  ObjRasch deb: {obj_raschetov_deb}")
else:
    print("  ObjRasch deb NE NAJDEN!")
if conn.ЗначениеЗаполнено(obj_raschetov_kred):
    print(f"  ObjRasch kred: {obj_raschetov_kred}")
else:
    print("  ObjRasch kred NE NAJDEN!")

# 6. Najti/obnovit' dokument
print("\n=== 6. Dokument ===")
q6 = conn.NewObject("Query")
q6.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка ИЗ Документ.ВзаимозачетЗадолженности КАК Док
ГДЕ Док.А_ДокументОснованиеДляБаланса = &Д И НЕ Док.ПометкаУдаления"""
q6.SetParameter("Д", dok_osnov)
sel6 = q6.Execute().Choose()

if sel6.Next():
    dok_obj = sel6.Ссылка.ПолучитьОбъект()
    if dok_obj.Проведен:
        dok_obj.Записать(conn.РежимЗаписиДокумента.ОтменаПроведения)
    dok_obj.ДебиторскаяЗадолженность.Очистить()
    dok_obj.КредиторскаяЗадолженность.Очистить()
    print(f"  Obnovlyaem: {sel6.Ссылка}")
else:
    dok_obj = conn.Документы.ВзаимозачетЗадолженности.СоздатьДокумент()
    print("  Novyj dokument")

# 7. Zapolnenie
print("\n=== 7. Zapolnenie ===")
dok_obj.Дата = dok_osnov.Дата
dok_obj.Организация = org_dok
dok_obj.ВидОперации = getattr(conn.Перечисления.ВидыОперацийВзаимозачетаЗадолженности, "Произвольный")
dok_obj.КонтрагентДебитор = kontr_deb
dok_obj.КонтрагентКредитор = kontr_kred
dok_obj.ТипДебитора = getattr(conn.Перечисления.ТипыУчастниковВзаимозачета, "Поставщик")
dok_obj.ТипКредитора = getattr(conn.Перечисления.ТипыУчастниковВзаимозачета, "Клиент")
dok_obj.СуммаРегл = summa
dok_obj.СуммаУпр = summa
dok_obj.А_ДокументОснованиеДляБаланса = dok_osnov
dok_obj.А_ВведенВЕРП = True

valuta = conn.Справочники.Валюты.НайтиПоКоду("980")

# DebitorskayaZadolzhennost
s = dok_obj.ДебиторскаяЗадолженность.Добавить()
s.Партнер = partner_deb
s.ТипРасчетов = getattr(conn.Перечисления.ТипыРасчетовСПартнерами, "РасчетыСПоставщиком")
s.СуммаВзаиморасчетов = summa
s.СуммаРегл = summa
s.СуммаУпр = summa
s.ВалютаВзаиморасчетов = valuta
s.Организация = org_dok
if conn.ЗначениеЗаполнено(obj_raschetov_deb):
    s.ОбъектРасчетов = obj_raschetov_deb
    print(f"  Deb ObjRasch = {obj_raschetov_deb}")
else:
    print("  Deb ObjRasch = pustoj!")

# KreditorskayaZadolzhennost
s2 = dok_obj.КредиторскаяЗадолженность.Добавить()
s2.Партнер = partner_kred
s2.ТипРасчетов = getattr(conn.Перечисления.ТипыРасчетовСПартнерами, "РасчетыСКлиентом")
s2.СуммаВзаиморасчетов = summa
s2.СуммаРегл = summa
s2.СуммаУпр = summa
s2.ВалютаВзаиморасчетов = valuta
s2.Организация = org_dok
if conn.ЗначениеЗаполнено(obj_raschetov_kred):
    s2.ОбъектРасчетов = obj_raschetov_kred
    print(f"  Kred ObjRasch = {obj_raschetov_kred}")
else:
    print("  Kred ObjRasch = pustoj!")

print(f"  Summa: {summa:.2f}")

# 8. Provedenie
print("\n=== 8. Provedenie ===")
try:
    dok_obj.Записать(conn.РежимЗаписиДокумента.Проведение)
    print(f"  OK proveden: {dok_obj.Ссылка}")
except Exception as e:
    print(f"  OSHIBKA: {e}")
    try:
        dok_obj.Записать(conn.РежимЗаписиДокумента.Запись)
        print(f"  Zapisan bez provedeniya: {dok_obj.Ссылка}")
    except Exception as e2:
        print(f"  OSHIBKA zapisi: {e2}")
        sys.exit()

# 9. Proverka dvizhenij
print("\n=== 9. Dvizheniya ===")
for reg in ["ПрочиеАктивыПассивы", "РасчетыСКлиентами", "РасчетыСПоставщиками",
            "РасчетыСКлиентамиПоДокументам", "ДвиженияКонтрагентКонтрагент"]:
    try:
        qr = conn.NewObject("Query")
        qr.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Р"
        qr.SetParameter("Р", dok_obj.Ссылка)
        sr = qr.Execute().Choose()
        sr.Next()
        cnt = int(sr.К)
        if cnt > 0:
            print(f"  {reg}: {cnt}")
    except:
        pass

print("\nGotovo!")
