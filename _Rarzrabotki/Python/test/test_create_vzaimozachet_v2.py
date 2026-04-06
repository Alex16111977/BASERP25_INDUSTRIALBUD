# -*- coding: utf-8 -*-
"""
Test v2: Sozdanie VzaimozachetZadolzhennosti s pravilnymi TipDebitora/TipKreditora
Dokument-osnova: SpisanieNedostachTovarov IB00-000752 ot 31.12.2025
Obnovlyaem uzhe sozdannyj 000C-000002
"""

import sys
import win32com.client
import pythoncom
from datetime import datetime

CONN_STR = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_STR)
    print("OK: Podklyuchenie k BaseERP")
    return conn


def najti_partnera(conn, podrazd):
    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка КАК Партнер
    ИЗ Справочник.Партнеры КАК П
    ГДЕ П.А_Подразделение = &Подразделение И НЕ П.ПометкаУдаления"""
    q.SetParameter("Подразделение", podrazd)
    res = q.Execute()
    sel = res.Choose()
    if sel.Next():
        return sel.Партнер
    # fallback
    q2 = conn.NewObject("Query")
    q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка КАК Партнер
    ИЗ Справочник.Партнеры КАК П
    ГДЕ П.Наименование = &Наименование И НЕ П.ПометкаУдаления"""
    q2.SetParameter("Наименование", str(podrazd))
    res2 = q2.Execute()
    sel2 = res2.Choose()
    if sel2.Next():
        return sel2.Партнер
    return None


def najti_kontragenta(conn, partner):
    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 К.Ссылка КАК Контрагент
    ИЗ Справочник.Контрагенты КАК К
    ГДЕ К.Партнер = &Партнер И НЕ К.ПометкаУдаления"""
    q.SetParameter("Партнер", partner)
    res = q.Execute()
    sel = res.Choose()
    if sel.Next():
        return sel.Контрагент
    return None


def najti_dogovor(conn, kontragent, tip_name):
    q = conn.NewObject("Query")
    q.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Договор"
        " ИЗ Справочник.ДоговорыКонтрагентов КАК Д"
        " ГДЕ Д.Контрагент = &Контрагент"
        "   И Д.ТипДоговора = ЗНАЧЕНИЕ(Перечисление.ТипыДоговоров." + tip_name + ")"
        "   И НЕ Д.ПометкаУдаления"
    )
    q.SetParameter("Контрагент", kontragent)
    res = q.Execute()
    sel = res.Choose()
    if sel.Next():
        return sel.Договор
    return None


def main():
    conn = connect()
    date_from = datetime(2025, 12, 1)
    date_to = datetime(2025, 12, 31, 23, 59, 59)

    # 1. Najti dokument-osnovu
    print("\n=== 1. Poisk dokumenta ===")
    q1 = conn.NewObject("Query")
    q1.Text = """ВЫБРАТЬ Док.Ссылка КАК Ссылка
    ИЗ Документ.СписаниеНедостачТоваров КАК Док
    ГДЕ Док.Номер ПОДОБНО &Номер И Док.Дата МЕЖДУ &Нач И &Кон"""
    q1.SetParameter("Номер", "%000752%")
    q1.SetParameter("Нач", date_from)
    q1.SetParameter("Кон", date_to)
    res1 = q1.Execute()
    sel1 = res1.Choose()
    if not sel1.Next():
        print("OSHIBKA: dokument ne najden!")
        return
    dok_osnov = sel1.Ссылка
    print(f"  Najdeno: {dok_osnov}")

    # 2. Disbalans
    print("\n=== 2. Disbalans ===")
    q2 = conn.NewObject("Query")
    q2.Text = """ВЫБРАТЬ
        АП.Подразделение КАК Подразделение,
        АП.Организация КАК Организация,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Дебет,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ) КАК Кредит,
        СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ -АП.Сумма КОНЕЦ) КАК Контроль
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
    ГДЕ АП.Регистратор = &Регистратор
        И НЕ АП.Статья В (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
    СГРУППИРОВАТЬ ПО АП.Подразделение, АП.Организация
    ИМЕЮЩИЕ СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ)
        <> СУММА(ВЫБОР КОГДА АП.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
            ТОГДА АП.Сумма ИНАЧЕ 0 КОНЕЦ)"""
    q2.SetParameter("Регистратор", dok_osnov)
    res2 = q2.Execute()
    sel2 = res2.Choose()

    podrazd_debitor = None
    podrazd_kreditor = None
    org_dok = None
    summa = 0

    while sel2.Next():
        dept = sel2.Подразделение
        dept_name = dept.Наименование if conn.ЗначениеЗаполнено(dept) else "<pusto>"
        kontrol = float(sel2.Контроль)
        role = "DEBITOR" if kontrol > 0 else "KREDITOR"
        print(f"  {dept_name}: Kontrol={kontrol:.2f} [{role}]")
        if kontrol > 0:
            podrazd_debitor = dept
            summa = kontrol
            org_dok = sel2.Организация
        elif kontrol < 0:
            podrazd_kreditor = dept

    if podrazd_debitor is None or podrazd_kreditor is None:
        print("  Menee 2 podrazdelenij")
        return

    # 3. Partnery / Kontragenty
    print("\n=== 3. Partnery/Kontragenty ===")
    partner_deb = najti_partnera(conn, podrazd_debitor)
    kontr_deb = najti_kontragenta(conn, partner_deb)
    partner_kred = najti_partnera(conn, podrazd_kreditor)
    kontr_kred = najti_kontragenta(conn, partner_kred)
    print(f"  Debitor: {kontr_deb.Наименование}, Kreditor: {kontr_kred.Наименование}")

    # 4. Dogovory
    print("\n=== 4. Dogovory ===")
    dogovor_deb = najti_dogovor(conn, kontr_deb, "СПоставщиком")
    dogovor_kred = najti_dogovor(conn, kontr_kred, "СПокупателем")
    print(f"  Dog deb: {dogovor_deb.Наименование if conn.ЗначениеЗаполнено(dogovor_deb) else 'NE NAJDEN'}")
    print(f"  Dog kred: {dogovor_kred.Наименование if conn.ЗначениеЗаполнено(dogovor_kred) else 'NE NAJDEN'}")

    # 5. Najti sushch. ili sozdat'
    print("\n=== 5. Sozdanie/obnovlenie ===")
    q5 = conn.NewObject("Query")
    q5.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка
    ИЗ Документ.ВзаимозачетЗадолженности КАК Док
    ГДЕ Док.А_ДокументОснованиеДляБаланса = &ДокОснование
        И НЕ Док.ПометкаУдаления"""
    q5.SetParameter("ДокОснование", dok_osnov)
    res5 = q5.Execute()
    sel5 = res5.Choose()

    if sel5.Next():
        dok_obj = sel5.Ссылка.ПолучитьОбъект()
        if dok_obj.Проведен:
            dok_obj.Записать(conn.РежимЗаписиДокумента.ОтменаПроведения)
        dok_obj.ДебиторскаяЗадолженность.Очистить()
        dok_obj.КредиторскаяЗадолженность.Очистить()
        print(f"  Obnovlyaem: {sel5.Ссылка}")
    else:
        dok_obj = conn.Документы.ВзаимозачетЗадолженности.СоздатьДокумент()
        print("  Sozdaem novyj")

    # 6. Zapolnenie shapki
    print("\n=== 6. Zapolnenie ===")
    dok_obj.Дата = dok_osnov.Дата
    dok_obj.Организация = org_dok
    dok_obj.ВидОперации = getattr(conn.Перечисления.ВидыОперацийВзаимозачетаЗадолженности, "Произвольный")
    dok_obj.КонтрагентДебитор = kontr_deb
    dok_obj.КонтрагентКредитор = kontr_kred
    dok_obj.СуммаРегл = summa
    dok_obj.СуммаУпр = summa
    dok_obj.А_ДокументОснованиеДляБаланса = dok_osnov
    dok_obj.А_ВведенВЕРП = True

    # KRITICHNO: TipDebitora i TipKreditora
    dok_obj.ТипДебитора = getattr(conn.Перечисления.ТипыУчастниковВзаимозачета, "Поставщик")
    dok_obj.ТипКредитора = getattr(conn.Перечисления.ТипыУчастниковВзаимозачета, "Клиент")

    valuta_uah = conn.Справочники.Валюты.НайтиПоКоду("980")

    # Debitorskaya
    str_deb = dok_obj.ДебиторскаяЗадолженность.Добавить()
    str_deb.Партнер = partner_deb
    str_deb.ТипРасчетов = getattr(conn.Перечисления.ТипыРасчетовСПартнерами, "РасчетыСПоставщиком")
    str_deb.СуммаВзаиморасчетов = summa
    str_deb.СуммаРегл = summa
    str_deb.СуммаУпр = summa
    str_deb.ВалютаВзаиморасчетов = valuta_uah
    str_deb.Организация = org_dok
    if conn.ЗначениеЗаполнено(dogovor_deb):
        str_deb.ОбъектРасчетов = dogovor_deb

    # Kreditorskaya
    str_kred = dok_obj.КредиторскаяЗадолженность.Добавить()
    str_kred.Партнер = partner_kred
    str_kred.ТипРасчетов = getattr(conn.Перечисления.ТипыРасчетовСПартнерами, "РасчетыСКлиентом")
    str_kred.СуммаВзаиморасчетов = summa
    str_kred.СуммаРегл = summa
    str_kred.СуммаУпр = summa
    str_kred.ВалютаВзаиморасчетов = valuta_uah
    str_kred.Организация = org_dok
    if conn.ЗначениеЗаполнено(dogovor_kred):
        str_kred.ОбъектРасчетов = dogovor_kred

    print(f"  Debitor: {kontr_deb.Наименование}")
    print(f"  Kreditor: {kontr_kred.Наименование}")
    print(f"  Summa: {summa:.2f}")
    print(f"  TipDebitora: Postavshchik")
    print(f"  TipKreditora: Klient")

    # 7. Provedenie
    print("\n=== 7. Provedenie ===")
    try:
        dok_obj.Записать(conn.РежимЗаписиДокумента.Проведение)
        print(f"  OK: Dokument proveden: {dok_obj.Ссылка}")
    except Exception as e:
        print(f"  OSHIBKA provedeniya: {e}")
        try:
            dok_obj.Записать(conn.РежимЗаписиДокумента.Запись)
            print(f"  Zapisan bez provedeniya: {dok_obj.Ссылка}")
        except Exception as e2:
            print(f"  OSHIBKA zapisi: {e2}")
            return

    # 8. Proverka dvizhenij
    print("\n=== 8. Proverka dvizhenij v ProchieAktivyPassivy ===")
    q8 = conn.NewObject("Query")
    q8.Text = """ВЫБРАТЬ
        АП.ВидДвижения КАК ВидДвижения,
        АП.Подразделение КАК Подразделение,
        АП.Статья КАК Статья,
        АП.Аналитика КАК Аналитика,
        АП.Сумма КАК Сумма
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК АП
    ГДЕ АП.Регистратор = &Регистратор"""
    q8.SetParameter("Регистратор", dok_obj.Ссылка)
    res8 = q8.Execute()
    sel8 = res8.Choose()
    count = 0
    while sel8.Next():
        count += 1
        dept = sel8.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel8.Подразделение) else ""
        stat = sel8.Статья.Наименование if conn.ЗначениеЗаполнено(sel8.Статья) else ""
        vid = str(sel8.ВидДвижения)
        print(f"  {vid} | {dept} | {stat} | {float(sel8.Сумма):.2f}")

    if count == 0:
        print("  NET dvizhenij v ProchieAktivyPassivy!")
    else:
        print(f"  Vsego {count} dvizhenij")

    print("\n=== Gotovo ===")


if __name__ == "__main__":
    main()
