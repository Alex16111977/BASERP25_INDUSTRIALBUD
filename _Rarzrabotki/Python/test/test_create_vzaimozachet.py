# -*- coding: utf-8 -*-
"""
Тест: повний цикл створення ВзаимозачетЗадолженности через COM
Документ-основа: СписаниеНедостачТоваров ІБ00-000752 от 31.12.2025
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


def main():
    conn = connect()

    date_from = datetime(2025, 12, 1)
    date_to = datetime(2025, 12, 31, 23, 59, 59)

    # === 1. Znajti dokument ===
    print("\n=== 1. Poisk dokumenta ===")
    q = conn.NewObject("Query")
    q.Text = """ВЫБРАТЬ Док.Ссылка КАК Ссылка
    ИЗ Документ.СписаниеНедостачТоваров КАК Док
    ГДЕ Док.Номер ПОДОБНО &Номер И Док.Дата МЕЖДУ &Нач И &Кон"""
    q.SetParameter("Номер", "%000752%")
    q.SetParameter("Нач", date_from)
    q.SetParameter("Кон", date_to)
    res = q.Execute()
    sel = res.Choose()
    if not sel.Next():
        print("OSHIBKA: dokument ne najden!")
        return
    dok_osnov = sel.Ссылка
    print(f"  Najdeno: {dok_osnov}")
    print(f"  Data: {dok_osnov.Дата}")

    # === 2. Disbalans po podrazdeleniyam ===
    print("\n=== 2. Disbalans po podrazdeleniyam ===")
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
        print("  Menee 2 podrazdelenij — vzaimozachet ne sozdaetsya")
        return

    print(f"  Summa: {summa:.2f}")

    # === 3. Najti partnerov i kontragentov ===
    print("\n=== 3. Poisk partnerov i kontragentov ===")

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
        # fallback po naimenovaniyu
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

    partner_deb = najti_partnera(conn, podrazd_debitor)
    if not conn.ЗначениеЗаполнено(partner_deb):
        print(f"  OSHIBKA: Partner ne najden dlya debitora")
        return
    print(f"  Partner debitor: {partner_deb.Наименование}")

    kontr_deb = najti_kontragenta(conn, partner_deb)
    if not conn.ЗначениеЗаполнено(kontr_deb):
        print(f"  OSHIBKA: Kontragent ne najden dlya debitora")
        return
    print(f"  Kontragent debitor: {kontr_deb.Наименование}")

    partner_kred = najti_partnera(conn, podrazd_kreditor)
    if not conn.ЗначениеЗаполнено(partner_kred):
        print(f"  OSHIBKA: Partner ne najden dlya kreditora")
        return
    print(f"  Partner kreditor: {partner_kred.Наименование}")

    kontr_kred = najti_kontragenta(conn, partner_kred)
    if not conn.ЗначениеЗаполнено(kontr_kred):
        print(f"  OSHIBKA: Kontragent ne najden dlya kreditora")
        return
    print(f"  Kontragent kreditor: {kontr_kred.Наименование}")

    # === 4. Najti dogovory ===
    print("\n=== 4. Poisk dogovorov ===")

    def najti_dogovor(conn, kontragent, tip_dogovora_name):
        q = conn.NewObject("Query")
        q.Text = (
            "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Договор"
            " ИЗ Справочник.ДоговорыКонтрагентов КАК Д"
            " ГДЕ Д.Контрагент = &Контрагент"
            "   И Д.ТипДоговора = ЗНАЧЕНИЕ(Перечисление.ТипыДоговоров." + tip_dogovora_name + ")"
            "   И НЕ Д.ПометкаУдаления"
        )
        q.SetParameter("Контрагент", kontragent)
        res = q.Execute()
        sel = res.Choose()
        if sel.Next():
            return sel.Договор
        return None

    dogovor_deb = najti_dogovor(conn, kontr_deb, "СПоставщиком")
    if conn.ЗначениеЗаполнено(dogovor_deb):
        print(f"  Dogovor debitora: {dogovor_deb.Наименование}")
    else:
        print("  Dogovor debitora NE najden — budem sozdavat'")
        # Sozdaem dogovor
        new_dog = conn.Справочники.ДоговорыКонтрагентов.СоздатьЭлемент()
        new_dog.Наименование = str(podrazd_debitor)
        new_dog.Контрагент = kontr_deb
        new_dog.Партнер = kontr_deb.Партнер
        new_dog.Организация = org_dok
        new_dog.Подразделение = podrazd_debitor
        new_dog.ТипДоговора = getattr(conn.Перечисления.ТипыДоговоров, "СПоставщиком")
        new_dog.ВалютаВзаиморасчетов = conn.Справочники.Валюты.НайтиПоКоду("980")
        new_dog.ПорядокРасчетов = conn.Перечисления.ПорядокРасчетов.ПоДоговорамКонтрагентов
        new_dog.ГруппаФинансовогоУчета = conn.Справочники.ГруппыФинансовогоУчетаРасчетов.НайтиПоНаименованию(
            "Расчеты с поставщиками в национальной валюте")
        new_dog.Статус = conn.Перечисления.СтатусыДоговоровКонтрагентов.Действует
        new_dog.Согласован = True
        new_dog.НалогообложениеНДС = conn.Перечисления.ТипыНалогообложенияНДС.ПродажаОсвобожденаОтНДС
        new_dog.ВариантОформленияЗакупок = conn.Перечисления.ВариантыОформленияЗакупок.НеРазделять
        new_dog.МоментОпределенияБазыНДС = conn.Перечисления.МоментыОпределенияНалоговойБазы.ПоОтгрузке
        new_dog.СпособДоставки = conn.Перечисления.СпособыДоставки.ОпределяетсяВРаспоряжении
        new_dog.Записать()
        dogovor_deb = new_dog.Ссылка
        print(f"  Sozdan dogovor debitora: {dogovor_deb.Наименование}")

    dogovor_kred = najti_dogovor(conn, kontr_kred, "СПокупателем")
    if conn.ЗначениеЗаполнено(dogovor_kred):
        print(f"  Dogovor kreditora: {dogovor_kred.Наименование}")
    else:
        print("  Dogovor kreditora NE najden — budem sozdavat'")
        new_dog = conn.Справочники.ДоговорыКонтрагентов.СоздатьЭлемент()
        new_dog.Наименование = str(podrazd_kreditor)
        new_dog.Контрагент = kontr_kred
        new_dog.Партнер = kontr_kred.Партнер
        new_dog.Организация = org_dok
        new_dog.Подразделение = podrazd_kreditor
        new_dog.ТипДоговора = getattr(conn.Перечисления.ТипыДоговоров, "СПокупателем")
        new_dog.ВалютаВзаиморасчетов = conn.Справочники.Валюты.НайтиПоКоду("980")
        new_dog.ПорядокРасчетов = conn.Перечисления.ПорядокРасчетов.ПоДоговорамКонтрагентов
        new_dog.ГруппаФинансовогоУчета = conn.Справочники.ГруппыФинансовогоУчетаРасчетов.НайтиПоНаименованию(
            "Расчеты с покупателями в национальной валюте")
        new_dog.Статус = conn.Перечисления.СтатусыДоговоровКонтрагентов.Действует
        new_dog.Согласован = True
        new_dog.НалогообложениеНДС = conn.Перечисления.ТипыНалогообложенияНДС.ПродажаОсвобожденаОтНДС
        new_dog.ВариантОформленияЗакупок = conn.Перечисления.ВариантыОформленияЗакупок.НеРазделять
        new_dog.МоментОпределенияБазыНДС = conn.Перечисления.МоментыОпределенияНалоговойБазы.ПоОтгрузке
        new_dog.СпособДоставки = conn.Перечисления.СпособыДоставки.ОпределяетсяВРаспоряжении
        new_dog.Записать()
        dogovor_kred = new_dog.Ссылка
        print(f"  Sozdan dogovor kreditora: {dogovor_kred.Наименование}")

    # === 5. Najti sushch. ili sozdat' VzaimozachetZadolzhennosti ===
    print("\n=== 5. Sozdanie/obnovlenie VzaimozachetZadolzhennosti ===")

    q5 = conn.NewObject("Query")
    q5.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка
    ИЗ Документ.ВзаимозачетЗадолженности КАК Док
    ГДЕ Док.А_ДокументОснованиеДляБаланса = &ДокОснование
        И НЕ Док.ПометкаУдаления"""
    q5.SetParameter("ДокОснование", dok_osnov)
    res5 = q5.Execute()
    sel5 = res5.Choose()

    dok_obj = None
    if sel5.Next():
        dok_obj = sel5.Ссылка.ПолучитьОбъект()
        if dok_obj.Проведен:
            dok_obj.Записать(conn.РежимЗаписиДокумента.ОтменаПроведения)
        dok_obj.ДебиторскаяЗадолженность.Очистить()
        dok_obj.КредиторскаяЗадолженность.Очистить()
        print(f"  Obnovlyaem sushchestvuyushchij: {sel5.Ссылка}")
    else:
        dok_obj = conn.Документы.ВзаимозачетЗадолженности.СоздатьДокумент()
        print("  Sozdaem novyj dokument")

    # === 6. Zapolnit' shapku ===
    print("\n=== 6. Zapolnenie dokumenta ===")

    dok_obj.Дата = dok_osnov.Дата
    dok_obj.Организация = org_dok
    dok_obj.ВидОперации = getattr(conn.Перечисления.ВидыОперацийВзаимозачетаЗадолженности, "Произвольный")
    dok_obj.КонтрагентДебитор = kontr_deb
    dok_obj.КонтрагентКредитор = kontr_kred
    dok_obj.СуммаРегл = summa
    dok_obj.СуммаУпр = summa
    dok_obj.А_ДокументОснованиеДляБаланса = dok_osnov

    valuta_uah = conn.Справочники.Валюты.НайтиПоКоду("980")

    # Debitorskaya
    str_deb = dok_obj.ДебиторскаяЗадолженность.Добавить()
    str_deb.Партнер = partner_deb
    str_deb.ТипРасчетов = conn.Перечисления.ТипыРасчетовСПартнерами.РасчетыСПоставщиком
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
    str_kred.ТипРасчетов = conn.Перечисления.ТипыРасчетовСПартнерами.РасчетыСКлиентом
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

    # === 7. Provesti ===
    print("\n=== 7. Provedenie ===")
    try:
        dok_obj.Записать(conn.РежимЗаписиДокумента.Проведение)
        print(f"  OK: Dokument proveden: {dok_obj.Ссылка}")
    except Exception as e:
        print(f"  OSHIBKA provedeniya: {e}")
        try:
            dok_obj.Записать(conn.РежимЗаписиДокумента.Запись)
            print(f"  Dokument zapisan bez provedeniya: {dok_obj.Ссылка}")
        except Exception as e2:
            print(f"  OSHIBKA zapisi: {e2}")

    print("\n=== Gotovo ===")


if __name__ == "__main__":
    main()
