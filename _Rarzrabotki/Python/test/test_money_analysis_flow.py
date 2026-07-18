# -*- coding: utf-8 -*-
"""
Діагностика: імітація ТОЧНОГО потоку АнализДокументов_Деньги
для ТАС_Будівн за грудень 2025.

Перевіряє кожен крок:
1. Конвертація Измерение1Ссылка → БанковскийСчетСсылка
2. ERP запит ДенежныеСредстваБезналичные.Обороты
3. ОбробитиДокументиЕРП (статуси ОК/Переобмін/Переказ)
4. ПеревіритиДокументиBuhBud_Деньги (зворотна перевірка)
"""

import win32com.client
import pythoncom
import pywintypes
from datetime import datetime


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    conn_erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    conn_buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

    dt_start = pywintypes.Time(datetime(2025, 12, 1))
    dt_end = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    print("=" * 70)
    print("КРОК 1: Імітація СвестиТаблицы_Деньги → Измерение1Ссылка")
    print("=" * 70)

    # Знаходимо банківський рахунок так само як обробка
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Ссылка, Наименование, НомерСчета "
        "ИЗ Справочник.БанковскиеСчетаОрганизаций "
        "ГДЕ Наименование ПОДОБНО &Имя"
    )
    q.УстановитьПараметр("Имя", "%ТАС_Будівн_%Індастріалбуд%")
    sel = q.Выполнить().Выбрать()
    bank_ref = None
    bank_uid = ""
    while sel.Следующий():
        name = str(sel.Наименование)
        if "$$$" not in name:
            bank_ref = sel.Ссылка
            bank_uid = str(conn_erp.XMLСтрока(sel.Ссылка.УникальныйИдентификатор()))
            print(f"  ERP рахунок: {name}")
            print(f"  UUID (як Строка): '{bank_uid}'")
            print(f"  UUID (верхній регістр): '{bank_uid.upper()}'")
            print(f"  НомерСчета: {sel.НомерСчета}")
            break

    # Це і є Измерение1Ссылка = Запись.Ключ = СокрЛП(Стр.БанковскийСчетUID)
    # В ПолучитьОстаткиBuhBud_Деньги: НС.БанковскийСчетUID = Строка(СсылкаЕРП.УникальныйИдентификатор())
    izmerenie1_ssylka = bank_uid
    print(f"\n  Измерение1Ссылка = '{izmerenie1_ssylka}'")

    print("\n" + "=" * 70)
    print("КРОК 2: Імітація АнализДокументов_Деньги — конвертація UUID")
    print("=" * 70)

    # В обробці: БанковскийСчетСсылка = Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(
    #    Новый УникальныйИдентификатор(СтрРасх.Измерение1Ссылка));
    try:
        bank_ref_restored = conn_erp.Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(
            conn_erp.NewObject("УникальныйИдентификатор", izmerenie1_ssylka))
        print(f"  ПолучитьСсылку: OK")
        print(f"  Наименование: {bank_ref_restored}")
        print(f"  Пуста: {not conn_erp.ЗначениеЗаполнено(bank_ref_restored)}")
    except Exception as e:
        print(f"  ПОМИЛКА конвертації UUID: {e}")
        return

    print("\n" + "=" * 70)
    print("КРОК 3: ERP запит ДенежныеСредстваБезналичные.Обороты")
    print("=" * 70)

    q2 = conn_erp.NewObject("Запрос")
    q2.Текст = (
        "ВЫБРАТЬ "
        "  Об.Регистратор КАК Регистратор, "
        "  ПРЕДСТАВЛЕНИЕ(Об.Регистратор) КАК РегистраторПредст, "
        "  Об.СуммаПриход КАК Приход, "
        "  Об.СуммаРасход КАК Расход, "
        "  Соотв.УникальныйИдентификаторПриемника КАК UIDвBuhBud, "
        "  Соотв.ТипПриемника КАК ТипПриемника, "
        "  Соотв.ТипИсточника КАК ТипИсточника "
        "ИЗ "
        "  РегистрНакопления.ДенежныеСредстваБезналичные.Обороты("
        "    &НачалоПериода, &КонецПериода, Авто, "
        "    БанковскийСчет = &БанковскийСчет) КАК Об "
        "  ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК Соотв "
        "  ПО Об.Регистратор = Соотв.УникальныйИдентификаторИсточника "
        "    И (Соотв.УзелИнформационнойБазы ССЫЛКА ПланОбмена.ОбменУправлениеПредприятиемБухгалтерия20)"
    )
    q2.УстановитьПараметр("НачалоПериода", dt_start)
    q2.УстановитьПараметр("КонецПериода", dt_end)
    q2.УстановитьПараметр("БанковскийСчет", bank_ref_restored)

    try:
        sel2 = q2.Выполнить().Выбрать()
    except Exception as e:
        print(f"  ПОМИЛКА ERP запиту: {e}")
        return

    erp_uids = set()
    total_p = 0
    total_r = 0
    cnt_ok = 0
    cnt_transfer = 0
    cnt_no_buh = 0
    cnt_diff = 0

    while sel2.Следующий():
        uid = str(conn_erp.XMLСтрока(sel2.Регистратор.УникальныйИдентификатор())).upper()
        erp_uids.add(uid)
        prihod = float(sel2.Приход) if sel2.Приход else 0
        rashod = float(sel2.Расход) if sel2.Расход else 0
        total_p += prihod
        total_r += rashod
        net_erp = prihod - rashod

        # Імітація ОбробитиДокументиЕРП
        # 1. Перевірка переказу
        is_transfer = False
        try:
            ho = sel2.Регистратор.ХозяйственнаяОперация
            hoz_op = str(conn_erp.String(ho))
            if "ДругогоСчета" in hoz_op or "ДругойСчет" in hoz_op:
                is_transfer = True
        except:
            pass

        if is_transfer:
            cnt_transfer += 1
            continue

        # 2. Перевірка відповідності
        uid_buh = ""
        typ_priem = ""
        try:
            if sel2.UIDвBuhBud:
                uid_buh = str(sel2.UIDвBuhBud).upper()
            if sel2.ТипПриемника:
                typ_priem = str(sel2.ТипПриемника).replace("ДокументСсылка.", "")
        except:
            pass

        if uid_buh and typ_priem:
            # Перевірка в BuhBud — сума по ВСІХ банках
            try:
                buh_doc_ref = conn_buh.Документы[typ_priem].ПолучитьСсылку(
                    conn_buh.NewObject("УникальныйИдентификатор", uid_buh))
                q_buh = conn_buh.NewObject("Запрос")
                q_buh.Текст = (
                    "ВЫБРАТЬ СУММА(Т.Приход) - СУММА(Т.Расход) КАК Сумма "
                    "ИЗ ("
                    "  ВЫБРАТЬ Д.Сумма КАК Приход, 0 КАК Расход "
                    "  ИЗ РегистрБухгалтерии.Хозрасчетный КАК Д "
                    "  ГДЕ Д.Регистратор = &Регистратор "
                    "    И Д.СчетДт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)) "
                    "  ОБЪЕДИНИТЬ ВСЕ "
                    "  ВЫБРАТЬ 0, Д.Сумма "
                    "  ИЗ РегистрБухгалтерии.Хозрасчетный КАК Д "
                    "  ГДЕ Д.Регистратор = &Регистратор "
                    "    И Д.СчетКт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))) КАК Т"
                )
                q_buh.УстановитьПараметр("Регистратор", buh_doc_ref)
                sel_buh = q_buh.Выполнить().Выбрать()
                net_buh_all = 0
                if sel_buh.Следующий():
                    net_buh_all = float(sel_buh.Сумма) if sel_buh.Сумма else 0

                if abs(round(net_erp - net_buh_all, 2)) == 0:
                    cnt_ok += 1
                else:
                    cnt_diff += 1
            except:
                cnt_diff += 1
        else:
            cnt_no_buh += 1

    print(f"  ERP документів: {len(erp_uids)}")
    print(f"  Приход: {total_p:,.2f}")
    print(f"  Расход: {total_r:,.2f}")
    print(f"\n  Статуси (імітація ОбробитиДокументиЕРП):")
    print(f"    ОК: {cnt_ok}")
    print(f"    Переказ: {cnt_transfer}")
    print(f"    Без відповідності: {cnt_no_buh}")
    print(f"    Різниця: {cnt_diff}")
    print(f"    ВСЬОГО рядків в ТаблицаДокументов: {cnt_ok + cnt_transfer + cnt_no_buh + cnt_diff}")

    print("\n" + "=" * 70)
    print("КРОК 4: Зворотна перевірка (ПеревіритиДокументиBuhBud_Деньги)")
    print("=" * 70)

    # Крок 4a: IBAN з ERP рахунку
    try:
        iban = str(bank_ref_restored.НомерСчета).strip()
        print(f"  IBAN: {iban}")
    except Exception as e:
        print(f"  ПОМИЛКА отримання IBAN: {e}")
        return

    # Крок 4b: Пошук BuhBud рахунку по IBAN (виключаємо $$$)
    q_bank = conn_buh.NewObject("Запрос")
    q_bank.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Наименование "
        "ИЗ Справочник.БанковскиеСчета "
        "ГДЕ НомерСчета = &НомерСчета "
        "И Наименование <> \"\" "
        "И Наименование НЕ ПОДОБНО \"%$$$%\" "
        "УПОРЯДОЧИТЬ ПО Наименование"
    )
    q_bank.УстановитьПараметр("НомерСчета", iban)
    sel_bank = q_bank.Выполнить().Выбрать()
    buh_bank_ref = None
    if sel_bank.Следующий():
        buh_bank_ref = sel_bank.Ссылка
        print(f"  BuhBud рахунок: {sel_bank.Наименование}")
    else:
        print(f"  ПОМИЛКА: BuhBud рахунок не знайдено!")
        return

    # Крок 4c: ОстаткиИОбороты з Субконто1=рахунок, Периодичність=Регистратор
    q_buh_docs = conn_buh.NewObject("Запрос")
    q_buh_docs.Текст = (
        "ВЫБРАТЬ "
        "  Т.Регистратор КАК Регистратор, "
        "  ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК Предст, "
        "  Т.СуммаОборотДт КАК Приход, "
        "  Т.СуммаОборотКт КАК Расход "
        "ИЗ "
        "  РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты("
        "    &НачалоПериода, &КонецПериода, Регистратор, , "
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , "
        "    Субконто1 = &БанкРахунок) КАК Т"
    )
    q_buh_docs.УстановитьПараметр("НачалоПериода", dt_start)
    q_buh_docs.УстановитьПараметр("КонецПериода", dt_end)
    q_buh_docs.УстановитьПараметр("БанкРахунок", buh_bank_ref)

    try:
        sel_bd = q_buh_docs.Выполнить().Выбрать()
    except Exception as e:
        print(f"  ПОМИЛКА запиту ОстаткиИОбороты: {e}")
        return

    buh_only_count = 0
    buh_only_with_diff = 0
    target_found = False

    while sel_bd.Следующий():
        try:
            reg = sel_bd.Регистратор
            if not conn_buh.ЗначениеЗаполнено(reg):
                continue
            uid = str(conn_buh.XMLСтрока(reg.УникальныйИдентификатор())).upper()
        except:
            continue

        if uid in erp_uids:
            continue  # matched

        prihod = float(sel_bd.Приход) if sel_bd.Приход else 0
        rashod = float(sel_bd.Расход) if sel_bd.Расход else 0
        net = prihod - rashod

        if abs(round(net, 2)) == 0:
            continue  # нульовий потік

        buh_only_count += 1
        buh_only_with_diff += 1

        predst = str(sel_bd.Предст)
        is_target = abs(net + 2824601) < 1
        if is_target:
            target_found = True

        if buh_only_count <= 5 or is_target:
            mark = " <<<< ШУКАНИЙ!" if is_target else ""
            print(f"  BuhBud-only: {predst[:55]:55s} Нетто={net:>14,.2f}{mark}")

    if buh_only_count > 5:
        print(f"  ... та ще {buh_only_count - 5} документів")

    print(f"\n  BuhBud-only документів: {buh_only_count}")
    print(f"  Документ -2,824,601 знайдено: {'ТАК' if target_found else 'НІ'}")

    print("\n" + "=" * 70)
    print("ПІДСУМОК")
    print("=" * 70)
    total_docs = cnt_ok + cnt_transfer + cnt_no_buh + cnt_diff + buh_only_count
    print(f"  Всього документів в ТаблицаДокументов: {total_docs}")
    print(f"    ОК: {cnt_ok}")
    print(f"    Переказ: {cnt_transfer}")
    print(f"    Без відповідності: {cnt_no_buh}")
    print(f"    Різниця ERP↔BuhBud: {cnt_diff}")
    print(f"    BuhBud-only: {buh_only_count}")
    print(f"  З галочкою 'Тільки розбіжності': {cnt_no_buh + cnt_diff + buh_only_count}")
    print(f"  Документ-причина 00019117 знайдено: {'ТАК ✓' if target_found else 'НІ ✗'}")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
