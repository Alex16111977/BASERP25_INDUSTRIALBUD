# -*- coding: utf-8 -*-
"""
Точна діагностика: перевірка кожного кроку АнализДокументов_Деньги
для ТАС_Будівн за грудень 2025.
Імітує 1С код максимально точно з перевіркою помилок на кожному кроці.
"""

import win32com.client
import pythoncom
import pywintypes
from datetime import datetime


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    conn_erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    conn_buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

    dt_start = pywintypes.Time(datetime(2025, 12, 1))
    dt_end = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    # === КРОК 0: Перевірити як СвестиТаблицы_Деньги зберігає Измерение1Ссылка ===
    print("=" * 70)
    print("КРОК 0: Перевірка Измерение1Ссылка (як в СвестиТаблицы_Деньги)")
    print("=" * 70)

    # Знаходимо банківський рахунок в ERP
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Ссылка, Наименование, НомерСчета "
        "ИЗ Справочник.БанковскиеСчетаОрганизаций "
        "ГДЕ Наименование ПОДОБНО &Имя И Наименование НЕ ПОДОБНО \"%$$$%\""
    )
    q.УстановитьПараметр("Имя", "%ТАС_Будівн_%Індастріалбуд%")
    sel = q.Выполнить().Выбрать()
    if not sel.Следующий():
        print("  ПОМИЛКА: рахунок не знайдено!")
        return
    bank_ref = sel.Ссылка
    bank_uid = str(conn_erp.XMLСтрока(sel.Ссылка.УникальныйИдентификатор()))
    print(f"  Рахунок: {sel.Наименование}")
    print(f"  UUID: {bank_uid}")
    print(f"  Это Измерение1Ссылка, записаний як Строка(СсылкаЕРП.УникальныйИдентификатор())")

    # === КРОК 1: Конвертація UUID → ссылка (як АнализДокументов_Деньги) ===
    print("\n" + "=" * 70)
    print("КРОК 1: ПолучитьСсылку (як рядок 1437-1438)")
    print("=" * 70)
    try:
        # Імітація: Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(
        #     Новый УникальныйИдентификатор(СтрРасх.Измерение1Ссылка));
        uid_obj = conn_erp.NewObject("УникальныйИдентификатор", bank_uid)
        bank_ref2 = conn_erp.Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(uid_obj)
        is_empty = not conn_erp.ЗначениеЗаполнено(bank_ref2)
        print(f"  ОК. Пуста: {is_empty}")
    except Exception as e:
        print(f"  ПОМИЛКА: {e}")
        return

    # === КРОК 2: ERP запит (рядки 1440-1462) ===
    print("\n" + "=" * 70)
    print("КРОК 2: ERP запит ДенежныеСредстваБезналичные.Обороты")
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
    q2.УстановитьПараметр("БанковскийСчет", bank_ref2)

    try:
        res2 = q2.Выполнить()
        tz = res2.Выгрузить()
        cnt = tz.Количество()
        print(f"  ОК. Документів: {cnt}")
    except Exception as e:
        print(f"  ПОМИЛКА запиту: {e}")
        return

    # === КРОК 3: ОбробитиДокументиЕРП (рядки 1645-1721) ===
    print("\n" + "=" * 70)
    print("КРОК 3: ОбробитиДокументиЕРП — перевірка кожного документа")
    print("=" * 70)

    erp_uids = set()
    cnt_transfer = 0
    cnt_ok = 0
    cnt_diff = 0
    cnt_no_match = 0
    cnt_error = 0
    errors = []

    for i in range(cnt):
        row = tz.Получить(i)
        doc_num = i + 1

        try:
            # 1. Рядок 1649-1656: Базові поля
            reg = row.Регистратор
            predst = str(row.РегистраторПредст)
            prihod = float(row.Приход) if row.Приход else 0
            rashod = float(row.Расход) if row.Расход else 0
            net_erp = prihod - rashod

            # UIDДокумента = Строка(СтрДок.Регистратор.УникальныйИдентификатор())
            uid = str(conn_erp.XMLСтрока(reg.УникальныйИдентификатор())).upper()
            erp_uids.add(uid)

            # 2. Рядки 1661-1668: ТипДокумента
            typ_ist = ""
            try:
                typ_ist = str(row.ТипИсточника)
                typ_ist = typ_ist.replace("ДокументСсылка.", "")
            except:
                try:
                    md = reg.Метаданные()
                    typ_ist = str(md.Имя)
                except:
                    typ_ist = "?"

            # 3. Рядки 1671-1685: Перевірка переказу (тільки для Гроші)
            is_transfer = False
            try:
                ho = reg.ХозяйственнаяОперация
                hoz_str = str(conn_erp.XMLСтрока(ho))
                if "ДругогоСчета" in hoz_str or "ДругойСчет" in hoz_str:
                    is_transfer = True
            except:
                pass

            if is_transfer:
                cnt_transfer += 1
                continue

            # 4. Рядки 1688-1704: UIDвBuhBud та ТипПриемника
            uid_buh = ""
            typ_priem = ""
            try:
                if conn_erp.ЗначениеЗаполнено(row.UIDвBuhBud):
                    uid_buh = str(row.UIDвBuhBud).strip()
            except:
                pass
            try:
                if conn_erp.ЗначениеЗаполнено(row.ТипПриемника):
                    typ_priem = str(row.ТипПриемника).strip()
                    typ_priem = typ_priem.replace("ДокументСсылка.", "")
            except:
                pass

            if uid_buh and typ_priem:
                # 5. ПроверитьДокументВBuhBud (рядки 1725-1820)
                try:
                    # Рядок 1733: V83.Документы[ТипПриемника].ПолучитьСсылку(...)
                    doc_mgr = getattr(conn_buh.Документы, typ_priem)
                    buh_ref = doc_mgr.ПолучитьСсылку(
                        conn_buh.NewObject("УникальныйИдентификатор", uid_buh))

                    # Рядок 1729-1731: Запит по документу
                    q_chk = conn_buh.NewObject("Запрос")
                    q_chk.Текст = (
                        f"ВЫБРАТЬ Ссылка, Проведен, ПометкаУдаления, "
                        f"ПРЕДСТАВЛЕНИЕ(Ссылка) КАК Представление "
                        f"ИЗ Документ.{typ_priem} ГДЕ Ссылка = &Ссылка"
                    )
                    q_chk.УстановитьПараметр("Ссылка", buh_ref)
                    sel_chk = q_chk.Выполнить().Выбрать()

                    if sel_chk.Следующий():
                        if sel_chk.ПометкаУдаления:
                            status = "Видалений"
                        elif not sel_chk.Проведен:
                            status = "Не проведений"
                        else:
                            # Рядки 1776-1798: Запит суми в BuhBud (UNION ALL)
                            try:
                                q_sum = conn_buh.NewObject("Запрос")
                                q_sum.Текст = (
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
                                q_sum.УстановитьПараметр("Регистратор", buh_ref)
                                sel_sum = q_sum.Выполнить().Выбрать()
                                net_buh = 0
                                if sel_sum.Следующий():
                                    net_buh = float(sel_sum.Сумма) if sel_sum.Сумма else 0

                                if abs(round(net_erp - net_buh, 2)) == 0:
                                    status = "ОК"
                                    cnt_ok += 1
                                else:
                                    status = f"Різниця (ERP={net_erp:,.2f}, Buh={net_buh:,.2f})"
                                    cnt_diff += 1
                            except Exception as e:
                                status = f"Помилка суми: {e}"
                                cnt_diff += 1
                    else:
                        status = "Не знайдено в BuhBud"
                        cnt_diff += 1
                except Exception as e:
                    status = f"Помилка COM: {e}"
                    cnt_error += 1
                    errors.append(f"  #{doc_num} {predst[:50]}: {e}")
            else:
                status = "Немає відповідності"
                cnt_no_match += 1

            if doc_num <= 5 or (cnt_error > 0 and cnt_error <= 3):
                print(f"  #{doc_num}: {predst[:55]} | {status[:30]}")

        except Exception as e:
            cnt_error += 1
            errors.append(f"  #{doc_num}: КРИТИЧНА ПОМИЛКА: {e}")
            if cnt_error <= 5:
                print(f"  #{doc_num}: КРИТИЧНА ПОМИЛКА: {e}")

    print(f"\n  Підсумок forward check:")
    print(f"    ОК: {cnt_ok}")
    print(f"    Переказ: {cnt_transfer}")
    print(f"    Без відповідності: {cnt_no_match}")
    print(f"    Різниця: {cnt_diff}")
    print(f"    Помилки: {cnt_error}")
    if errors:
        print(f"\n  Перші помилки:")
        for e in errors[:5]:
            print(f"    {e}")

    # === КРОК 4: Зворотна перевірка ===
    print("\n" + "=" * 70)
    print("КРОК 4: ПеревіритиДокументиBuhBud_Деньги")
    print("=" * 70)

    # 4a: IBAN
    try:
        iban = str(bank_ref2.НомерСчета).strip()
        print(f"  IBAN: {iban}")
    except Exception as e:
        print(f"  ПОМИЛКА IBAN: {e}")
        return

    # 4b: BuhBud рахунок
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
    if not sel_bank.Следующий():
        print(f"  ПОМИЛКА: BuhBud рахунок не знайдено!")
        return
    buh_bank_ref = sel_bank.Ссылка
    print(f"  BuhBud рахунок: {sel_bank.Наименование}")

    # 4c: ОстаткиИОбороты
    q_buh = conn_buh.NewObject("Запрос")
    q_buh.Текст = (
        "ВЫБРАТЬ "
        "  Т.Регистратор КАК Регистратор, "
        "  ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК Представление, "
        "  Т.СуммаОборотДт КАК Приход, "
        "  Т.СуммаОборотКт КАК Расход "
        "ИЗ "
        "  РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты("
        "    &НачалоПериода, &КонецПериода, Регистратор, , "
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , "
        "    Субконто1 = &БанкРахунок) КАК Т"
    )
    q_buh.УстановитьПараметр("НачалоПериода", dt_start)
    q_buh.УстановитьПараметр("КонецПериода", dt_end)
    q_buh.УстановитьПараметр("БанкРахунок", buh_bank_ref)

    try:
        sel_bd = q_buh.Выполнить().Выбрать()
    except Exception as e:
        print(f"  ПОМИЛКА ОстаткиИОбороты: {e}")
        return

    buh_only = 0
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
            continue

        prihod = float(sel_bd.Приход) if sel_bd.Приход else 0
        rashod = float(sel_bd.Расход) if sel_bd.Расход else 0
        net = prihod - rashod
        if abs(round(net, 2)) == 0:
            continue

        buh_only += 1
        predst = str(sel_bd.Представление)
        is_target = abs(net + 2824601) < 1
        if is_target:
            target_found = True
        if buh_only <= 3 or is_target:
            mark = " <<<< ШУКАНИЙ!" if is_target else ""
            print(f"  BuhBud-only: {predst[:55]} Нетто={net:>14,.2f}{mark}")

    if buh_only > 3:
        print(f"  ... та ще {buh_only - 3} документів")
    print(f"  BuhBud-only всього: {buh_only}")
    print(f"  Документ 00019117 знайдено: {'ТАК' if target_found else 'НІ'}")

    # === ПІДСУМОК ===
    print("\n" + "=" * 70)
    print("ПІДСУМОК")
    print("=" * 70)
    total = cnt_ok + cnt_transfer + cnt_no_match + cnt_diff + cnt_error + buh_only
    total_visible = cnt_no_match + cnt_diff + cnt_error + buh_only
    print(f"  Всього документів: {total}")
    print(f"    ОК (Синхронизировать=False): {cnt_ok}")
    print(f"    Переказ (Синхронизировать=False): {cnt_transfer}")
    print(f"    Без відповідності (Синхронизировать=True): {cnt_no_match}")
    print(f"    Різниця (Синхронизировать=True): {cnt_diff}")
    print(f"    Помилки (Синхронизировать=True): {cnt_error}")
    print(f"    BuhBud-only (Синхронизировать=True): {buh_only}")
    print(f"  Видимих з фільтром 'Тільки розбіжності': {total_visible}")

    if total > 0 and cnt_error == 0:
        print(f"\n  ВИСНОВОК: Код має працювати! Таблиця повинна заповнюватись ({total} рядків).")
        print(f"  Якщо в 1С таблиця порожня — можливо обробка не перезавантажена з файлів.")
    elif cnt_error > 0:
        print(f"\n  ВИСНОВОК: Є {cnt_error} помилок — вони можуть спричинити збій в 1С.")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
