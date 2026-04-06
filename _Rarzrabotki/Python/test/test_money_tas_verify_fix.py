# -*- coding: utf-8 -*-
"""
Верифікація: виправлена логіка зворотної перевірки (Гроші)
знаходить документ-причину розбіжності 2,824,601 на ТАС_Будівн.

Імітує НОВУ логіку ПеревіритиДокументиBuhBud_Деньги:
1. Знаходить BuhBud банк.рахунок по IBAN
2. Запит ОстаткиИОбороты з Субконто1=конкретний рахунок, Периодичність=Регистратор
3. Порівнює з ERP документами за UUID
4. Показує BuhBud-only документи
"""

import win32com.client
import pythoncom
import pywintypes
from datetime import datetime


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    print("=" * 80)
    print("ВЕРИФІКАЦІЯ: Виправлена зворотна перевірка Гроші")
    print("=" * 80)

    conn_erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    conn_buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

    dt_start = pywintypes.Time(datetime(2025, 12, 1))
    dt_end = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    # 1. Знайти ERP рахунок
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Ссылка, Наименование, НомерСчета "
        "ИЗ Справочник.БанковскиеСчетаОрганизаций "
        "ГДЕ Наименование ПОДОБНО &Имя"
    )
    q.УстановитьПараметр("Имя", "%ТАС_Будівн_%Індастріалбуд%")
    sel = q.Выполнить().Выбрать()
    bank_ref = None
    iban = ""
    while sel.Следующий():
        name = str(sel.Наименование)
        if "$$$" not in name:
            bank_ref = sel.Ссылка
            iban = str(sel.НомерСчета).strip()
            print(f"  ERP рахунок: {name}")
            print(f"  IBAN: {iban}")
            break

    if not bank_ref:
        print("  ERP рахунок не знайдено!")
        return

    # 2. Отримати ERP документи (масUIDерп)
    print("\n--- ERP документи ---")
    q_erp = conn_erp.NewObject("Запрос")
    q_erp.Текст = (
        "ВЫБРАТЬ "
        "  Об.Регистратор КАК Регистратор "
        "ИЗ "
        "  РегистрНакопления.ДенежныеСредстваБезналичные.Обороты("
        "    &НачалоПериода, &КонецПериода, Авто, "
        "    БанковскийСчет = &БанковскийСчет) КАК Об"
    )
    q_erp.УстановитьПараметр("НачалоПериода", dt_start)
    q_erp.УстановитьПараметр("КонецПериода", dt_end)
    q_erp.УстановитьПараметр("БанковскийСчет", bank_ref)
    sel_erp = q_erp.Выполнить().Выбрать()

    mas_uid_erp = set()
    while sel_erp.Следующий():
        uid = str(conn_erp.XMLСтрока(sel_erp.Регистратор.УникальныйИдентификатор())).upper()
        mas_uid_erp.add(uid)

    print(f"  ERP документів: {len(mas_uid_erp)}")

    # 3. Знайти BuhBud рахунок по IBAN (НОВА ЛОГІКА)
    print("\n--- BuhBud: пошук рахунку по IBAN ---")
    q_bank = conn_buh.NewObject("Запрос")
    q_bank.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Наименование "
        "ИЗ Справочник.БанковскиеСчета "
        "ГДЕ НомерСчета = &НомерСчета И Наименование <> \"\" "
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
        print("  BuhBud рахунок не знайдено!")
        return

    # 4. НОВА ЛОГІКА: ОстаткиИОбороты з Субконто1, Периодичність=Регистратор
    print("\n--- BuhBud: ОстаткиИОбороты (Регистратор, Субконто1=рахунок) ---")
    q_buh = conn_buh.NewObject("Запрос")
    q_buh.Текст = (
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
    q_buh.УстановитьПараметр("НачалоПериода", dt_start)
    q_buh.УстановитьПараметр("КонецПериода", dt_end)
    q_buh.УстановитьПараметр("БанкРахунок", buh_bank_ref)

    sel_buh = q_buh.Выполнить().Выбрать()

    buh_only_docs = []
    total_buh = 0
    total_matched = 0
    total_buh_only_net = 0

    while sel_buh.Следующий():
        try:
            reg = sel_buh.Регистратор
            if not conn_buh.ЗначениеЗаполнено(reg):
                continue
            uid = str(conn_buh.XMLСтрока(reg.УникальныйИдентификатор())).upper()
        except:
            continue

        total_buh += 1
        prihod = float(sel_buh.Приход) if sel_buh.Приход else 0
        rashod = float(sel_buh.Расход) if sel_buh.Расход else 0
        net = prihod - rashod

        if uid in mas_uid_erp:
            total_matched += 1
            continue

        # Пропустити нульові
        if abs(round(net, 2)) == 0:
            continue

        buh_only_docs.append({
            'uid': uid,
            'predst': str(sel_buh.Предст),
            'prihod': prihod,
            'rashod': rashod,
            'net': net,
        })
        total_buh_only_net += net

    print(f"  Всього BuhBud документів по рахунку: {total_buh}")
    print(f"  Matched з ERP: {total_matched}")
    print(f"  BuhBud-only (ненульових): {len(buh_only_docs)}")

    # 5. Результат
    print("\n" + "=" * 80)
    print("РЕЗУЛЬТАТ: BuhBud-only документи (виправлена логіка)")
    print("=" * 80)

    for doc in sorted(buh_only_docs, key=lambda x: abs(x['net']), reverse=True):
        mark = " <<<< ШУКАНА РОЗБІЖНІСТЬ!" if abs(doc['net'] + 2824601) < 1 else ""
        print(f"  {doc['predst'][:60]:60s} | П={doc['prihod']:>14,.2f} | Р={doc['rashod']:>14,.2f} | Нетто={doc['net']:>14,.2f}{mark}")

    print(f"\n  РАЗОМ BuhBud-only нетто: {total_buh_only_net:,.2f}")

    # Перевірка: чи знайдено документ 00019117 з сумою -2,824,601
    found_target = any(abs(d['net'] + 2824601) < 1 for d in buh_only_docs)
    print(f"\n  Документ з розбіжністю -2,824,601 знайдено: {'ТАК ✓' if found_target else 'НІ ✗'}")

    if found_target:
        print(f"\n  ВИСНОВОК: Виправлена логіка УСПІШНО знаходить документ-причину розбіжності!")
        print(f"  Стара логіка пропускала його через:")
        print(f"    1. Фільтр ЄПереказ (обидва боки — банківські рахунки)")
        print(f"    2. Нульовий чистий потік по ВСІХ банках")
        print(f"  Нова логіка фільтрує по КОНКРЕТНОМУ рахунку через Субконто1.")
    else:
        print(f"\n  ПОПЕРЕДЖЕННЯ: Документ не знайдено — потрібна додаткова діагностика!")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
