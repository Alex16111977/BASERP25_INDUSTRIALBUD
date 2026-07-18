# -*- coding: utf-8 -*-
"""
Діагностика: чому обробка не знаходить документ-причину розбіжності 2,824,601
на рахунку "ТАС_Будівн_ Індастріалбуд (UA693395002600501537072000001)" за грудень 2025.
"""

import win32com.client
import pythoncom
import pywintypes
from datetime import datetime


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    print("=" * 80)
    print("ДІАГНОСТИКА: Розбіжність 2,824,601 на ТАС_Будівн за грудень 2025")
    print("=" * 80)

    conn_erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    conn_buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

    dt_start = pywintypes.Time(datetime(2025, 12, 1))
    dt_end = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    # ===================================================================
    # КРОК 1: Знайти банківський рахунок ТАС_Будівн в ЕРП
    # (вибираємо ПЕРШИЙ — основний, без $$$)
    # ===================================================================
    print("\n--- КРОК 1: Знайти банківський рахунок ---")
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Ссылка, Наименование, НомерСчета "
        "ИЗ Справочник.БанковскиеСчетаОрганизаций "
        "ГДЕ НомерСчета ПОДОБНО &Номер "
        "УПОРЯДОЧИТЬ ПО Наименование"
    )
    q.УстановитьПараметр("Номер", "%UA693395002600501537072000001%")
    sel = q.Выполнить().Выбрать()
    all_banks = []
    while sel.Следующий():
        uid = str(conn_erp.XMLСтрока(sel.Ссылка.УникальныйИдентификатор())).upper()
        name = str(sel.Наименование)
        all_banks.append({'ref': sel.Ссылка, 'uid': uid, 'name': name})
        print(f"  [{len(all_banks)}] {name} | UUID: {uid}")

    # Вибираємо рахунок "ТАС_Будівн_" (без $$$)
    bank_ref = None
    bank_uid = ""
    for b in all_banks:
        if "ТАС_Будівн_" in b['name'] and "$$$" not in b['name']:
            bank_ref = b['ref']
            bank_uid = b['uid']
            print(f"\n  >>> Вибрано: {b['name']}")
            break
    if not bank_ref and all_banks:
        bank_ref = all_banks[0]['ref']
        bank_uid = all_banks[0]['uid']
        print(f"\n  >>> Вибрано (перший): {all_banks[0]['name']}")

    if not bank_ref:
        print("  РАХУНОК НЕ ЗНАЙДЕНО!")
        return

    # ===================================================================
    # КРОК 2: ЕРП — всі документи по ТАС_Будівн за грудень 2025
    # ===================================================================
    print("\n--- КРОК 2: ЕРП документи (ДенежныеСредстваБезналичные.Обороты) ---")
    q2 = conn_erp.NewObject("Запрос")
    q2.Текст = (
        "ВЫБРАТЬ "
        "  Об.Регистратор КАК Регистратор, "
        "  ПРЕДСТАВЛЕНИЕ(Об.Регистратор) КАК Предст, "
        "  Об.СуммаПриход КАК Приход, "
        "  Об.СуммаРасход КАК Расход, "
        "  Соотв.УникальныйИдентификаторПриемника КАК UIDвBuhBud, "
        "  Соотв.ТипПриемника КАК ТипПриемника "
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
    q2.УстановитьПараметр("БанковскийСчет", bank_ref)
    sel2 = q2.Выполнить().Выбрать()

    erp_docs = {}
    total_erp_prihod = 0
    total_erp_rashod = 0
    transfer_count = 0

    while sel2.Следующий():
        uid = str(conn_erp.XMLСтрока(sel2.Регистратор.УникальныйИдентификатор())).upper()
        prihod = float(sel2.Приход) if sel2.Приход else 0
        rashod = float(sel2.Расход) if sel2.Расход else 0
        total_erp_prihod += prihod
        total_erp_rashod += rashod

        uid_buh = ""
        try:
            if sel2.UIDвBuhBud:
                uid_buh = str(sel2.UIDвBuhBud).upper()
        except:
            pass

        # Перевірити тип приемника
        typ_priem = ""
        try:
            if sel2.ТипПриемника:
                typ_priem = str(sel2.ТипПриемника)
        except:
            pass

        # Перевірити хозоперацію
        hoz_op = ""
        is_transfer = False
        try:
            ho = sel2.Регистратор.ХозяйственнаяОперация
            hoz_op = str(conn_erp.String(ho))
            if "ДругогоСчета" in hoz_op or "ДругойСчет" in hoz_op:
                is_transfer = True
                transfer_count += 1
        except:
            pass

        erp_docs[uid] = {
            'prihod': prihod,
            'rashod': rashod,
            'predst': str(sel2.Предст),
            'uid_buh': uid_buh,
            'typ_priem': typ_priem,
            'hoz_op': hoz_op,
            'is_transfer': is_transfer,
        }

    print(f"  Всього документів ЕРП: {len(erp_docs)}")
    print(f"  З них переказів: {transfer_count}")
    print(f"  Приход: {total_erp_prihod:,.2f}")
    print(f"  Расход: {total_erp_rashod:,.2f}")
    print(f"  Нетто: {total_erp_prihod - total_erp_rashod:,.2f}")

    # ===================================================================
    # КРОК 3: BuhBud — знайти рахунок і отримати документи
    # ===================================================================
    print("\n--- КРОК 3: BuhBud документи ---")

    # 3a. Знайти рахунок BuhBud (вибираємо "ТАС_Будівн_" без $$$)
    q_buh_acc = conn_buh.NewObject("Запрос")
    q_buh_acc.Текст = (
        "ВЫБРАТЬ Ссылка, Наименование, НомерСчета "
        "ИЗ Справочник.БанковскиеСчета "
        "ГДЕ НомерСчета ПОДОБНО &Номер "
        "УПОРЯДОЧИТЬ ПО Наименование"
    )
    q_buh_acc.УстановитьПараметр("Номер", "%UA693395002600501537072000001%")
    sel_acc = q_buh_acc.Выполнить().Выбрать()
    buh_banks = []
    while sel_acc.Следующий():
        name = str(sel_acc.Наименование)
        buh_banks.append({'ref': sel_acc.Ссылка, 'name': name})
        print(f"  BuhBud [{len(buh_banks)}] '{name}' | Номер: {sel_acc.НомерСчета}")

    buh_bank_ref = None
    for b in buh_banks:
        if "ТАС_Будівн_" in b['name'] and "$$$" not in b['name']:
            buh_bank_ref = b['ref']
            print(f"\n  >>> Вибрано BuhBud: {b['name']}")
            break
    if not buh_bank_ref and buh_banks:
        buh_bank_ref = buh_banks[0]['ref']
        print(f"\n  >>> Вибрано BuhBud (перший): {buh_banks[0]['name']}")

    # 3b. Загальні цифри через ОстаткиИОбороты (з Субконто1)
    print("\n  BuhBud ОстаткиИОбороты:")
    q_boh = conn_buh.NewObject("Запрос")
    q_boh.Текст = (
        "ВЫБРАТЬ "
        "  Т.СуммаОборотДт КАК Приход, "
        "  Т.СуммаОборотКт КАК Расход, "
        "  Т.СуммаНачальныйОстаток КАК НачОст, "
        "  Т.СуммаКонечныйОстаток КАК КонОст "
        "ИЗ "
        "  РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты("
        "    &НачалоПериода, &КонецПериода, , , "
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , "
        "    Субконто1 = &БанкРахунок) КАК Т"
    )
    q_boh.УстановитьПараметр("НачалоПериода", dt_start)
    q_boh.УстановитьПараметр("КонецПериода", dt_end)
    q_boh.УстановитьПараметр("БанкРахунок", buh_bank_ref)
    sel_boh = q_boh.Выполнить().Выбрать()
    buh_total_prihod = 0
    buh_total_rashod = 0
    while sel_boh.Следующий():
        buh_total_prihod = float(sel_boh.Приход)
        buh_total_rashod = float(sel_boh.Расход)
        print(f"    НачОст: {float(sel_boh.НачОст):,.2f}")
        print(f"    Приход (Дт): {buh_total_prihod:,.2f}")
        print(f"    Расход (Кт): {buh_total_rashod:,.2f}")
        print(f"    КонОст: {float(sel_boh.КонОст):,.2f}")

    # 3c. Документи BuhBud через ОстаткиИОбороты з Периодичность=Регистратор
    print("\n  Запит BuhBud документів (ОстаткиИОбороты, Периодичність=Регистратор)...")
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

    buh_docs = {}
    total_buh_docs_p = 0
    total_buh_docs_r = 0
    try:
        sel_bd = q_buh_docs.Выполнить().Выбрать()
        while sel_bd.Следующий():
            try:
                reg = sel_bd.Регистратор
                if reg is None or not conn_buh.ЗначениеЗаполнено(reg):
                    continue
                uid = str(conn_buh.XMLСтрока(reg.УникальныйИдентификатор())).upper()
            except:
                continue
            prihod = float(sel_bd.Приход) if sel_bd.Приход else 0
            rashod = float(sel_bd.Расход) if sel_bd.Расход else 0
            total_buh_docs_p += prihod
            total_buh_docs_r += rashod
            buh_docs[uid] = {
                'prihod': prihod,
                'rashod': rashod,
                'predst': str(sel_bd.Предст),
            }
    except Exception as e:
        print(f"  ПОМИЛКА: {e}")
        import traceback; traceback.print_exc()
        return

    print(f"  Всього документів BuhBud: {len(buh_docs)}")
    print(f"  Приход: {total_buh_docs_p:,.2f}")
    print(f"  Расход: {total_buh_docs_r:,.2f}")
    print(f"  Нетто: {total_buh_docs_p - total_buh_docs_r:,.2f}")

    # ===================================================================
    # КРОК 4: Побудова відповідності UUID ЕРП ↔ BuhBud
    # ===================================================================
    print("\n" + "=" * 80)
    print("КРОК 4: Порівняння документів ERP ↔ BuhBud")
    print("=" * 80)

    erp_to_buh = {}
    buh_to_erp = {}
    for uid_erp, doc in erp_docs.items():
        if doc['uid_buh']:
            erp_to_buh[uid_erp] = doc['uid_buh']
            buh_to_erp[doc['uid_buh']] = uid_erp

    # 4a. Документи тільки в BuhBud
    buh_only = []
    for uid_buh, doc in buh_docs.items():
        if uid_buh not in buh_to_erp:
            buh_only.append((uid_buh, doc))

    print(f"\n--- 4a. Документи тільки в BuhBud: {len(buh_only)} ---")
    sum_buh_only_p = 0
    sum_buh_only_r = 0
    for uid, doc in sorted(buh_only, key=lambda x: abs(x[1]['prihod'] - x[1]['rashod']), reverse=True):
        net = doc['prihod'] - doc['rashod']
        sum_buh_only_p += doc['prihod']
        sum_buh_only_r += doc['rashod']
        print(f"  {doc['predst'][:60]:60s} | П={doc['prihod']:>14,.2f} | Р={doc['rashod']:>14,.2f} | Н={net:>14,.2f}")
    print(f"  РАЗОМ BuhBud-only: Приход={sum_buh_only_p:,.2f} Расход={sum_buh_only_r:,.2f}")

    # 4b. Документи тільки в ЕРП
    erp_only = []
    for uid_erp, doc in erp_docs.items():
        uid_buh = erp_to_buh.get(uid_erp, "")
        if uid_buh and uid_buh in buh_docs:
            continue  # matched
        erp_only.append((uid_erp, doc))

    print(f"\n--- 4b. Документи тільки в ЕРП: {len(erp_only)} ---")
    for uid, doc in erp_only[:20]:
        net = doc['prihod'] - doc['rashod']
        reason = "no_corr" if not doc['uid_buh'] else "corr_but_no_buh_movement"
        transfer_mark = " [ПЕРЕКАЗ]" if doc['is_transfer'] else ""
        print(f"  {doc['predst'][:55]:55s}{transfer_mark} | П={doc['prihod']:>14,.2f} | Р={doc['rashod']:>14,.2f} | {reason}")

    # 4c. Matched documents with amount differences
    diff_docs = []
    for uid_erp, doc_erp in erp_docs.items():
        uid_buh = erp_to_buh.get(uid_erp, "")
        if not uid_buh or uid_buh not in buh_docs:
            continue
        doc_buh = buh_docs[uid_buh]
        dp = doc_erp['prihod'] - doc_buh['prihod']
        dr = doc_erp['rashod'] - doc_buh['rashod']
        if abs(dp) > 0.01 or abs(dr) > 0.01:
            diff_docs.append((uid_erp, doc_erp, doc_buh, dp, dr))

    print(f"\n--- 4c. Документи з різницею в сумах: {len(diff_docs)} ---")
    sum_diff_p = 0
    sum_diff_r = 0
    for uid, doc_erp, doc_buh, dp, dr in diff_docs:
        sum_diff_p += dp
        sum_diff_r += dr
        transfer_mark = " [ПЕРЕКАЗ]" if doc_erp['is_transfer'] else ""
        print(f"  {doc_erp['predst'][:55]:55s}{transfer_mark}")
        print(f"    ЕРП:    П={doc_erp['prihod']:>14,.2f}  Р={doc_erp['rashod']:>14,.2f}")
        print(f"    BuhBud: П={doc_buh['prihod']:>14,.2f}  Р={doc_buh['rashod']:>14,.2f}")
        print(f"    Різн:   dP={dp:>14,.2f}  dR={dr:>14,.2f}")
    print(f"  РАЗОМ різниця: dP={sum_diff_p:,.2f} dR={sum_diff_r:,.2f}")

    # 4d. Переказ-документи окремо
    transfers = [(u, d) for u, d in erp_docs.items() if d['is_transfer']]
    print(f"\n--- 4d. Переказ-документи: {len(transfers)} ---")
    for uid, doc in transfers[:20]:
        uid_buh = erp_to_buh.get(uid, "")
        matched = uid_buh in buh_docs if uid_buh else False
        buh_info = ""
        if matched:
            db = buh_docs[uid_buh]
            buh_info = f" BuhBud: П={db['prihod']:>12,.2f} Р={db['rashod']:>12,.2f}"
        print(f"  {doc['predst'][:50]:50s} ЕРП: П={doc['prihod']:>12,.2f} Р={doc['rashod']:>12,.2f} |{buh_info} | {'matched' if matched else 'NOT matched'}")

    # ===================================================================
    # КРОК 5: Пояснення розбіжності
    # ===================================================================
    print("\n" + "=" * 80)
    print("КРОК 5: Баланс розбіжності")
    print("=" * 80)

    # Matched OK
    matched_erp_p = 0
    matched_erp_r = 0
    matched_buh_p = 0
    matched_buh_r = 0
    for uid_erp, doc_erp in erp_docs.items():
        uid_buh = erp_to_buh.get(uid_erp, "")
        if uid_buh and uid_buh in buh_docs:
            matched_erp_p += doc_erp['prihod']
            matched_erp_r += doc_erp['rashod']
            doc_buh = buh_docs[uid_buh]
            matched_buh_p += doc_buh['prihod']
            matched_buh_r += doc_buh['rashod']

    erp_only_p = sum(d['prihod'] for _, d in erp_only)
    erp_only_r = sum(d['rashod'] for _, d in erp_only)

    print(f"\n  ЕРП загалом:          П={total_erp_prihod:>16,.2f}  Р={total_erp_rashod:>16,.2f}")
    print(f"  BuhBud загалом:       П={total_buh_docs_p:>16,.2f}  Р={total_buh_docs_r:>16,.2f}")
    print(f"  Різниця (BuhBud-ЕРП): П={total_buh_docs_p-total_erp_prihod:>16,.2f}  Р={total_buh_docs_r-total_erp_rashod:>16,.2f}")
    print()
    print(f"  Matched ЕРП:       П={matched_erp_p:>16,.2f}  Р={matched_erp_r:>16,.2f}")
    print(f"  Matched BuhBud:    П={matched_buh_p:>16,.2f}  Р={matched_buh_r:>16,.2f}")
    print(f"  Matched різниця:   П={matched_buh_p-matched_erp_p:>16,.2f}  Р={matched_buh_r-matched_erp_r:>16,.2f}")
    print()
    print(f"  ERP-only:          П={erp_only_p:>16,.2f}  Р={erp_only_r:>16,.2f}")
    print(f"  BuhBud-only:       П={sum_buh_only_p:>16,.2f}  Р={sum_buh_only_r:>16,.2f}")

    # ===================================================================
    # КРОК 6: Як обробка бачить ці документи (імітація)
    # ===================================================================
    print("\n" + "=" * 80)
    print("КРОК 6: Як обробка бачить документи")
    print("=" * 80)

    # Обробка для кожного ЕРП документа:
    # 1. Якщо ХозОперація = переказ → "Переказ між рахунками", пропустити
    # 2. Якщо є відповідність → перевірити в BuhBud:
    #    КоличествоБух = Сума(Дт банк) - Сума(Кт банк) ПО ВСІХ банках
    #    Якщо КоличествоЕРП ≈ КоличествоБух → ОК
    # 3. Якщо нема відповідності → "Немає в BuhBud"

    for uid_erp, doc in erp_docs.items():
        status = ""
        if doc['is_transfer']:
            status = "Переказ між рахунками"
        elif not doc['uid_buh']:
            status = "Немає в BuhBud (немає відповідності)"
        else:
            uid_buh = doc['uid_buh']
            # Обробка рахує КоличествоБух по ВСІХ банківських рахунках (не тільки ТАС_Будівн)
            # Ми перевіряємо: чи дає запит по ВСІХ рахунках інший результат?
            if uid_buh in buh_docs:
                net_erp = doc['prihod'] - doc['rashod']
                net_buh_specific = buh_docs[uid_buh]['prihod'] - buh_docs[uid_buh]['rashod']
                # Тепер запит по ВСІХ банківських рахунках
                try:
                    typ = doc['typ_priem'].replace("ДокументСсылка.", "") if doc['typ_priem'] else ""
                    if typ:
                        buh_doc_ref = conn_buh.Документы[typ].ПолучитьСсылку(
                            conn_buh.NewObject("УникальныйИдентификатор", uid_buh))
                        q_all = conn_buh.NewObject("Запрос")
                        q_all.Текст = (
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
                        q_all.УстановитьПараметр("Регистратор", buh_doc_ref)
                        sel_all = q_all.Выполнить().Выбрать()
                        net_buh_all = 0
                        if sel_all.Следующий():
                            net_buh_all = float(sel_all.Сумма) if sel_all.Сумма else 0

                        if abs(round(net_erp - net_buh_all, 2)) == 0:
                            status = "ОК (обробка)"
                        else:
                            status = f"Різниця! ЕРП={net_erp:,.2f} BuhBudAll={net_buh_all:,.2f} BuhBudSpec={net_buh_specific:,.2f}"
                    else:
                        status = "Тип приемника не визначено"
                except Exception as e:
                    status = f"Помилка: {e}"
            else:
                status = f"uid_buh={uid_buh} не знайдено в BuhBud рухах"

        transfer_mark = " [ПЕРЕКАЗ]" if doc['is_transfer'] else ""
        short = doc['predst'][:50]
        print(f"  {short:50s}{transfer_mark:12s} | {status}")

    pythoncom.CoUninitialize()
    print("\n  ГОТОВО")


if __name__ == "__main__":
    main()
