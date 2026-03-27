"""
Тест лютий 2026: порівняння залишків Грошей між звітом та обробкою.

Відтворює логіку:
1. Звіту А_СравнитьОстаткиДенегЕРПсBASБухгалтерия (еталон)
2. Обробки СинхронизироватьТовары (що перевіряємо)

Очікуваний результат за лютий 2026:
- Одна реальна розбіжність: ОТП_ТОВ ІНДАСТРІАЛБУД, расход -293706.85
- Обробка має показати ту саму розбіжність що і звіт
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_BUH = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'

START_DATE = datetime(2026, 2, 1)
END_DATE = datetime(2026, 2, 28)


def get_erp_balances_report(conn_erp):
    """
    Логіка звіту: ПолучитьОстаткиЕРП_Бух (рядки 134-143 звіту).
    Джерело: РегистрНакопления.ДенежныеСредстваБезналичные.ОстаткиИОбороты
    Ключ: БанковскийСчет (reference -> UUID)
    """
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ '
        '   Т.БанковскийСчет КАК БанковскийСчетКасса, '
        '   ПРЕДСТАВЛЕНИЕ(Т.БанковскийСчет) КАК Предст, '
        '   Т.СуммаНачальныйОстаток КАК НачОстаток, '
        '   Т.СуммаПриход КАК Приход, '
        '   Т.СуммаРасход КАК Расход, '
        '   Т.СуммаКонечныйОстаток КАК КонОстаток '
        'ИЗ '
        '   РегистрНакопления.ДенежныеСредстваБезналичные.ОстаткиИОбороты( '
        '       &НачалоПериода, &КонецПериода, , , ) КАК Т'
    )
    q.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    data = {}
    sel = q.Выполнить().Выбрать()
    while sel.Следующий():
        uuid = str(conn_erp.XMLСтрока(sel.БанковскийСчетКасса.УникальныйИдентификатор())).lower()
        data[uuid] = {
            'name': str(sel.Предст),
            'nach': float(sel.НачОстаток or 0),
            'prih': float(sel.Приход or 0),
            'rash': float(sel.Расход or 0),
            'kon': float(sel.КонОстаток or 0),
        }
    return data


def get_buhbud_balances_report(V83, conn_erp):
    """
    Логіка звіту: ПолучитьОстаткиBuhBud (рядки 231-326 звіту).
    Джерело: РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты з 7 параметрами
    Ключ: Субконто1 -> UUID через V83.String(row.УникальныйИдентификатор())
    Спец. обробка рахунку 315 по ЄДРПОУ
    Пряма ітерація по COM ТЗ (як у звіті)
    """
    q = V83.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ '
        '   Остатки.Субконто1 КАК БанковскийСчетКасса, '
        '   ПРЕДСТАВЛЕНИЕ(Остатки.Субконто1) КАК БанковскийСчетПредставление, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаНачальныйОстаток '
        '       ИНАЧЕ Остатки.СуммаНачальныйОстаток '
        '   КОНЕЦ КАК НачОстаток, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаОборотДт '
        '       ИНАЧЕ Остатки.СуммаОборотДт '
        '   КОНЕЦ КАК Приход, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаОборотКт '
        '       ИНАЧЕ Остатки.СуммаОборотКт '
        '   КОНЕЦ КАК Расход, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаКонечныйОстаток '
        '       ИНАЧЕ Остатки.СуммаКонечныйОстаток '
        '   КОНЕЦ КАК КонОстаток, '
        '   Остатки.Счет.Код КАК СчетКод, '
        '   Остатки.Организация.КодПоЕДРПОУ КАК ОрганизацияКодПоЕДРПОУ '
        'ИЗ '
        '   РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты( '
        '       &НачалоПериода, &КонецПериода, , , '
        '       Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , ) КАК Остатки '
        'ГДЕ '
        '   (Остатки.СуммаНачальныйОстаток <> 0 '
        '       ИЛИ Остатки.СуммаОборотДт <> 0 '
        '       ИЛИ Остатки.СуммаОборотКт <> 0)'
    )
    q.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    ТЗ = q.Выполнить().Выгрузить()

    data = {}
    uuid_ok = 0
    uuid_fail = 0
    acc315 = 0

    for i in range(ТЗ.Количество()):
        row = ТЗ.Получить(i)
        name = str(V83.String(row.БанковскийСчетПредставление)).strip()
        code = str(V83.String(row.СчетКод)).strip()
        nach = float(row.НачОстаток or 0)
        prih = float(row.Приход or 0)
        rash = float(row.Расход or 0)
        kon = float(row.КонОстаток or 0)

        uuid = ""
        try:
            if code == "315":
                # Спец. обробка 315 — пошук в ERP по ЄДРПОУ (як в звіті рядок 290-307)
                edrpou = str(V83.String(row.ОрганизацияКодПоЕДРПОУ)).strip()
                q315 = conn_erp.NewObject("Запрос")
                q315.Текст = (
                    'ВЫБРАТЬ '
                    '   БСО.Ссылка КАК Ссылка '
                    'ИЗ '
                    '   Справочник.БанковскиеСчетаОрганизаций КАК БСО '
                    'ГДЕ '
                    '   БСО.А_ОрганизацияБухгалтерия.КодПоЕДРПОУ = &КодПоЕДРПОУ '
                    '   И БСО.А_СчетБухгалтерский.Код = "315" '
                    '   И БСО.А_ПоУмолчаниюДляПустыхСчетовИзБухгалтерии'
                )
                q315.УстановитьПараметр("КодПоЕДРПОУ", edrpou)
                sel315 = q315.Выполнить().Выбрать()
                if sel315.Следующий():
                    uuid = str(conn_erp.XMLСтрока(sel315.Ссылка.УникальныйИдентификатор())).lower()
                    acc315 += 1
            elif V83.ЗначениеЗаполнено(row.БанковскийСчетКасса):
                uuid = str(V83.String(row.БанковскийСчетКасса.УникальныйИдентификатор())).lower()
                uuid_ok += 1
        except Exception as e:
            uuid_fail += 1
            print(f"    UUID FAIL [{i+1}] {name} (сч.{code}): {e}")

        if uuid:
            if uuid not in data:
                data[uuid] = {'name': name, 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0}
            data[uuid]['nach'] += nach
            data[uuid]['prih'] += prih
            data[uuid]['rash'] += rash
            data[uuid]['kon'] += kon

    print(f"    BuhBud (звіт): {ТЗ.Количество()} рядків, UUID OK={uuid_ok}, 315={acc315}, FAIL={uuid_fail}")
    return data


def get_erp_balances_processor(conn_erp):
    """
    Логіка обробки: ПолучитьОстаткиЕРП_Деньги (рядки 431-471).
    Те саме що звіт але UUID через Строка(УникальныйИдентификатор())
    """
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ '
        '   Т.БанковскийСчет КАК БанковскийСчет, '
        '   ПРЕДСТАВЛЕНИЕ(Т.БанковскийСчет) КАК Предст, '
        '   Т.СуммаНачальныйОстаток КАК НачОстаток, '
        '   Т.СуммаПриход КАК Приход, '
        '   Т.СуммаРасход КАК Расход, '
        '   Т.СуммаКонечныйОстаток КАК КонОстаток '
        'ИЗ '
        '   РегистрНакопления.ДенежныеСредстваБезналичные.ОстаткиИОбороты( '
        '       &НачалоПериода, &КонецПериода, , , ) КАК Т'
    )
    q.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    data = {}
    sel = q.Выполнить().Выбрать()
    while sel.Следующий():
        # Обробка використовує Строка() а не XMLСтрока()
        uuid = str(conn_erp.XMLСтрока(sel.БанковскийСчет.УникальныйИдентификатор())).lower()
        data[uuid] = {
            'name': str(sel.Предст),
            'nach': float(sel.НачОстаток or 0),
            'prih': float(sel.Приход or 0),
            'rash': float(sel.Расход or 0),
            'kon': float(sel.КонОстаток or 0),
        }
    return data


def get_buhbud_balances_processor(V83, conn_erp):
    """
    Логіка обробки: ПолучитьОстаткиBuhBud_Деньги (рядки 473-609).
    Відмінність від звіту: фільтр по масКодПоЕДРПОУ (організації з А_ВБалансе=ІСТИНА)
    Пряма COM ітерація (після виправлення, рядки 561-605)
    """
    # Крок 1: отримати список організацій з А_ВБалансе
    q_org = conn_erp.NewObject("Запрос")
    q_org.Текст = (
        'ВЫБРАТЬ '
        '   Организации.КодПоЕДРПОУ КАК КодПоЕДРПОУ '
        'ИЗ '
        '   Справочник.Организации КАК Организации '
        'ГДЕ '
        '   Организации.А_ВБалансе = ИСТИНА '
        '   И Организации.КодПоЕДРПОУ <> ""'
    )
    edrpou_list = []
    sel_org = q_org.Выполнить().Выбрать()
    while sel_org.Следующий():
        edrpou_list.append(str(sel_org.КодПоЕДРПОУ))
    print(f"    Організації А_ВБалансе: {len(edrpou_list)} ({', '.join(edrpou_list[:5])}...)")

    # Крок 2: BuhBud запит з фільтром
    масКодПоЕДРПОУ_V83 = V83.NewObject("Массив")
    for code in edrpou_list:
        масКодПоЕДРПОУ_V83.Добавить(code)

    q = V83.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ '
        '   Остатки.Субконто1 КАК БанковскийСчетКасса, '
        '   ПРЕДСТАВЛЕНИЕ(Остатки.Субконто1) КАК БанковскийСчетПредставление, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаНачальныйОстаток '
        '       ИНАЧЕ Остатки.СуммаНачальныйОстаток '
        '   КОНЕЦ КАК НачОстаток, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаОборотДт '
        '       ИНАЧЕ Остатки.СуммаОборотДт '
        '   КОНЕЦ КАК Приход, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаОборотКт '
        '       ИНАЧЕ Остатки.СуммаОборотКт '
        '   КОНЕЦ КАК Расход, '
        '   ВЫБОР '
        '       КОГДА Остатки.Счет.Валютный '
        '           ТОГДА Остатки.ВалютнаяСуммаКонечныйОстаток '
        '       ИНАЧЕ Остатки.СуммаКонечныйОстаток '
        '   КОНЕЦ КАК КонОстаток, '
        '   Остатки.Счет.Код КАК СчетКод, '
        '   Остатки.Организация.КодПоЕДРПОУ КАК ОрганизацияКодПоЕДРПОУ '
        'ИЗ '
        '   РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты( '
        '       &НачалоПериода, &КонецПериода, , , '
        '       Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , ) КАК Остатки '
        'ГДЕ '
        '   Остатки.Организация.КодПоЕДРПОУ В (&масКодПоЕДРПОУ) '
        '   И (Остатки.СуммаНачальныйОстаток <> 0 '
        '       ИЛИ Остатки.СуммаОборотДт <> 0 '
        '       ИЛИ Остатки.СуммаОборотКт <> 0)'
    )
    q.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))
    q.УстановитьПараметр("масКодПоЕДРПОУ", масКодПоЕДРПОУ_V83)

    ТЗ = q.Выполнить().Выгрузить()

    data = {}
    uuid_ok = 0
    uuid_fail = 0
    acc315 = 0

    for i in range(ТЗ.Количество()):
        row = ТЗ.Получить(i)
        name = str(V83.String(row.БанковскийСчетПредставление)).strip()
        code = str(V83.String(row.СчетКод)).strip()
        nach = float(row.НачОстаток or 0)
        prih = float(row.Приход or 0)
        rash = float(row.Расход or 0)
        kon = float(row.КонОстаток or 0)

        uuid = ""
        try:
            if code == "315":
                # Спец. обробка 315 — пошук в ERP по ЄДРПОУ (рядки 572-589 обробки)
                edrpou = str(V83.String(row.ОрганизацияКодПоЕДРПОУ)).strip()
                q315 = conn_erp.NewObject("Запрос")
                q315.Текст = (
                    'ВЫБРАТЬ '
                    '   БСО.Ссылка КАК Ссылка '
                    'ИЗ '
                    '   Справочник.БанковскиеСчетаОрганизаций КАК БСО '
                    'ГДЕ '
                    '   БСО.А_ОрганизацияБухгалтерия.КодПоЕДРПОУ = &КодПоЕДРПОУ '
                    '   И БСО.А_СчетБухгалтерский.Код = "315" '
                    '   И БСО.А_ПоУмолчаниюДляПустыхСчетовИзБухгалтерии'
                )
                q315.УстановитьПараметр("КодПоЕДРПОУ", edrpou)
                sel315 = q315.Выполнить().Выбрать()
                if sel315.Следующий():
                    uuid = str(conn_erp.XMLСтрока(sel315.Ссылка.УникальныйИдентификатор())).lower()
                    acc315 += 1
            elif V83.ЗначениеЗаполнено(row.БанковскийСчетКасса):
                # Як в обробці рядок 592-598
                buh_uuid = str(V83.String(row.БанковскийСчетКасса.УникальныйИдентификатор())).lower()
                uuid = buh_uuid  # ПолучитьСправочникUID повертає ту саму UUID
                uuid_ok += 1
        except Exception as e:
            uuid_fail += 1
            print(f"    UUID FAIL [{i+1}] {name} (сч.{code}): {e}")

        if uuid:
            if uuid not in data:
                data[uuid] = {'name': name, 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0}
            data[uuid]['nach'] += nach
            data[uuid]['prih'] += prih
            data[uuid]['rash'] += rash
            data[uuid]['kon'] += kon

    print(f"    BuhBud (обробка): {ТЗ.Количество()} рядків, UUID OK={uuid_ok}, 315={acc315}, FAIL={uuid_fail}")
    return data


def reconcile(erp_data, buh_data, label):
    """Зведення: порівняння ERP та BuhBud залишків по UUID."""
    all_keys = set(erp_data.keys()) | set(buh_data.keys())
    discrepancies = []
    matches = 0

    for uuid in sorted(all_keys):
        erp = erp_data.get(uuid, {'name': '??? (тільки BuhBud)', 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0})
        buh = buh_data.get(uuid, {'name': '??? (тільки ERP)', 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0})

        diff_nach = round(erp['nach'] - buh['nach'], 2)
        diff_prih = round(erp['prih'] - buh['prih'], 2)
        diff_rash = round(erp['rash'] - buh['rash'], 2)
        diff_kon = round(erp['kon'] - buh['kon'], 2)

        if diff_nach == 0 and diff_prih == 0 and diff_rash == 0 and diff_kon == 0:
            matches += 1
        else:
            name = erp['name'] if uuid in erp_data else buh['name']
            discrepancies.append({
                'uuid': uuid,
                'name': name,
                'erp': erp,
                'buh': buh,
                'diff_nach': diff_nach,
                'diff_prih': diff_prih,
                'diff_rash': diff_rash,
                'diff_kon': diff_kon,
            })

    return matches, discrepancies


def print_results(label, matches, discrepancies):
    """Вивід результатів зведення."""
    print(f"\n{'=' * 80}")
    print(f"  {label}")
    print(f"{'=' * 80}")
    print(f"  Збігів: {matches}, Розбіжностей: {len(discrepancies)}")

    if discrepancies:
        print(f"\n  {'Рахунок':<55} {'НачОст':>12} {'Приход':>14} {'Расход':>14} {'КонОст':>12}")
        print(f"  {'-'*55} {'-'*12} {'-'*14} {'-'*14} {'-'*12}")

        for d in discrepancies:
            print(f"\n  {d['name'][:55]}")
            print(f"    ERP:  {d['erp']['nach']:>12,.2f} {d['erp']['prih']:>14,.2f} {d['erp']['rash']:>14,.2f} {d['erp']['kon']:>12,.2f}")
            print(f"    BUH:  {d['buh']['nach']:>12,.2f} {d['buh']['prih']:>14,.2f} {d['buh']['rash']:>14,.2f} {d['buh']['kon']:>12,.2f}")
            print(f"    DIFF: {d['diff_nach']:>12,.2f} {d['diff_prih']:>14,.2f} {d['diff_rash']:>14,.2f} {d['diff_kon']:>12,.2f}")
    else:
        print("  Всі рахунки збігаються!")

    return discrepancies


def test_feb2026():
    """Головний тест: порівняння звіту та обробки за лютий 2026."""
    print("=" * 80)
    print(f"ТЕСТ: Порівняння залишків Грошей за лютий 2026")
    print(f"Період: {START_DATE.strftime('%d.%m.%Y')} - {END_DATE.strftime('%d.%m.%Y')}")
    print("=" * 80)

    pythoncom.CoInitialize()
    v8_conn = win32com.client.Dispatch("V83.COMConnector")

    # Підключення
    conn_erp = v8_conn.Connect(CONN_ERP)
    print("OK: ERP підключено")
    V83 = v8_conn.Connect(CONN_BUH)
    print("OK: BuhBud підключено")

    # ==========================================
    # ЧАСТИНА 1: Логіка ЗВІТУ (еталон)
    # ==========================================
    print("\n--- ЧАСТИНА 1: Логіка звіту А_СравнитьОстаткиДенегЕРПсBASБухгалтерия ---")

    print("  ERP (звіт)...")
    erp_report = get_erp_balances_report(conn_erp)
    print(f"    ERP: {len(erp_report)} рахунків")

    print("  BuhBud (звіт)...")
    buh_report = get_buhbud_balances_report(V83, conn_erp)
    print(f"    BuhBud: {len(buh_report)} рахунків")

    matches_r, disc_r = reconcile(erp_report, buh_report, "ЗВІТ")
    print_results("ЗВІТ (еталон): А_СравнитьОстаткиДенегЕРПсBASБухгалтерия", matches_r, disc_r)

    # ==========================================
    # ЧАСТИНА 2: Логіка ОБРОБКИ (що перевіряємо)
    # ==========================================
    print("\n--- ЧАСТИНА 2: Логіка обробки СинхронизироватьТовары ---")

    print("  ERP (обробка)...")
    erp_proc = get_erp_balances_processor(conn_erp)
    print(f"    ERP: {len(erp_proc)} рахунків")

    print("  BuhBud (обробка)...")
    buh_proc = get_buhbud_balances_processor(V83, conn_erp)
    print(f"    BuhBud: {len(buh_proc)} рахунків")

    matches_p, disc_p = reconcile(erp_proc, buh_proc, "ОБРОБКА")
    print_results("ОБРОБКА: СинхронизироватьТовары (виправлена логіка)", matches_p, disc_p)

    # ==========================================
    # ЧАСТИНА 3: Порівняння звіту та обробки
    # ==========================================
    print("\n" + "=" * 80)
    print("  ПОРІВНЯННЯ: Звіт vs Обробка")
    print("=" * 80)

    # Перевіряємо чи розбіжності однакові
    disc_r_uuids = {d['uuid'] for d in disc_r}
    disc_p_uuids = {d['uuid'] for d in disc_p}

    only_in_report = disc_r_uuids - disc_p_uuids
    only_in_proc = disc_p_uuids - disc_r_uuids
    in_both = disc_r_uuids & disc_p_uuids

    if only_in_report:
        print(f"\n  Розбіжності ТІЛЬКИ у звіті ({len(only_in_report)}):")
        for d in disc_r:
            if d['uuid'] in only_in_report:
                print(f"    - {d['name']}: diff_rash={d['diff_rash']}, diff_kon={d['diff_kon']}")

    if only_in_proc:
        print(f"\n  Розбіжності ТІЛЬКИ в обробці ({len(only_in_proc)}):")
        for d in disc_p:
            if d['uuid'] in only_in_proc:
                print(f"    - {d['name']}: diff_rash={d['diff_rash']}, diff_kon={d['diff_kon']}")

    if in_both:
        print(f"\n  Спільні розбіжності ({len(in_both)}):")
        for uuid in in_both:
            dr = next(d for d in disc_r if d['uuid'] == uuid)
            dp = next(d for d in disc_p if d['uuid'] == uuid)
            match_ok = (dr['diff_nach'] == dp['diff_nach'] and
                        dr['diff_prih'] == dp['diff_prih'] and
                        dr['diff_rash'] == dp['diff_rash'] and
                        dr['diff_kon'] == dp['diff_kon'])
            status = "ЗБІГ" if match_ok else "РІЗНИЦЯ"
            print(f"    [{status}] {dr['name']}")
            if not match_ok:
                print(f"      Звіт:    diff_rash={dr['diff_rash']}, diff_kon={dr['diff_kon']}")
                print(f"      Обробка: diff_rash={dp['diff_rash']}, diff_kon={dp['diff_kon']}")

    if not only_in_report and not only_in_proc:
        if len(disc_r) == len(disc_p):
            print(f"\n  РЕЗУЛЬТАТ: Звіт і обробка показують ОДНАКОВІ розбіжності ({len(disc_r)} шт)")
        else:
            print(f"\n  УВАГА: Кількість розбіжностей різна! Звіт={len(disc_r)}, Обробка={len(disc_p)}")
    else:
        print(f"\n  УВАГА: Є розбіжності які бачить тільки один з них!")

    # Очікуваний результат
    print("\n" + "-" * 80)
    print("  ОЧІКУВАНІ дані (з наданого звіту за лютий 2026):")
    print("  ОТП_ТОВ ІНДАСТРІАЛБУД UA97...: diff_расход = -293706.85, diff_кон = 293706.85")
    print("-" * 80)

    V83 = None
    conn_erp = None
    v8_conn = None


if __name__ == "__main__":
    try:
        test_feb2026()
    except Exception as e:
        print(f"FAIL: {e}")
        import traceback
        traceback.print_exc()
