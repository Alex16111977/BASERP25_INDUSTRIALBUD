"""
Діагностичний тест: порівняння залишків Грошей між ERP та BuhBud.

Мета: з'ясувати чому обробка СинхронизироватьТовары показує хибні розбіжності
(BuhBud колонки = 0), а звіт А_СравнитьОстаткиДенегЕРПсBASБухгалтерия — 0 розбіжностей.

Гіпотези:
1. Запит ОстаткиИОбороты з 7 параметрами НЕ працює через COM (Попытка ловить помилку)
2. Серіалізація ЗначениеВСтрокуВнутр/ЗначениеИзСтрокиВнутр ламає UUID посилань
3. V83.String() на десеріалізованому UUID не працює коректно
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'

START_DATE = datetime(2026, 1, 1)
END_DATE = datetime(2026, 1, 31)


def test_1_buhbud_query_params():
    """Тест 1: Чи працює запит ОстаткиИОбороты з 7 параметрами через COM?"""
    print("=" * 70)
    print("ТЕСТ 1: Запит ОстаткиИОбороты з 7 параметрами через COM")
    print("=" * 70)

    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_BUH)
    print("OK: Підключення до BuhBud")

    q = conn.NewObject("Запрос")

    # Запит з 7 параметрами (як в обробці, рядок 530-532)
    q.Текст = (
        'ВЫБРАТЬ '
        '   Остатки.Субконто1 КАК БанковскийСчетКасса, '
        '   ПРЕДСТАВЛЕНИЕ(Остатки.Субконто1) КАК БанковскийСчетПредставление, '
        '   Остатки.СуммаНачальныйОстаток КАК НачОстаток, '
        '   Остатки.СуммаОборотДт КАК Приход, '
        '   Остатки.СуммаОборотКт КАК Расход, '
        '   Остатки.СуммаКонечныйОстаток КАК КонОстаток, '
        '   Остатки.Счет.Код КАК СчетКод '
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

    try:
        result = q.Выполнить()
        tz = result.Выгрузить()
        count = tz.Количество()
        print(f"OK: Запит з 7 параметрами ПРАЦЮЄ! Рядків: {count}")

        sel = result.Выбрать()
        # Перевиконуємо бо вже Выгрузили
        result2 = q.Выполнить()
        sel2 = result2.Выбрать()
        i = 0
        while sel2.Следующий():
            i += 1
            print(f"  [{i}] Рахунок: {sel2.БанковскийСчетПредставление}  "
                  f"Код: {sel2.СчетКод}  "
                  f"НачОст={sel2.НачОстаток}  Прих={sel2.Приход}  "
                  f"Расх={sel2.Расход}  КонОст={sel2.КонОстаток}")
        return True
    except Exception as e:
        print(f"FAIL: Запит з 7 параметрами НЕ ПРАЦЮЄ: {e}")
        print("  -> Це причина порожніх BuhBud колонок в обробці!")
        return False
    finally:
        conn = None
        v8 = None


def test_2_serialization():
    """Тест 2: Чи зберігається UUID після ЗначениеВСтрокуВнутр/ЗначениеИзСтрокиВнутр?"""
    print()
    print("=" * 70)
    print("ТЕСТ 2: Серіалізація ЗначениеВСтрокуВнутр → ЗначениеИзСтрокиВнутр")
    print("=" * 70)

    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_BUH)
    print("OK: Підключення до BuhBud")

    q = conn.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ ПЕРВЫЕ 3 '
        '   Остатки.Субконто1 КАК БанковскийСчетКасса, '
        '   ПРЕДСТАВЛЕНИЕ(Остатки.Субконто1) КАК БанковскийСчетПредставление, '
        '   Остатки.СуммаКонечныйОстаток КАК КонОстаток '
        'ИЗ '
        '   РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты( '
        '       &НачалоПериода, &КонецПериода, , , '
        '       Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , ) КАК Остатки '
        'ГДЕ '
        '   Остатки.СуммаКонечныйОстаток <> 0'
    )
    q.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    tz = q.Выполнить().Выгрузить()
    print(f"OK: Отримано {tz.Количество()} рядків з BuhBud")

    # Спосіб 1: пряме зчитування UUID (як у звіті)
    print("\n--- Спосіб 1: Пряме зчитування UUID з COM ТЗ (як у звіті) ---")
    for i in range(tz.Количество()):
        row = tz.Получить(i)
        name = str(conn.String(row.БанковскийСчетПредставление))
        try:
            uuid_str = str(conn.String(row.БанковскийСчетКасса.УникальныйИдентификатор()))
            print(f"  [{i+1}] {name}: UUID = {uuid_str}")
        except Exception as e:
            print(f"  [{i+1}] {name}: ПОМИЛКА UUID = {e}")

    # Спосіб 2: серіалізація (як в обробці)
    print("\n--- Спосіб 2: Після ЗначениеВСтрокуВнутр/ЗначениеИзСтрокиВнутр ---")
    try:
        str_tz = conn.ЗначениеВСтрокуВнутр(tz)
        print(f"  OK: ЗначениеВСтрокуВнутр — {len(str_tz)} символів")

        # Десеріалізація в BuhBud контексті (для порівняння)
        tz_restored = conn.ЗначениеИзСтрокиВнутр(str_tz)
        print(f"  OK: ЗначениеИзСтрокиВнутр (в BuhBud контексті) — {tz_restored.Количество()} рядків")

        for i in range(tz_restored.Количество()):
            row = tz_restored.Получить(i)
            name = str(conn.String(row.БанковскийСчетПредставление))
            try:
                uuid_str = str(conn.String(row.БанковскийСчетКасса.УникальныйИдентификатор()))
                print(f"  [{i+1}] {name}: UUID = {uuid_str}")
            except Exception as e:
                print(f"  [{i+1}] {name}: ПОМИЛКА UUID = {e}")

    except Exception as e:
        print(f"  FAIL серіалізації: {e}")

    # Спосіб 3: серіалізація в BuhBud, десеріалізація в ERP (як в обробці!)
    print("\n--- Спосіб 3: Серіалізація BuhBud → Десеріалізація ERP (як в обробці) ---")
    conn_erp = v8.Connect(CONN_ERP)
    print("  OK: Підключення до ERP")

    try:
        # Серіалізація в BuhBud контексті
        str_tz_buh = conn.ЗначениеВСтрокуВнутр(tz)
        print(f"  OK: Серіалізовано в BuhBud — {len(str_tz_buh)} символів")

        # Десеріалізація в ERP контексті
        tz_erp = conn_erp.ЗначениеИзСтрокиВнутр(str_tz_buh)
        print(f"  OK: Десеріалізовано в ERP — {tz_erp.Количество()} рядків")

        for i in range(tz_erp.Количество()):
            row = tz_erp.Получить(i)
            name = str(conn_erp.String(row.БанковскийСчетПредставление))
            try:
                ref = row.БанковскийСчетКасса
                ref_type = str(conn_erp.String(conn_erp.ТипЗнч(ref)))
                print(f"  [{i+1}] {name}: ТипЗнч = {ref_type}")

                if str(conn_erp.String(conn_erp.ЗначениеЗаполнено(ref))) == "Да":
                    try:
                        uuid_obj = ref.УникальныйИдентификатор()
                        uuid_str = str(conn_erp.XMLСтрока(uuid_obj))
                        print(f"         UUID (XMLСтрока) = {uuid_str}")
                    except Exception as e2:
                        print(f"         UUID XMLСтрока ПОМИЛКА: {e2}")

                    try:
                        uuid_str2 = str(conn.String(ref.УникальныйИдентификатор()))
                        print(f"         UUID (V83.String) = {uuid_str2}")
                    except Exception as e3:
                        print(f"         UUID V83.String ПОМИЛКА: {e3}")
                else:
                    print(f"         Значення НЕ заповнене!")
            except Exception as e:
                print(f"  [{i+1}] {name}: ПОМИЛКА = {e}")

    except Exception as e:
        print(f"  FAIL: {e}")
        import traceback
        traceback.print_exc()

    conn = None
    conn_erp = None
    v8 = None


def test_3_uuid_comparison():
    """Тест 3: Порівняння UUID банківських рахунків ERP vs BuhBud."""
    print()
    print("=" * 70)
    print("ТЕСТ 3: Порівняння UUID банківських рахунків ERP vs BuhBud")
    print("=" * 70)

    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    # ERP: отримати банківські рахунки з UUID
    conn_erp = v8.Connect(CONN_ERP)
    print("OK: Підключення до ERP")

    q_erp = conn_erp.NewObject("Запрос")
    q_erp.Текст = (
        'ВЫБРАТЬ '
        '   БСО.Ссылка КАК Ссылка, '
        '   БСО.Наименование КАК Наименование, '
        '   БСО.НомерСчета КАК НомерСчета '
        'ИЗ '
        '   Справочник.БанковскиеСчетаОрганизаций КАК БСО '
        'ГДЕ '
        '   БСО.ПометкаУдаления = ЛОЖЬ'
    )
    erp_accounts = {}  # uuid -> name
    sel_erp = q_erp.Выполнить().Выбрать()
    while sel_erp.Следующий():
        uuid = str(conn_erp.XMLСтрока(sel_erp.Ссылка.УникальныйИдентификатор()))
        name = str(sel_erp.Наименование)
        num = str(sel_erp.НомерСчета)
        erp_accounts[uuid.lower()] = f"{name} ({num})"
    print(f"  ERP: {len(erp_accounts)} банківських рахунків")

    # BuhBud: отримати банківські рахунки з UUID
    conn_buh = v8.Connect(CONN_BUH)
    print("OK: Підключення до BuhBud")

    q_buh = conn_buh.NewObject("Запрос")
    q_buh.Текст = (
        'ВЫБРАТЬ '
        '   БСО.Ссылка КАК Ссылка, '
        '   БСО.Наименование КАК Наименование, '
        '   БСО.НомерСчета КАК НомерСчета '
        'ИЗ '
        '   Справочник.БанковскиеСчетаОрганизаций КАК БСО '
        'ГДЕ '
        '   БСО.ПометкаУдаления = ЛОЖЬ'
    )
    buh_accounts = {}  # uuid -> name
    sel_buh = q_buh.Выполнить().Выбрать()
    while sel_buh.Следующий():
        uuid = str(conn_buh.XMLСтрока(sel_buh.Ссылка.УникальныйИдентификатор()))
        name = str(sel_buh.Наименование)
        num = str(sel_buh.НомерСчета)
        buh_accounts[uuid.lower()] = f"{name} ({num})"
    print(f"  BuhBud: {len(buh_accounts)} банківських рахунків")

    # Порівняння UUID
    print("\n--- Збіги UUID ---")
    matched = 0
    for uuid, buh_name in sorted(buh_accounts.items(), key=lambda x: x[1]):
        erp_name = erp_accounts.get(uuid)
        if erp_name:
            matched += 1
            print(f"  MATCH: {buh_name} <-> {erp_name}  UUID={uuid}")
    print(f"  Всього збігів: {matched} з {len(buh_accounts)}")

    print("\n--- BuhBud рахунки БЕЗ збігу в ERP ---")
    unmatched_buh = 0
    for uuid, buh_name in sorted(buh_accounts.items(), key=lambda x: x[1]):
        if uuid not in erp_accounts:
            unmatched_buh += 1
            print(f"  NO MATCH: {buh_name}  UUID={uuid}")
    print(f"  Без збігу: {unmatched_buh}")

    print("\n--- ERP рахунки БЕЗ збігу в BuhBud ---")
    unmatched_erp = 0
    for uuid, erp_name in sorted(erp_accounts.items(), key=lambda x: x[1]):
        if uuid not in buh_accounts:
            unmatched_erp += 1
            print(f"  NO MATCH: {erp_name}  UUID={uuid}")
    print(f"  Без збігу: {unmatched_erp}")

    conn_erp = None
    conn_buh = None
    v8 = None


def test_4_processor_simulation():
    """Тест 4: Повна симуляція логіки обробки — де саме ламається?"""
    print()
    print("=" * 70)
    print("ТЕСТ 4: Симуляція логіки обробки (рядки 473-610)")
    print("=" * 70)

    pythoncom.CoInitialize()
    v8_connector = win32com.client.Dispatch("V83.COMConnector")

    # Підключення як в обробці: V83 = BuhBud COM
    V83 = v8_connector.Connect(CONN_BUH)
    print("OK: V83 (BuhBud)")

    # Запит до BuhBud (як в обробці, рядки 503-537)
    Запрос = V83.NewObject("Запрос")
    Запрос.Текст = (
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
    Запрос.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    Запрос.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    # Крок 1: Виконання запиту
    try:
        ТЗ = Запрос.Выполнить().Выгрузить()
        print(f"  Крок 1 (запит): OK — {ТЗ.Количество()} рядків")
    except Exception as e:
        print(f"  Крок 1 (запит): FAIL — {e}")
        print("  -> ЦЕ ПРИЧИНА! Запит не працює, повертається порожня таблиця")
        return

    # Крок 2: Серіалізація ЗначениеВСтрокуВнутр
    try:
        стр_ТЗ = V83.ЗначениеВСтрокуВнутр(ТЗ)
        print(f"  Крок 2 (серіалізація): OK — {len(стр_ТЗ)} символів")
    except Exception as e:
        print(f"  Крок 2 (серіалізація): FAIL — {e}")
        return

    # Крок 3: Десеріалізація (в контексті ERP? Ні — в обробці це теж виконується в ERP)
    # Але з Python ми не можемо виконати ЗначениеИзСтрокиВнутр в контексті ERP
    # Спробуємо в контексті BuhBud для перевірки
    try:
        ТаблицаRestored = V83.ЗначениеИзСтрокиВнутр(стр_ТЗ)
        print(f"  Крок 3 (десеріалізація BuhBud): OK — {ТаблицаRestored.Количество()} рядків")
    except Exception as e:
        print(f"  Крок 3 (десеріалізація BuhBud): FAIL — {e}")
        return

    # Крок 4: Зчитування UUID з десеріалізованої таблиці
    print(f"\n  Крок 4: Зчитування UUID з ТЗ (рядки 563-604 обробки)")
    for i in range(ТаблицаRestored.Количество()):
        СтрокаТЗ = ТаблицаRestored.Получить(i)
        name = str(V83.String(СтрокаТЗ.БанковскийСчетПредставление)).strip()
        код = str(V83.String(СтрокаТЗ.СчетКод)).strip()

        try:
            ref = СтрокаТЗ.БанковскийСчетКасса
            is_filled = V83.ЗначениеЗаполнено(ref)

            if is_filled:
                # Як в обробці рядок 596: V83.String(СтрокаТЗ.БанковскийСчетКасса.УникальныйИдентификатор())
                uuid_str = str(V83.String(ref.УникальныйИдентификатор()))
                print(f"  [{i+1}] {name} (сч.{код}): UUID = {uuid_str}  "
                      f"НачОст={СтрокаТЗ.НачОстаток}  КонОст={СтрокаТЗ.КонОстаток}")
            else:
                print(f"  [{i+1}] {name} (сч.{код}): НЕ ЗАПОВНЕНО!  "
                      f"НачОст={СтрокаТЗ.НачОстаток}  КонОст={СтрокаТЗ.КонОстаток}")
        except Exception as e:
            print(f"  [{i+1}] {name} (сч.{код}): ПОМИЛКА = {e}  "
                  f"НачОст={СтрокаТЗ.НачОстаток}  КонОст={СтрокаТЗ.КонОстаток}")

    # Крок 5: Десеріалізація в контексті ERP (як реально робить обробка)
    print(f"\n  Крок 5: Десеріалізація в контексті ERP (як в обробці)")
    conn_erp = v8_connector.Connect(CONN_ERP)
    print("  OK: Підключення до ERP")

    try:
        ТаблицаERP = conn_erp.ЗначениеИзСтрокиВнутр(стр_ТЗ)
        print(f"  OK: Десеріалізовано в ERP — {ТаблицаERP.Количество()} рядків")

        for i in range(ТаблицаERP.Количество()):
            СтрокаТЗ = ТаблицаERP.Получить(i)
            name = str(conn_erp.String(СтрокаТЗ.БанковскийСчетПредставление)).strip()

            try:
                ref = СтрокаТЗ.БанковскийСчетКасса
                type_name = str(conn_erp.String(conn_erp.ТипЗнч(ref)))

                if conn_erp.ЗначениеЗаполнено(ref):
                    try:
                        uuid_str = str(conn_erp.XMLСтрока(ref.УникальныйИдентификатор()))
                        print(f"  [{i+1}] {name}: Тип={type_name}  UUID={uuid_str}")

                        # Тепер перевіримо V83.String як в обробці
                        try:
                            uuid_v83 = str(V83.String(ref.УникальныйИдентификатор()))
                            print(f"         V83.String(UUID) = {uuid_v83}")
                        except Exception as e_v83:
                            print(f"         V83.String(UUID) FAIL: {e_v83}")

                    except Exception as e2:
                        print(f"  [{i+1}] {name}: Тип={type_name}  UUID FAIL: {e2}")
                else:
                    print(f"  [{i+1}] {name}: Тип={type_name}  НЕ ЗАПОВНЕНО!")
            except Exception as e:
                print(f"  [{i+1}] {name}: ПОМИЛКА = {e}")

    except Exception as e:
        print(f"  Десеріалізація в ERP FAIL: {e}")
        import traceback
        traceback.print_exc()

    conn_erp = None
    V83 = None
    v8_connector = None


if __name__ == "__main__":
    print(f"Період: {START_DATE.strftime('%d.%m.%Y')} - {END_DATE.strftime('%d.%m.%Y')}")
    print()

    try:
        test_1_buhbud_query_params()
    except Exception as e:
        print(f"FAIL тест 1: {e}")
        import traceback
        traceback.print_exc()

    try:
        test_2_serialization()
    except Exception as e:
        print(f"FAIL тест 2: {e}")
        import traceback
        traceback.print_exc()

    try:
        test_3_uuid_comparison()
    except Exception as e:
        print(f"FAIL тест 3: {e}")
        import traceback
        traceback.print_exc()

    try:
        test_4_processor_simulation()
    except Exception as e:
        print(f"FAIL тест 4: {e}")
        import traceback
        traceback.print_exc()

    print()
    print("=" * 70)
    print("ДІАГНОСТИКА ЗАВЕРШЕНА")
    print("=" * 70)
