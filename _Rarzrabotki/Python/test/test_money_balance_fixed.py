"""
Верифікаційний тест: симуляція виправленої логіки ПолучитьОстаткиBuhBud_Деньги.

Перевіряє що UUID зчитуються коректно з BuhBud COM таблиці (без серіалізації)
та збігаються з UUID банківських рахунків ERP.
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_BUH = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'

START_DATE = datetime(2026, 1, 1)
END_DATE = datetime(2026, 1, 31)


def test_fixed_logic():
    """Симуляція виправленої логіки: пряма ітерація по COM ТЗ."""
    print("=" * 70)
    print("Верифікація: виправлена логіка (пряма ітерація COM)")
    print(f"Період: {START_DATE.strftime('%d.%m.%Y')} - {END_DATE.strftime('%d.%m.%Y')}")
    print("=" * 70)

    pythoncom.CoInitialize()
    v8_conn = win32com.client.Dispatch("V83.COMConnector")

    # --- ERP: отримати залишки ---
    conn_erp = v8_conn.Connect(CONN_ERP)
    print("OK: ERP")

    q_erp = conn_erp.NewObject("Запрос")
    q_erp.Текст = (
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
    q_erp.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q_erp.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    erp_data = {}  # uuid -> {name, nач, прих, расх, кон}
    sel = q_erp.Выполнить().Выбрать()
    while sel.Следующий():
        uuid = str(conn_erp.XMLСтрока(sel.БанковскийСчет.УникальныйИдентификатор())).lower()
        erp_data[uuid] = {
            'name': str(sel.Предст),
            'nach': float(sel.НачОстаток or 0),
            'prih': float(sel.Приход or 0),
            'rash': float(sel.Расход or 0),
            'kon': float(sel.КонОстаток or 0),
        }
    print(f"  ERP: {len(erp_data)} рахунків з залишками")

    # --- BuhBud: отримати залишки (пряма ітерація COM) ---
    V83 = v8_conn.Connect(CONN_BUH)
    print("OK: BuhBud")

    q_buh = V83.NewObject("Запрос")
    q_buh.Текст = (
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
    q_buh.УстановитьПараметр("НачалоПериода", pywintypes.Time(START_DATE))
    q_buh.УстановитьПараметр("КонецПериода", pywintypes.Time(END_DATE))

    ТЗ = q_buh.Выполнить().Выгрузить()
    print(f"  BuhBud: {ТЗ.Количество()} рядків")

    # Симуляція виправленої логіки
    buh_data = {}  # uuid -> {name, nач, прих, расх, кон}
    uuid_ok = 0
    uuid_fail = 0
    uuid_empty = 0

    print("\n--- BuhBud рахунки (пряма ітерація COM) ---")
    for i in range(ТЗ.Количество()):
        row = ТЗ.Получить(i)
        name = str(V83.String(row.БанковскийСчетПредставление)).strip()
        code = str(V83.String(row.СчетКод)).strip()
        nach = float(row.НачОстаток or 0)
        kon = float(row.КонОстаток or 0)

        uuid = ""
        try:
            if code == "315":
                # 315 обробляється по ЄДРПОУ (залишаємо порожнім для тесту)
                uuid = "315-special"
            elif V83.ЗначениеЗаполнено(row.БанковскийСчетКасса):
                uuid = str(V83.String(row.БанковскийСчетКасса.УникальныйИдентификатор())).lower()
                uuid_ok += 1
            else:
                uuid_empty += 1
        except Exception as e:
            uuid_fail += 1
            print(f"  [{i+1}] {name} (сч.{code}): UUID FAIL = {e}")

        if uuid and uuid != "315-special":
            if uuid not in buh_data:
                buh_data[uuid] = {'name': name, 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0}
            buh_data[uuid]['nach'] += nach
            buh_data[uuid]['prih'] += float(row.Приход or 0)
            buh_data[uuid]['rash'] += float(row.Расход or 0)
            buh_data[uuid]['kon'] += kon

        match_erp = "MATCH" if uuid and uuid in erp_data else ("315" if uuid == "315-special" else "NO-MATCH")
        print(f"  [{i+1}] {name} (сч.{code}): UUID={uuid[:20]}...  КонОст={kon}  -> {match_erp}")

    print(f"\n  UUID OK: {uuid_ok}, Empty: {uuid_empty}, Fail: {uuid_fail}")

    # --- Зведення: порівняння залишків ---
    print("\n" + "=" * 70)
    print("ЗВЕДЕННЯ: порівняння залишків (як робить СвестиТаблицы_Деньги)")
    print("=" * 70)

    all_keys = set(erp_data.keys()) | set(buh_data.keys())
    discrepancies = 0
    matches = 0

    for uuid in sorted(all_keys):
        erp = erp_data.get(uuid, {'name': '???', 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0})
        buh = buh_data.get(uuid, {'name': '???', 'nach': 0, 'prih': 0, 'rash': 0, 'kon': 0})

        diff_nach = round(erp['nach'] - buh['nach'], 2)
        diff_prih = round(erp['prih'] - buh['prih'], 2)
        diff_rash = round(erp['rash'] - buh['rash'], 2)
        diff_kon = round(erp['kon'] - buh['kon'], 2)

        if diff_nach == 0 and diff_prih == 0 and diff_rash == 0 and diff_kon == 0:
            matches += 1
        else:
            discrepancies += 1
            name = erp['name'] if uuid in erp_data else buh['name']
            print(f"  ROZB: {name}")
            print(f"    ERP:  Nач={erp['nach']}  Пр={erp['prih']}  Рас={erp['rash']}  Кон={erp['kon']}")
            print(f"    BUH:  Nач={buh['nach']}  Пр={buh['prih']}  Рас={buh['rash']}  Кон={buh['kon']}")
            print(f"    DIFF: Nач={diff_nach}  Пр={diff_prih}  Рас={diff_rash}  Кон={diff_kon}")

    print(f"\nРезультат: {matches} збігів, {discrepancies} розбіжностей")
    if discrepancies == 0:
        print("OK: 0 розбіжностей (як і в звіті)!")
    else:
        print(f"УВАГА: {discrepancies} розбіжностей (потрібна додаткова діагностика)")

    V83 = None
    conn_erp = None
    v8_conn = None


if __name__ == "__main__":
    try:
        test_fixed_logic()
    except Exception as e:
        print(f"FAIL: {e}")
        import traceback
        traceback.print_exc()
