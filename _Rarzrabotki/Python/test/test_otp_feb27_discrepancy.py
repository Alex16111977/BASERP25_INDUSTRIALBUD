"""
Діагностика розбіжності 293,706.85 на ОТП рахунку за лютий 2026.

ERP: расход = 203,967,570.70, кон.остаток = 293,706.85
BuhBud: расход = 204,261,277.55, кон.остаток = 0
Різниця: BuhBud має +293,706.85 расхода

Тест знаходить конкретний рух/документ що створює розбіжність.
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime
import re

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_BUH = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'

START_DATE = datetime(2026, 2, 1)
END_DATE = datetime(2026, 2, 28)


def normalize_date(raw):
    """Нормалізувати дату в формат yyyy-mm-dd."""
    s = str(raw)[:10]
    # dd.mm.yyyy -> yyyy-mm-dd
    m = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', s)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return s  # already yyyy-mm-dd


def test_otp_discrepancy():
    print("=" * 80)
    print("ДІАГНОСТИКА: розбіжність 293,706.85 на ОТП рахунку")
    print(f"Період: {START_DATE:%d.%m.%Y} - {END_DATE:%d.%m.%Y}")
    print("=" * 80)

    pythoncom.CoInitialize()
    v8_conn = win32com.client.Dispatch("V83.COMConnector")

    conn_erp = v8_conn.Connect(CONN_ERP)
    print("OK: ERP connected")

    conn_buh = v8_conn.Connect(CONN_BUH)
    print("OK: BuhBud connected")

    # ==========================================
    # КРОК 0: Знайти UUID ОТП рахунку
    # ==========================================
    print("\n--- КРОК 0: UUID ОТП рахунку ---")

    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ Б.Ссылка, Б.Наименование, Б.НомерСчета '
        'ИЗ Справочник.БанковскиеСчетаОрганизаций КАК Б '
        'ГДЕ Б.НомерСчета ПОДОБНО "%26004000010559%" И НЕ Б.ПометкаУдаления'
    )
    sel = q.Выполнить().Выбрать()
    otp_ref_erp = None
    otp_uuid = ""
    while sel.Следующий():
        ref = sel.Ссылка
        uuid = str(conn_erp.XMLСтрока(ref.УникальныйИдентификатор())).lower()
        # Перевіряємо чи має рухи
        q2 = conn_erp.NewObject("Запрос")
        q2.Текст = (
            'ВЫБРАТЬ СУММА(Д.Сумма) КАК С '
            'ИЗ РегистрНакопления.ДенежныеСредстваБезналичные КАК Д '
            'ГДЕ Д.БанковскийСчет = &Б И Д.Период МЕЖДУ &Н И &К'
        )
        q2.УстановитьПараметр("Б", ref)
        q2.УстановитьПараметр("Н", pywintypes.Time(START_DATE))
        q2.УстановитьПараметр("К", pywintypes.Time(datetime(2026, 2, 28, 23, 59, 59)))
        s2 = q2.Выполнить().Выбрать()
        total = 0
        if s2.Следующий():
            total = float(s2.С or 0)
        print(f"  {sel.Наименование}: UUID={uuid}, рухи={total:,.2f}")
        if total > 0:
            otp_ref_erp = ref
            otp_uuid = uuid

    print(f"  Використовую UUID: {otp_uuid}")

    # BuhBud ref
    otp_ref_buh = conn_buh.Справочники.БанковскиеСчета.ПолучитьСсылку(
        conn_buh.NewObject("УникальныйИдентификатор", otp_uuid)
    )

    # ==========================================
    # ТЕСТ 1: Денні підсумки ERP (віртуальна таблиця)
    # ==========================================
    print("\n--- ТЕСТ 1: ERP денні підсумки ---")

    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ '
        '   Т.Период КАК Період, '
        '   Т.СуммаПриход КАК Приход, '
        '   Т.СуммаРасход КАК Расход '
        'ИЗ '
        '   РегистрНакопления.ДенежныеСредстваБезналичные.ОстаткиИОбороты( '
        '       &Н, &К, День, , БанковскийСчет = &Б) КАК Т'
    )
    q.УстановитьПараметр("Н", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("К", pywintypes.Time(END_DATE))
    q.УстановитьПараметр("Б", otp_ref_erp)

    sel = q.Выполнить().Выбрать()
    erp_daily = {}
    while sel.Следующий():
        day = normalize_date(sel.Період)
        erp_daily[day] = {
            'prih': float(sel.Приход or 0),
            'rash': float(sel.Расход or 0),
        }

    # ==========================================
    # ТЕСТ 2: Денні підсумки BuhBud (віртуальна таблиця)
    # ==========================================
    print("--- ТЕСТ 2: BuhBud денні підсумки ---")

    q = conn_buh.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ '
        '   О.Период КАК Період, '
        '   О.Субконто1 КАК Субконто1, '
        '   ВЫБОР КОГДА О.Счет.Валютный '
        '       ТОГДА О.ВалютнаяСуммаОборотДт ИНАЧЕ О.СуммаОборотДт '
        '   КОНЕЦ КАК Приход, '
        '   ВЫБОР КОГДА О.Счет.Валютный '
        '       ТОГДА О.ВалютнаяСуммаОборотКт ИНАЧЕ О.СуммаОборотКт '
        '   КОНЕЦ КАК Расход '
        'ИЗ '
        '   РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты( '
        '       &Н, &К, День, , '
        '       Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)), , ) КАК О'
    )
    q.УстановитьПараметр("Н", pywintypes.Time(START_DATE))
    q.УстановитьПараметр("К", pywintypes.Time(END_DATE))

    ТЗ = q.Выполнить().Выгрузить()
    buh_daily = {}
    for i in range(ТЗ.Количество()):
        row = ТЗ.Получить(i)
        try:
            if not conn_buh.ЗначениеЗаполнено(row.Субконто1):
                continue
            sub_uuid = str(conn_buh.String(row.Субконто1.УникальныйИдентификатор())).lower()
        except:
            continue
        if sub_uuid != otp_uuid:
            continue
        day = normalize_date(row.Період)
        if day not in buh_daily:
            buh_daily[day] = {'prih': 0, 'rash': 0}
        buh_daily[day]['prih'] += float(row.Приход or 0)
        buh_daily[day]['rash'] += float(row.Расход or 0)

    # ==========================================
    # ТЕСТ 3: Порівняння по днях
    # ==========================================
    print("\n--- ТЕСТ 3: Порівняння по днях ---")
    print(f"  {'День':<12} {'ERP_Пр':>15} {'BUH_Пр':>15} {'D_Пр':>10}  {'ERP_Рас':>15} {'BUH_Рас':>15} {'D_Рас':>10}")
    print("  " + "-" * 100)

    all_days = sorted(set(erp_daily.keys()) | set(buh_daily.keys()))
    discrepancy_days = []
    for day in all_days:
        e = erp_daily.get(day, {'prih': 0, 'rash': 0})
        b = buh_daily.get(day, {'prih': 0, 'rash': 0})
        dp = round(e['prih'] - b['prih'], 2)
        dr = round(e['rash'] - b['rash'], 2)
        mark = " ***" if dp != 0 or dr != 0 else ""
        if dp != 0 or dr != 0:
            discrepancy_days.append(day)
        print(f"  {day:<12} {e['prih']:>15,.2f} {b['prih']:>15,.2f} {dp:>10,.2f}  {e['rash']:>15,.2f} {b['rash']:>15,.2f} {dr:>10,.2f}{mark}")

    erp_sp = sum(v['prih'] for v in erp_daily.values())
    erp_sr = sum(v['rash'] for v in erp_daily.values())
    buh_sp = sum(v['prih'] for v in buh_daily.values())
    buh_sr = sum(v['rash'] for v in buh_daily.values())
    print("  " + "-" * 100)
    print(f"  {'ВСЬОГО':<12} {erp_sp:>15,.2f} {buh_sp:>15,.2f} {erp_sp - buh_sp:>10,.2f}  {erp_sr:>15,.2f} {buh_sr:>15,.2f} {erp_sr - buh_sr:>10,.2f}")

    if not discrepancy_days:
        print("\n  Розбіжностей немає!")
        return

    print(f"\n  Дні з розбіжностями: {', '.join(discrepancy_days)}")

    # ==========================================
    # ТЕСТ 4: Рухи ERP на день розбіжності
    # ==========================================
    for disc_day in discrepancy_days:
        print(f"\n{'=' * 80}")
        print(f"ТЕСТ 4: Рухи на {disc_day}")
        print("=" * 80)

        e = erp_daily.get(disc_day, {'prih': 0, 'rash': 0})
        b = buh_daily.get(disc_day, {'prih': 0, 'rash': 0})
        print(f"  ERP:  П={e['prih']:,.2f}  Р={e['rash']:,.2f}")
        print(f"  BUH:  П={b['prih']:,.2f}  Р={b['rash']:,.2f}")
        print(f"  DIFF: П={e['prih'] - b['prih']:,.2f}  Р={e['rash'] - b['rash']:,.2f}")

        # 4a: ERP рухи на цей день
        print(f"\n  --- ERP рухи ({disc_day}) ---")
        q = conn_erp.NewObject("Запрос")
        q.Текст = (
            'ВЫБРАТЬ '
            '   Д.Регистратор КАК Регистратор, '
            '   ПРЕДСТАВЛЕНИЕ(Д.Регистратор) КАК РегПредст, '
            '   Д.ВидДвижения КАК Вид, '
            '   Д.Сумма КАК Сумма '
            'ИЗ РегистрНакопления.ДенежныеСредстваБезналичные КАК Д '
            'ГДЕ Д.БанковскийСчет = &Б '
            '   И Д.Период МЕЖДУ &Н И &К '
            'УПОРЯДОЧИТЬ ПО Д.Период'
        )
        q.УстановитьПараметр("Б", otp_ref_erp)
        # Parse disc_day (yyyy-mm-dd) to datetime
        parts = disc_day.split("-")
        day_start = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
        day_end = datetime(int(parts[0]), int(parts[1]), int(parts[2]), 23, 59, 59)
        q.УстановитьПараметр("Н", pywintypes.Time(day_start))
        q.УстановитьПараметр("К", pywintypes.Time(day_end))

        ТЗ_erp = q.Выполнить().Выгрузить()
        erp_regs = {}  # uuid -> {name, prih, rash}
        for i in range(ТЗ_erp.Количество()):
            row = ТЗ_erp.Получить(i)
            reg_uuid = str(conn_erp.XMLСтрока(row.Регистратор.УникальныйИдентификатор())).lower()
            reg_name = str(row.РегПредст)
            suma = float(row.Сумма or 0)

            # ВидДвижения: XMLСтрока повертає "Receipt" або "Expense"
            vid_xml = str(conn_erp.XMLСтрока(row.Вид))
            is_prih = vid_xml.lower() in ("receipt", "приход")

            if reg_uuid not in erp_regs:
                erp_regs[reg_uuid] = {'name': reg_name, 'prih': 0, 'rash': 0}
            if is_prih:
                erp_regs[reg_uuid]['prih'] += suma
            else:
                erp_regs[reg_uuid]['rash'] += suma

        erp_day_prih = sum(v['prih'] for v in erp_regs.values())
        erp_day_rash = sum(v['rash'] for v in erp_regs.values())
        print(f"  ERP: {len(erp_regs)} документів, П={erp_day_prih:,.2f}, Р={erp_day_rash:,.2f}")

        # 4b: BuhBud рухи на цей день (всі банк. рахунки)
        print(f"\n  --- BuhBud рухи ({disc_day}) ---")
        q = conn_buh.NewObject("Запрос")
        q.Текст = (
            'ВЫБРАТЬ '
            '   Д.Регистратор КАК Регистратор, '
            '   ПРЕДСТАВЛЕНИЕ(Д.Регистратор) КАК РегПредст, '
            '   Д.СчетДт.Код КАК СчетДтКод, '
            '   Д.СчетКт.Код КАК СчетКтКод, '
            '   Д.Сумма КАК Сумма '
            'ИЗ РегистрБухгалтерии.Хозрасчетный КАК Д '
            'ГДЕ Д.Период МЕЖДУ &Н И &К '
            '   И (Д.СчетДт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)) '
            '       ИЛИ Д.СчетКт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))) '
            'УПОРЯДОЧИТЬ ПО Д.Регистратор'
        )
        q.УстановитьПараметр("Н", pywintypes.Time(day_start))
        q.УстановитьПараметр("К", pywintypes.Time(day_end))

        ТЗ_buh = q.Выполнить().Выгрузить()
        buh_regs = {}  # uuid -> {name, suma_total, movements: [{sdt, skt, suma}]}
        for i in range(ТЗ_buh.Количество()):
            row = ТЗ_buh.Получить(i)
            try:
                reg_uuid = str(conn_buh.XMLСтрока(row.Регистратор.УникальныйИдентификатор())).lower()
            except:
                continue
            reg_name = str(conn_buh.String(row.РегПредст))
            sdt = str(conn_buh.String(row.СчетДтКод)).strip()
            skt = str(conn_buh.String(row.СчетКтКод)).strip()
            suma = float(row.Сумма or 0)

            if reg_uuid not in buh_regs:
                buh_regs[reg_uuid] = {'name': reg_name, 'moves': []}
            buh_regs[reg_uuid]['moves'].append({'sdt': sdt, 'skt': skt, 'suma': suma})

        print(f"  BuhBud: {len(buh_regs)} документів (всі банк.), {ТЗ_buh.Количество()} рухів")

        # 4c: Порівняння по UUID регістратора
        print(f"\n  --- Порівняння регістраторів ---")

        # BuhBud documents WITHOUT ERP counterpart
        buh_only_uuids = set(buh_regs.keys()) - set(erp_regs.keys())
        if buh_only_uuids:
            print(f"\n  BuhBud документи БЕЗ ERP відповідника: {len(buh_only_uuids)}")
            for uuid in sorted(buh_only_uuids):
                data = buh_regs[uuid]
                total = sum(m['suma'] for m in data['moves'])
                move_desc = "; ".join(f"Дт{m['sdt']}/Кт{m['skt']}={m['suma']:,.2f}" for m in data['moves'][:5])
                if len(data['moves']) > 5:
                    move_desc += f" ...+{len(data['moves'])-5} more"
                print(f"    {data['name'][:65]}")
                print(f"      UUID: {uuid}, Сума: {total:,.2f}, Рухів: {len(data['moves'])}")
                print(f"      {move_desc}")

        # ERP documents WITHOUT BuhBud counterpart
        erp_only_uuids = set(erp_regs.keys()) - set(buh_regs.keys())
        if erp_only_uuids:
            print(f"\n  ERP документи БЕЗ BuhBud відповідника: {len(erp_only_uuids)}")
            for uuid in sorted(erp_only_uuids):
                data = erp_regs[uuid]
                print(f"    {data['name'][:65]}")
                print(f"      UUID: {uuid}, П={data['prih']:,.2f}, Р={data['rash']:,.2f}")

        # Documents in BOTH but with different amounts
        common_uuids = set(erp_regs.keys()) & set(buh_regs.keys())
        diff_docs = []
        for uuid in common_uuids:
            e_data = erp_regs[uuid]
            b_data = buh_regs[uuid]
            b_total = sum(m['suma'] for m in b_data['moves'])
            e_total = e_data['prih'] + e_data['rash']
            if round(e_total - b_total, 2) != 0:
                diff_docs.append((uuid, e_data, b_data, e_total, b_total))

        if diff_docs:
            print(f"\n  Документи з РІЗНИЦЕЮ сум: {len(diff_docs)}")
            for uuid, e_data, b_data, e_total, b_total in diff_docs:
                print(f"    {e_data['name'][:65]}")
                print(f"      ERP: П={e_data['prih']:,.2f} Р={e_data['rash']:,.2f} (всього={e_total:,.2f})")
                print(f"      BUH: всього={b_total:,.2f}")
                print(f"      DIFF: {e_total - b_total:,.2f}")

        # Шукаємо документ з сумою 293,706.85 (або близькою)
        print(f"\n  --- Пошук документів з сумою ~293,706.85 ---")
        target = 293706.85

        print(f"\n  В ERP:")
        for uuid, data in erp_regs.items():
            if abs(data['prih'] - target) < 0.01 or abs(data['rash'] - target) < 0.01:
                mark = " *** ЗНАЙДЕНО ***"
                in_buh = uuid in buh_regs
                print(f"    {data['name'][:65]}{mark}")
                print(f"      UUID: {uuid}, П={data['prih']:,.2f}, Р={data['rash']:,.2f}")
                print(f"      Є в BuhBud: {in_buh}")
                if in_buh:
                    b = buh_regs[uuid]
                    b_total = sum(m['suma'] for m in b['moves'])
                    print(f"      BuhBud сума: {b_total:,.2f}")
                    for m in b['moves']:
                        print(f"        Дт{m['sdt']}/Кт{m['skt']} = {m['suma']:,.2f}")

        print(f"\n  В BuhBud:")
        for uuid, data in buh_regs.items():
            total = sum(m['suma'] for m in data['moves'])
            for m in data['moves']:
                if abs(m['suma'] - target) < 0.01:
                    in_erp = uuid in erp_regs
                    print(f"    {data['name'][:65]} *** ЗНАЙДЕНО ***")
                    print(f"      UUID: {uuid}, Дт{m['sdt']}/Кт{m['skt']} = {m['suma']:,.2f}")
                    print(f"      Є в ERP: {in_erp}")
                    if in_erp:
                        e = erp_regs[uuid]
                        print(f"      ERP: П={e['prih']:,.2f}, Р={e['rash']:,.2f}")

    # ==========================================
    # ПІДСУМОК
    # ==========================================
    print("\n" + "=" * 80)
    print("ПІДСУМОК")
    print("=" * 80)
    print(f"  ERP:    Приход={erp_sp:,.2f}  Расход={erp_sr:,.2f}  Залишок={erp_sp - erp_sr:,.2f}")
    print(f"  BuhBud: Приход={buh_sp:,.2f}  Расход={buh_sr:,.2f}  Залишок={buh_sp - buh_sr:,.2f}")
    print(f"  РІЗНИЦЯ: Приход={erp_sp - buh_sp:,.2f}  Расход={erp_sr - buh_sr:,.2f}")

    conn_erp = None
    conn_buh = None
    v8_conn = None


if __name__ == "__main__":
    try:
        test_otp_discrepancy()
    except Exception as e:
        print(f"FAIL: {e}")
        import traceback
        traceback.print_exc()
