# -*- coding: utf-8 -*-
"""
Тестовий скрипт: перевірка запитів для звіту порівняння залишків ОС
між ERP (РегістрБухгалтерії.Хозрасчетный) та BuhBud (COM).

Тести:
  1. BuhBud — залишки ОС по рахунках 10x (з А_ИдКод)
  2. BuhBud — амортизація по рахунку 131
  3. BuhBud — об'єднання ОС + Амортизація по А_ИдКод
  4. ERP — залишки ОС по рахунках 10x (з А_ИдКод)
  5. ERP — амортизація по рахунку 131
  6. Маппінг А_ИдКод — перетин BuhBud та ERP
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONNECTION_STRING_BUHBUD = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
CONNECTION_STRING_ERP = 'Srvr="SQLSERVER";Ref="BASERP25";Usr="cfo";Pwd="2442"'
DATE_OSTATKI = datetime(2026, 2, 28, 23, 59, 59)


def make_time(dt):
    return pywintypes.Time(dt)


def connect(conn_string, label):
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(conn_string)
    print(f"[OK] Підключення до {label} встановлено")
    return conn


# ============================================================================
# ТЕСТИ BuhBud (COM)
# ============================================================================

def test1_buhbud_ostatki_os(conn):
    """Тест 1: BuhBud — залишки ОС по рахунках 10x (з А_ИдКод)"""
    print("\n" + "=" * 80)
    print("ТЕСТ 1: BuhBud — Залишки ОС (рах. 10x) з А_ИдКод")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Наименование КАК ОС_Наименование,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Код КАК ОС_Код,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).А_ИдКод КАК ОС_А_ИдКод,\n"
        "  Ост.Счет.Код КАК КодСчета,\n"
        "  СУММА(Ост.СуммаОстатокДт) КАК ПервоначальнаяСтоимость,\n"
        "  СУММА(Ост.КоличествоОстатокДт) КАК КолВо\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ОсновныеСредства)),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "ГДЕ\n"
        "  Ост.Субконто1 ССЫЛКА Справочник.ОсновныеСредства\n"
        "  И Ост.СуммаОстатокДт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Наименование,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Код,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).А_ИдКод,\n"
        "  Ост.Счет.Код"
    )
    query.УстановитьПараметр("Дата", make_time(DATE_OSTATKI))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        total_perv = 0
        total_kol = 0
        data = {}  # А_ИдКод -> {КодСчета, ПервоначальнаяСтоимость, Количество, Наименование}

        print(f"\n  {'№':>4s} | {'Код':10s} | {'Наименование':<40s} | {'Рах':>4s} | {'Кіл':>5s} | {'Перв.вартість':>15s} | {'А_ИдКод'}")
        print(f"  {'-'*4} | {'-'*10} | {'-'*40} | {'-'*4} | {'-'*5} | {'-'*15} | {'-'*36}")

        while result.Следующий():
            count += 1
            naim = str(result.ОС_Наименование or "").strip()
            kod = str(result.ОС_Код or "").strip()
            id_kod = str(result.ОС_А_ИдКод or "").strip()
            schet = str(result.КодСчета or "").strip()
            perv = float(result.ПервоначальнаяСтоимость or 0)
            kol = float(result.КолВо or 0)
            total_perv += perv
            total_kol += kol

            if count <= 10:
                print(f"  {count:4d} | {kod:10s} | {naim[:40]:<40s} | {schet:>4s} | {kol:5.0f} | {perv:>15.2f} | {id_kod}")

            if id_kod:
                data[id_kod] = {
                    "КодСчета": schet,
                    "ПервоначальнаяСтоимость": perv,
                    "Количество": kol,
                    "Наименование": naim
                }

        if count > 10:
            print(f"  ... (ще {count - 10} записів)")

        print(f"\n  Всього записів: {count}")
        print(f"  Всього кількість: {total_kol:.0f}")
        print(f"  Всього перв. вартість: {total_perv:>15.2f}")
        print(f"  Записів з А_ИдКод: {len(data)}")
        empty_id = count - len(data)
        if empty_id > 0:
            print(f"  [!] Записів БЕЗ А_ИдКод: {empty_id}")
        print(f"[OK] Тест 1 пройдено. Записів: {count}")
        return data
    except Exception as e:
        print(f"[ПОМИЛКА] Тест 1: {e}")
        import traceback
        traceback.print_exc()
        return {}


def test2_buhbud_amortizacia(conn):
    """Тест 2: BuhBud — амортизація по рахунку 131 (з А_ИдКод)"""
    print("\n" + "=" * 80)
    print("ТЕСТ 2: BuhBud — Амортизація (рах. 131) з А_ИдКод")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Наименование КАК ОС_Наименование,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Код КАК ОС_Код,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).А_ИдКод КАК ОС_А_ИдКод,\n"
        "  СУММА(Ост.СуммаОстатокКт) КАК НакопленнаяАмортизация\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ИзносОсновныхСредств)),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "ГДЕ\n"
        "  Ост.Субконто1 ССЫЛКА Справочник.ОсновныеСредства\n"
        "  И Ост.СуммаОстатокКт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Наименование,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Код,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).А_ИдКод"
    )
    query.УстановитьПараметр("Дата", make_time(DATE_OSTATKI))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        total_amort = 0
        data = {}  # А_ИдКод -> НакопленнаяАмортизация

        print(f"\n  {'№':>4s} | {'Код':10s} | {'Наименование':<40s} | {'Амортизація':>15s} | {'А_ИдКод'}")
        print(f"  {'-'*4} | {'-'*10} | {'-'*40} | {'-'*15} | {'-'*36}")

        while result.Следующий():
            count += 1
            naim = str(result.ОС_Наименование or "").strip()
            kod = str(result.ОС_Код or "").strip()
            id_kod = str(result.ОС_А_ИдКод or "").strip()
            amort = float(result.НакопленнаяАмортизация or 0)
            total_amort += amort

            if count <= 10:
                print(f"  {count:4d} | {kod:10s} | {naim[:40]:<40s} | {amort:>15.2f} | {id_kod}")

            if id_kod:
                data[id_kod] = amort

        if count > 10:
            print(f"  ... (ще {count - 10} записів)")

        print(f"\n  Всього записів: {count}")
        print(f"  Всього амортизація: {total_amort:>15.2f}")
        print(f"  Записів з А_ИдКод: {len(data)}")
        print(f"[OK] Тест 2 пройдено. Записів: {count}")
        return data
    except Exception as e:
        print(f"[ПОМИЛКА] Тест 2: {e}")
        import traceback
        traceback.print_exc()
        return {}


def test3_buhbud_combined(data_os, data_amort):
    """Тест 3: BuhBud — об'єднання ОС + Амортизація по А_ИдКод"""
    print("\n" + "=" * 80)
    print("ТЕСТ 3: BuhBud — Об'єднання ОС + Амортизація (перевірка ОстаточнаяСтоимость)")
    print("=" * 80)

    if not data_os:
        print("[!] Немає даних ОС з тесту 1, пропускаємо")
        return {}

    combined = {}
    total_perv = 0
    total_amort = 0
    total_ost = 0

    for id_kod, os_info in data_os.items():
        perv = os_info["ПервоначальнаяСтоимость"]
        amort = data_amort.get(id_kod, 0)
        ost = perv - amort
        combined[id_kod] = {
            "Наименование": os_info["Наименование"],
            "КодСчета": os_info["КодСчета"],
            "Количество": os_info["Количество"],
            "ПервоначальнаяСтоимость": perv,
            "НакопленнаяАмортизация": amort,
            "ОстаточнаяСтоимость": ost
        }
        total_perv += perv
        total_amort += amort
        total_ost += ost

    # Показати перші 10
    print(f"\n  {'№':>4s} | {'Наименование':<40s} | {'Рах':>4s} | {'Перв.вартість':>15s} | {'Амортизація':>15s} | {'Залишкова':>15s}")
    print(f"  {'-'*4} | {'-'*40} | {'-'*4} | {'-'*15} | {'-'*15} | {'-'*15}")

    for i, (id_kod, info) in enumerate(combined.items()):
        if i >= 10:
            break
        print(f"  {i+1:4d} | {info['Наименование'][:40]:<40s} | {info['КодСчета']:>4s} | {info['ПервоначальнаяСтоимость']:>15.2f} | {info['НакопленнаяАмортизация']:>15.2f} | {info['ОстаточнаяСтоимость']:>15.2f}")

    if len(combined) > 10:
        print(f"  ... (ще {len(combined) - 10} записів)")

    # ОС без амортизації (нові ОС або ті, де амортизація = 0)
    os_bez_amort = sum(1 for id_kod in data_os if id_kod not in data_amort)
    amort_bez_os = sum(1 for id_kod in data_amort if id_kod not in data_os)

    print(f"\n  Всього ОС (з А_ИдКод): {len(combined)}")
    print(f"  Перв. вартість:  {total_perv:>15.2f}")
    print(f"  Амортизація:     {total_amort:>15.2f}")
    print(f"  Залишкова:       {total_ost:>15.2f}")
    print(f"  ОС без амортизації: {os_bez_amort}")
    print(f"  Амортизація без ОС: {amort_bez_os}")

    # Перевірка
    if total_perv > 0:
        print(f"  Коефіцієнт зносу: {total_amort / total_perv * 100:.1f}%")

    print(f"[OK] Тест 3 пройдено. Об'єднано: {len(combined)} записів")
    return combined


# ============================================================================
# ТЕСТИ ERP (COM)
# ============================================================================

def test4_erp_ostatki_os(conn):
    """Тест 4: ERP — залишки ОС по рахунках 10x (з А_ИдКод)"""
    print("\n" + "=" * 80)
    print("ТЕСТ 4: ERP — Залишки ОС (рах. 10x) з А_ИдКод")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ\n"
        "  ОЭ.Наименование КАК ОС_Наименование,\n"
        "  ОЭ.Код КАК ОС_Код,\n"
        "  ОЭ.А_ИдКод КАК ОС_А_ИдКод,\n"
        "  Ост.Счет.Код КАК КодСчета,\n"
        "  СУММА(Ост.СуммаОстатокДт) КАК ПервоначальнаяСтоимость,\n"
        "  СУММА(Ост.КоличествоОстатокДт) КАК КолВо\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ОсновныеСредства)),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "  ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ОбъектыЭксплуатации КАК ОЭ\n"
        "  ПО Ост.Субконто1 = ОЭ.Ссылка\n"
        "ГДЕ\n"
        "  Ост.Субконто1 ССЫЛКА Справочник.ОбъектыЭксплуатации\n"
        "  И Ост.СуммаОстатокДт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  ОЭ.Наименование,\n"
        "  ОЭ.Код,\n"
        "  ОЭ.А_ИдКод,\n"
        "  Ост.Счет.Код"
    )
    query.УстановитьПараметр("Дата", make_time(DATE_OSTATKI))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        total_perv = 0
        total_kol = 0
        data = {}

        print(f"\n  {'№':>4s} | {'Код':10s} | {'Наименование':<40s} | {'Рах':>4s} | {'Кіл':>5s} | {'Перв.вартість':>15s} | {'А_ИдКод'}")
        print(f"  {'-'*4} | {'-'*10} | {'-'*40} | {'-'*4} | {'-'*5} | {'-'*15} | {'-'*36}")

        while result.Следующий():
            count += 1
            naim = str(result.ОС_Наименование or "").strip()
            kod = str(result.ОС_Код or "").strip()
            id_kod = str(result.ОС_А_ИдКод or "").strip()
            schet = str(result.КодСчета or "").strip()
            perv = float(result.ПервоначальнаяСтоимость or 0)
            kol = float(result.КолВо or 0)
            total_perv += perv
            total_kol += kol

            if count <= 10:
                print(f"  {count:4d} | {kod:10s} | {naim[:40]:<40s} | {schet:>4s} | {kol:5.0f} | {perv:>15.2f} | {id_kod}")

            if id_kod:
                data[id_kod] = {
                    "КодСчета": schet,
                    "ПервоначальнаяСтоимость": perv,
                    "Количество": kol,
                    "Наименование": naim
                }

        if count > 10:
            print(f"  ... (ще {count - 10} записів)")

        print(f"\n  Всього записів: {count}")
        print(f"  Всього кількість: {total_kol:.0f}")
        print(f"  Всього перв. вартість: {total_perv:>15.2f}")
        print(f"  Записів з А_ИдКод: {len(data)}")
        empty_id = count - len(data)
        if empty_id > 0:
            print(f"  [!] Записів БЕЗ А_ИдКод: {empty_id}")
        print(f"[OK] Тест 4 пройдено. Записів: {count}")
        return data
    except Exception as e:
        print(f"[ПОМИЛКА] Тест 4: {e}")
        import traceback
        traceback.print_exc()
        return {}


def test5_erp_amortizacia(conn):
    """Тест 5: ERP — амортизація по рахунку 131 (з А_ИдКод)"""
    print("\n" + "=" * 80)
    print("ТЕСТ 5: ERP — Амортизація (рах. 131) з А_ИдКод")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ\n"
        "  ОЭ.Наименование КАК ОС_Наименование,\n"
        "  ОЭ.Код КАК ОС_Код,\n"
        "  ОЭ.А_ИдКод КАК ОС_А_ИдКод,\n"
        "  СУММА(Ост.СуммаОстатокКт) КАК НакопленнаяАмортизация\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ИзносОсновныхСредств)),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "  ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ОбъектыЭксплуатации КАК ОЭ\n"
        "  ПО Ост.Субконто1 = ОЭ.Ссылка\n"
        "ГДЕ\n"
        "  Ост.Субконто1 ССЫЛКА Справочник.ОбъектыЭксплуатации\n"
        "  И Ост.СуммаОстатокКт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  ОЭ.Наименование,\n"
        "  ОЭ.Код,\n"
        "  ОЭ.А_ИдКод"
    )
    query.УстановитьПараметр("Дата", make_time(DATE_OSTATKI))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        total_amort = 0
        data = {}

        print(f"\n  {'№':>4s} | {'Код':10s} | {'Наименование':<40s} | {'Амортизація':>15s} | {'А_ИдКод'}")
        print(f"  {'-'*4} | {'-'*10} | {'-'*40} | {'-'*15} | {'-'*36}")

        while result.Следующий():
            count += 1
            naim = str(result.ОС_Наименование or "").strip()
            kod = str(result.ОС_Код or "").strip()
            id_kod = str(result.ОС_А_ИдКод or "").strip()
            amort = float(result.НакопленнаяАмортизация or 0)
            total_amort += amort

            if count <= 10:
                print(f"  {count:4d} | {kod:10s} | {naim[:40]:<40s} | {amort:>15.2f} | {id_kod}")

            if id_kod:
                data[id_kod] = amort

        if count > 10:
            print(f"  ... (ще {count - 10} записів)")

        print(f"\n  Всього записів: {count}")
        print(f"  Всього амортизація: {total_amort:>15.2f}")
        print(f"  Записів з А_ИдКод: {len(data)}")
        print(f"[OK] Тест 5 пройдено. Записів: {count}")
        return data
    except Exception as e:
        print(f"[ПОМИЛКА] Тест 5: {e}")
        import traceback
        traceback.print_exc()
        return {}


def test6_mapping(data_buhbud, data_erp):
    """Тест 6: Маппінг А_ИдКод — перетин BuhBud та ERP"""
    print("\n" + "=" * 80)
    print("ТЕСТ 6: Маппінг А_ИдКод — Перетин BuhBud та ERP")
    print("=" * 80)

    if not data_buhbud:
        print("[!] Немає даних BuhBud, пропускаємо")
        return
    if not data_erp:
        print("[!] Немає даних ERP, пропускаємо")
        return

    keys_buhbud = set(data_buhbud.keys())
    keys_erp = set(data_erp.keys())

    matched = keys_buhbud & keys_erp
    only_buhbud = keys_buhbud - keys_erp
    only_erp = keys_erp - keys_buhbud

    print(f"\n  BuhBud А_ИдКод: {len(keys_buhbud)}")
    print(f"  ERP А_ИдКод:    {len(keys_erp)}")
    print(f"  Збіглися:       {len(matched)}")
    print(f"  Тільки BuhBud:  {len(only_buhbud)}")
    print(f"  Тільки ERP:     {len(only_erp)}")

    if len(keys_buhbud) > 0:
        pct = len(matched) / len(keys_buhbud) * 100
        print(f"\n  Відсоток маппінгу (від BuhBud): {pct:.1f}%")
    if len(keys_erp) > 0:
        pct = len(matched) / len(keys_erp) * 100
        print(f"  Відсоток маппінгу (від ERP):    {pct:.1f}%")

    # Показати перші незбіжні
    if only_buhbud:
        print(f"\n  Перші 10 з BuhBud без пари в ERP:")
        for i, id_kod in enumerate(sorted(only_buhbud)):
            if i >= 10:
                break
            info = data_buhbud[id_kod]
            print(f"    {id_kod} — {info['Наименование'][:50]}")

    if only_erp:
        print(f"\n  Перші 10 з ERP без пари в BuhBud:")
        for i, id_kod in enumerate(sorted(only_erp)):
            if i >= 10:
                break
            info = data_erp[id_kod]
            print(f"    {id_kod} — {info['Наименование'][:50]}")

    # Показати розбіжності по сумах для збігшихся
    print(f"\n  Розбіжності по сумах (перші 10):")
    print(f"  {'А_ИдКод':<38s} | {'Наименование':<30s} | {'Перв.BuhBud':>12s} | {'Перв.ERP':>12s} | {'Різниця':>12s}")
    print(f"  {'-'*38} | {'-'*30} | {'-'*12} | {'-'*12} | {'-'*12}")
    diff_count = 0
    for id_kod in sorted(matched):
        perv_buh = data_buhbud[id_kod]["ПервоначальнаяСтоимость"]
        perv_erp = data_erp[id_kod]["ПервоначальнаяСтоимость"]
        diff = perv_buh - perv_erp
        if abs(diff) > 0.01:
            diff_count += 1
            if diff_count <= 10:
                naim = data_buhbud[id_kod]["Наименование"][:30]
                print(f"  {id_kod:<38s} | {naim:<30s} | {perv_buh:>12.2f} | {perv_erp:>12.2f} | {diff:>12.2f}")

    print(f"\n  Всього з розбіжностями по перв. вартості: {diff_count}")
    print(f"[OK] Тест 6 пройдено.")


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 80)
    print("  Тестування запитів для звіту порівняння залишків ОС")
    print(f"  Дата залишків: {DATE_OSTATKI:%d.%m.%Y}")
    print("=" * 80)

    conn_buhbud = None
    conn_erp = None

    try:
        # --- BuhBud тести ---
        conn_buhbud = connect(CONNECTION_STRING_BUHBUD, "BuhBud")
        data_os_buh = test1_buhbud_ostatki_os(conn_buhbud)
        data_amort_buh = test2_buhbud_amortizacia(conn_buhbud)
        data_combined_buh = test3_buhbud_combined(data_os_buh, data_amort_buh)

        # --- ERP тести ---
        try:
            conn_erp = connect(CONNECTION_STRING_ERP, "ERP")
            data_os_erp = test4_erp_ostatki_os(conn_erp)
            data_amort_erp = test5_erp_amortizacia(conn_erp)

            # --- Маппінг ---
            test6_mapping(data_combined_buh, data_os_erp)
        except Exception as e:
            print(f"\n[!] Не вдалося підключитися до ERP: {e}")
            print("[!] Тести 4-6 (ERP) пропущено. Перевірте CONNECTION_STRING_ERP")
            import traceback
            traceback.print_exc()

    except Exception as e:
        print(f"\n[КРИТИЧНА ПОМИЛКА] {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn_buhbud = None
        conn_erp = None
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        print("\n" + "=" * 80)
        print("[OK] Завершено")
        print("=" * 80)


if __name__ == "__main__":
    main()
