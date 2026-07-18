# -*- coding: utf-8 -*-
"""
GREEN-фаза: загрузить EPF через COM и вызвать ПолучитьРасшифровкуПлатежей().

Проверяет:
  1. EPF загружается без ошибок
  2. Экспорт-функция ПолучитьРасшифровкуПлатежей возвращает структуру
     с двумя ТЗ: Расшифровка и РасшифровкаДокументы
  3. РасшифровкаДокументы = 204 строки (детализация)
  4. Расшифровка = свёртка (не больше 204, обычно ~10–30)
  5. Сумма по Расшифровка ≈ 10 033 642,31
  6. Среди контрагентов есть "Європабуд ТОВ"
"""

import sys
import datetime
from pathlib import Path
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EPF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\ЗаполнениеСтатьиДенегВДоговорах.epf"
STATIA_UUID = "3e2ba4c7-7a7e-11eb-a208-000c299fb278"

EXPECTED_DOCS              = 237  # все платежи договоров (включая другие статьи)
EXPECTED_DOCS_OUR_STATIA   = 204  # только наша статья
EXPECTED_TOTAL             = 10_033_642.31  # сумма по нашей статье (свёртка Расшифровка)
EXPECTED_MULTI_CONTRACTS   = 9   # договоры, у которых были другие статьи ДДС за период
EXPECTED_MULTI_PARTNERS    = ["БЕТОН-СЕРВІС", "Кійко", "Терещенко"]  # должны быть среди многоцелевых
EXPECTED_DISTINCT_STATEI   = 6   # разных статей в детализации
STATIA_NAME                = "Аренда сторонней техники"


def fail(msg):
    print(f"FAIL: {msg}")
    sys.exit(1)


def main():
    if not Path(EPF_PATH).exists():
        fail(f"EPF не собран: {EPF_PATH}")

    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)
    print("OK: connected to BaseERP")

    vn_obr = erp.ВнешниеОбработки.Создать(EPF_PATH, True)
    print(f"OK: EPF loaded ({Path(EPF_PATH).stat().st_size} bytes)")

    uid = erp.NewObject("УникальныйИдентификатор", STATIA_UUID)
    statia_ref = erp.Справочники.СтатьиДвиженияДенежныхСредств.ПолучитьСсылку(uid)
    nach = datetime.datetime(2025, 12, 1)
    okon = datetime.datetime(2026, 3, 31, 23, 59, 59)

    try:
        result = vn_obr.ПолучитьРасшифровкуПлатежей(nach, okon, statia_ref)
    except Exception as e:
        if hasattr(e, "excepinfo") and e.excepinfo:
            fail(f"ПолучитьРасшифровкуПлатежей упала: {e.excepinfo[2]}")
        else:
            fail(f"ПолучитьРасшифровкуПлатежей упала: {e}")
    print("OK: экспорт-функция отработала")

    tz_dokumenty = result.РасшифровкаДокументы
    tz_dogovory  = result.Расшифровка

    rows_doc = tz_dokumenty.Количество()
    rows_dog = tz_dogovory.Количество()
    print(f"OK: РасшифровкаДокументы = {rows_doc} строк, Расшифровка = {rows_dog} строк")

    if rows_doc != EXPECTED_DOCS:
        fail(f"Ожидалось {EXPECTED_DOCS} документов, получено {rows_doc}")
    if rows_dog == 0 or rows_dog > rows_doc:
        fail(f"Свёртка некорректна: {rows_dog} строк (должно 1..{rows_doc})")

    # === В детализации есть СтатьяДДС и микс статей ===
    first_doc = tz_dokumenty.Получить(0)
    if not hasattr(first_doc, "СтатьяДвиженияДенежныхСредств"):
        fail("ТЗ РасшифровкаДокументы не содержит колонку 'СтатьяДвиженияДенежныхСредств'")
    distinct_statei = set()
    docs_our_statia = 0
    for i in range(rows_doc):
        row = tz_dokumenty.Получить(i)
        st_str = str(erp.String(row.СтатьяДвиженияДенежныхСредств))
        distinct_statei.add(st_str)
        if STATIA_NAME in st_str:
            docs_our_statia += 1
    print(f"OK: разных статей в детализации = {len(distinct_statei)}")
    if len(distinct_statei) != EXPECTED_DISTINCT_STATEI:
        fail(f"Ожидалось {EXPECTED_DISTINCT_STATEI} разных статей, получено {len(distinct_statei)}: {distinct_statei}")
    if docs_our_statia != EXPECTED_DOCS_OUR_STATIA:
        fail(f"Документов нашей статьи {docs_our_statia} != эталон {EXPECTED_DOCS_OUR_STATIA}")
    print(f"OK: документов по '{STATIA_NAME}' = {docs_our_statia}")

    # === Сумма по агрегату ===
    total = 0.0
    has_partner = False
    has_glob_div = False
    for i in range(rows_dog):
        row = tz_dogovory.Получить(i)
        total += float(row.Сумма)
        partner_str = str(erp.String(row.Контрагент))
        div_str = str(erp.String(row.Подразделение))
        if "Європабуд" in partner_str:
            has_partner = True
        if "Глобино-2" in div_str:
            has_glob_div = True

    print(f"OK: сумма по свёртке = {total:,.2f}")

    if abs(total - EXPECTED_TOTAL) > 0.5:
        fail(f"Сумма {total:.2f} != эталон {EXPECTED_TOTAL:.2f}")
    if not has_partner:
        fail("'Європабуд' не найден среди контрагентов")
    if not has_glob_div:
        fail("'Глобино-2' не найдено среди подразделений")
    print("OK: 'Європабуд' и 'Глобино-2' найдены")

    # === Колонка Обработан проставляется автоматически: Истина для моноцелевых ===
    first_row = tz_dogovory.Получить(0)
    if not hasattr(first_row, "Обработан"):
        fail("ТЗ Расшифровка не содержит колонку 'Обработан'")
    obr_count = 0
    obr_match_flag = 0  # Обработан = НЕ ИспользоватьВДругихСтатьях
    for i in range(rows_dog):
        row = tz_dogovory.Получить(i)
        is_obr = bool(row.Обработан)
        is_multi = bool(row.ИспользоватьВДругихСтатьях)
        if is_obr:
            obr_count += 1
        if is_obr == (not is_multi):
            obr_match_flag += 1
    print(f"OK: Обработан=Истина у {obr_count} строк (моноцелевых)")
    if obr_count != rows_dog - EXPECTED_MULTI_CONTRACTS:
        fail(f"Обработан должен быть Истина у {rows_dog - EXPECTED_MULTI_CONTRACTS} (моноцелевых), а у {obr_count}")
    if obr_match_flag != rows_dog:
        fail(f"Логика Обработан = НЕ ИспользоватьВДругихСтатьях нарушена: {obr_match_flag}/{rows_dog}")
    print("OK: Обработан = НЕ ИспользоватьВДругихСтатьях для всех строк")

    # === Колонка ТекущаяСтатья (А_СтатьяДвиженияДенежныхСредствОснавнаяПринудительно) ===
    if not hasattr(first_row, "А_СтатьяДвиженияДенежныхСредствОснавнаяПринудительно"):
        fail("ТЗ Расшифровка не содержит колонку 'А_СтатьяДвиженияДенежныхСредствОснавнаяПринудительно'")
    print("OK: колонка А_СтатьяДДСОснавнаяПринудительно присутствует")

    # === Флаг ИспользоватьВДругихСтатьях ===
    if not hasattr(first_row, "ИспользоватьВДругихСтатьях"):
        fail("ТЗ Расшифровка не содержит колонку 'ИспользоватьВДругихСтатьях'")
    print("OK: колонка ИспользоватьВДругихСтатьях присутствует")

    multi_count = 0
    multi_partners_seen = set()
    for i in range(rows_dog):
        row = tz_dogovory.Получить(i)
        if bool(row.ИспользоватьВДругихСтатьях):
            multi_count += 1
            partner_str = str(erp.String(row.Контрагент))
            for needle in EXPECTED_MULTI_PARTNERS:
                if needle in partner_str:
                    multi_partners_seen.add(needle)

    print(f"OK: договоров с флагом 'многоцелевой' = {multi_count}")
    if multi_count != EXPECTED_MULTI_CONTRACTS:
        fail(f"Ожидалось {EXPECTED_MULTI_CONTRACTS} многоцелевых, получено {multi_count}")

    missing = [p for p in EXPECTED_MULTI_PARTNERS if p not in multi_partners_seen]
    if missing:
        fail(f"Ожидаемые многоцелевые партнёры не найдены: {missing}")
    print(f"OK: среди многоцелевых найдены: {sorted(multi_partners_seen)}")

    # === В РасшифровкаДокументы первая строка (по сортировке Дата) ===
    first_doc = tz_dokumenty.Получить(0)
    print(f"  пример док: {erp.String(first_doc.ДокументДенежныхСредств)[:60]}, "
          f"сумма={float(first_doc.СуммаДокумента):,.2f}")

    print("\nSUCCESS: GREEN-тест прошёл")


if __name__ == "__main__":
    main()
