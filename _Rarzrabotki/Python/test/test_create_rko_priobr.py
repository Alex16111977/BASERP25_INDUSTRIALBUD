# -*- coding: utf-8 -*-
"""
Тест-верификация: проверка документов ПКО, РКО и Приобретений,
созданных обработкой СоздатьПриходДенегПоФинАгенту.

Тестовый документ: А_ПриходДенегОтФинАгента N00023533 от 01.12.2025
Ожидаемый результат: 12 ПКО + 12 РКО + 12 Приобретений

ВАЖНО: Сначала запустите обработку из 1С, затем этот тест для проверки.
COM-подключение НЕ может создавать документы из-за А_СобытияОбъектов
(ExternalConnection=false). Обработка работает только из 1С клиента.
"""

import win32com.client
import pythoncom
import sys

# ============ НАСТРОЙКИ ============
CONN_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
DOC_NUMBER = "N00023533"


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_STRING)
    print(f"[OK] Подключение к BaseERP установлено")
    return conn


def execute_query(conn, text, params=None):
    """Выполнить запрос 1С, вернуть список словарей."""
    q = conn.NewObject("Запрос")
    q.Текст = text
    if params:
        for k, v in params.items():
            q.УстановитьПараметр(k, v)
    result = q.Выполнить()
    if result.Пустой():
        return []
    sel = result.Выбрать()
    rows = []
    while sel.Следующий():
        row = {}
        for i in range(result.Колонки.Количество()):
            col_name = result.Колонки.Получить(i).Имя
            row[col_name] = getattr(sel, col_name)
        rows.append(row)
    return rows


# ============ ШАГ 1: Найти документ-источник ============
def find_source_doc(conn):
    print("\n=== ШАГ 1: Поиск документа-источника ===")
    rows = execute_query(conn,
        'ВЫБРАТЬ Ссылка, Номер, Дата, СуммаДокумента '
        'ИЗ Документ.А_ПриходДенегОтФинАгента '
        'ГДЕ Номер = &Номер',
        {"Номер": DOC_NUMBER}
    )
    if not rows:
        print(f"[FAIL] Документ {DOC_NUMBER} не найден!")
        return None
    doc = rows[0]
    print(f"[OK] Найден: {doc['Номер']} от {doc['Дата']}, сумма={doc['СуммаДокумента']}")
    return doc


# ============ ШАГ 2: Исходные строки РасшифровкаПлатежа ============
def get_source_rows(conn, doc_ref):
    print("\n=== ШАГ 2: Строки РасшифровкаПлатежа с Сумма > 0 ===")
    rows = execute_query(conn,
        'ВЫБРАТЬ НомерСтроки, Контрагент.Наименование КАК Контрагент, '
        'Сумма, СуммаКомиссии '
        'ИЗ Документ.А_ПриходДенегОтФинАгента.РасшифровкаПлатежа '
        'ГДЕ Ссылка = &Ссылка И Сумма > 0 '
        'УПОРЯДОЧИТЬ ПО НомерСтроки',
        {"Ссылка": doc_ref}
    )
    print(f"[OK] Строк: {len(rows)}")
    total_sum = sum(r['Сумма'] for r in rows)
    total_commission = sum(r['СуммаКомиссии'] for r in rows)
    print(f"     Итого сумма: {total_sum:.2f}, комиссия: {total_commission:.2f}")
    for r in rows:
        print(f"  #{r['НомерСтроки']}: {r['Контрагент']} "
              f"сумма={r['Сумма']:.2f} комиссия={r['СуммаКомиссии']:.2f}")
    return rows, total_sum, total_commission


# ============ ШАГ 3: Верификация созданных документов ============
def verify_pko(conn, doc_ref, expected_count, expected_sum):
    """Проверить созданные ПКО."""
    print(f"\n--- ПКО (ПриходныйКассовыйОрдер) ---")
    rows = execute_query(conn,
        'ВЫБРАТЬ Номер, Дата, СуммаДокумента, '
        'Контрагент.Наименование КАК Контрагент, '
        'Проведен, ХозяйственнаяОперация, '
        'А_КодСтрокиПриходДенегОтФинАгента КАК НомерСтроки, '
        'А_ПроцентФинАгента, А_СуммаКомиссии, '
        'А_НаправлениеДеятельности, А_Подразделение, '
        'Касса, Валюта, Подразделение, НаправлениеДеятельности '
        'ИЗ Документ.ПриходныйКассовыйОрдер '
        'ГДЕ А_ПриходДенегОтФинАгента = &Осн И НЕ ПометкаУдаления '
        'УПОРЯДОЧИТЬ ПО А_КодСтрокиПриходДенегОтФинАгента',
        {"Осн": doc_ref}
    )

    total = 0.0
    errors = []
    for r in rows:
        total += r['СуммаДокумента']
        status = "V" if r['Проведен'] else "X"
        print(f"  [{status}] {r['Номер']}: {r['СуммаДокумента']:.2f} - {r['Контрагент']} "
              f"(стр.{r['НомерСтроки']}, хозоп={r['ХозяйственнаяОперация']})")

        # Проверить ХозОперацию
        хозоп_str = str(r['ХозяйственнаяОперация'])
        if "ВозвратДенежныхСредствОтПоставщика" not in хозоп_str and "Возврат ДС от поставщика" not in хозоп_str:
            errors.append(f"ПКО {r['Номер']}: неверная ХозОперация '{хозоп_str}'")

        if not r['Проведен']:
            errors.append(f"ПКО {r['Номер']}: не проведён")

    print(f"  Итого: {len(rows)} шт., сумма {total:.2f}")

    if len(rows) != expected_count:
        errors.append(f"Количество ПКО: ожидалось {expected_count}, получено {len(rows)}")
    if abs(total - expected_sum) > 0.01:
        errors.append(f"Сумма ПКО: ожидалось {expected_sum:.2f}, получено {total:.2f}")

    for e in errors:
        print(f"  [FAIL] {e}")
    if not errors:
        print(f"  [PASS] ПКО OK")

    return rows, errors


def verify_rko(conn, doc_ref, expected_count, expected_sum):
    """Проверить созданные РКО."""
    print(f"\n--- РКО (РасходныйКассовыйОрдер) ---")
    rows = execute_query(conn,
        'ВЫБРАТЬ Номер, Дата, СуммаДокумента, '
        'Контрагент.Наименование КАК Контрагент, '
        'Проведен, ХозяйственнаяОперация, '
        'А_ПриходныйКассовыйОрдерОтФинАгента КАК СвязьПКО '
        'ИЗ Документ.РасходныйКассовыйОрдер '
        'ГДЕ А_ПриходДенегОтФинАгента = &Осн И НЕ ПометкаУдаления '
        'УПОРЯДОЧИТЬ ПО Номер',
        {"Осн": doc_ref}
    )

    total = 0.0
    errors = []
    no_pko_count = 0
    for r in rows:
        total += r['СуммаДокумента']
        status = "V" if r['Проведен'] else "X"
        has_pko = conn.ЗначениеЗаполнено(r['СвязьПКО'])
        pko_mark = "ПКО+" if has_pko else "ПКО-"
        print(f"  [{status}] {r['Номер']}: {r['СуммаДокумента']:.2f} - {r['Контрагент']} ({pko_mark})")

        if not has_pko:
            no_pko_count += 1
        if not r['Проведен']:
            errors.append(f"РКО {r['Номер']}: не проведён")

    print(f"  Итого: {len(rows)} шт., сумма {total:.2f}")

    if len(rows) != expected_count:
        errors.append(f"Количество РКО: ожидалось {expected_count}, получено {len(rows)}")
    if abs(total - expected_sum) > 0.01:
        errors.append(f"Сумма РКО: ожидалось {expected_sum:.2f}, получено {total:.2f}")
    if no_pko_count > 0:
        errors.append(f"РКО без связки с ПКО: {no_pko_count}")

    for e in errors:
        print(f"  [FAIL] {e}")
    if not errors:
        print(f"  [PASS] РКО OK")

    return rows, errors


def verify_priobr(conn, doc_ref, expected_count, expected_sum):
    """Проверить созданные Приобретения."""
    print(f"\n--- Приобретения (ПриобретениеТоваровУслуг) ---")
    rows = execute_query(conn,
        'ВЫБРАТЬ Номер, Дата, СуммаДокумента, '
        'Контрагент.Наименование КАК Контрагент, '
        'Проведен, ХозяйственнаяОперация, '
        'Курс, Кратность, СуммаВзаиморасчетов, '
        'А_ПриходныйКассовыйОрдерКомиссия КАК СвязьПКО '
        'ИЗ Документ.ПриобретениеТоваровУслуг '
        'ГДЕ А_ПриходДенегОтФинАгента = &Осн И НЕ ПометкаУдаления '
        'УПОРЯДОЧИТЬ ПО Номер',
        {"Осн": doc_ref}
    )

    total = 0.0
    errors = []
    no_pko_count = 0
    for r in rows:
        total += r['СуммаДокумента']
        status = "V" if r['Проведен'] else "X"
        has_pko = conn.ЗначениеЗаполнено(r['СвязьПКО'])
        pko_mark = "ПКО+" if has_pko else "ПКО-"
        print(f"  [{status}] {r['Номер']}: {r['СуммаДокумента']:.2f} - {r['Контрагент']} "
              f"(Курс={r['Курс']}, Кратн={r['Кратность']}, "
              f"СуммаВзаим={r['СуммаВзаиморасчетов']:.2f}, {pko_mark})")

        if not has_pko:
            no_pko_count += 1
        if r['Курс'] != 1:
            errors.append(f"Приобр {r['Номер']}: Курс={r['Курс']} (ожид. 1)")
        if r['Кратность'] != 1:
            errors.append(f"Приобр {r['Номер']}: Кратность={r['Кратность']} (ожид. 1)")
        if abs(r['СуммаВзаиморасчетов'] - r['СуммаДокумента']) > 0.01:
            errors.append(f"Приобр {r['Номер']}: СуммаВзаиморасчетов ({r['СуммаВзаиморасчетов']:.2f}) "
                          f"!= СуммаДокумента ({r['СуммаДокумента']:.2f})")
        if not r['Проведен']:
            errors.append(f"Приобр {r['Номер']}: не проведён")

    print(f"  Итого: {len(rows)} шт., сумма {total:.2f}")

    if len(rows) != expected_count:
        errors.append(f"Количество Приобретений: ожидалось {expected_count}, получено {len(rows)}")
    if abs(total - expected_sum) > 0.01:
        errors.append(f"Сумма Приобретений: ожидалось {expected_sum:.2f}, получено {total:.2f}")
    if no_pko_count > 0:
        errors.append(f"Приобретений без связки с ПКО: {no_pko_count}")

    for e in errors:
        print(f"  [FAIL] {e}")
    if not errors:
        print(f"  [PASS] Приобретения OK")

    return rows, errors


# ============ MAIN ============
def main():
    print("=" * 60)
    print("ТЕСТ-ВЕРИФИКАЦИЯ: ПКО + РКО + Приобретений")
    print(f"по А_ПриходДенегОтФинАгента {DOC_NUMBER}")
    print("=" * 60)
    print("ВАЖНО: Сначала запустите обработку из 1С!")
    print()

    conn = connect()

    # Шаг 1: Найти источник
    doc = find_source_doc(conn)
    if not doc:
        sys.exit(1)
    doc_ref = doc['Ссылка']

    # Шаг 2: Исходные строки
    src_rows, total_sum, total_commission = get_source_rows(conn, doc_ref)
    if not src_rows:
        print("[FAIL] Нет строк с Сумма > 0")
        sys.exit(1)

    expected_count = len(src_rows)
    # Количество строк с комиссией > 0
    rows_with_commission = sum(1 for r in src_rows if r['СуммаКомиссии'] > 0)

    print(f"\n=== ШАГ 3: Верификация созданных документов ===")
    print(f"  Ожидаем: {expected_count} ПКО, {rows_with_commission} РКО, "
          f"{rows_with_commission} Приобретений")

    # Шаг 3: Верификация
    all_errors = []

    pko_rows, pko_errors = verify_pko(conn, doc_ref, expected_count, total_sum)
    all_errors.extend(pko_errors)

    rko_rows, rko_errors = verify_rko(conn, doc_ref, rows_with_commission, total_commission)
    all_errors.extend(rko_errors)

    priobr_rows, priobr_errors = verify_priobr(conn, doc_ref, rows_with_commission, total_commission)
    all_errors.extend(priobr_errors)

    # Итог
    print(f"\n{'=' * 60}")
    print(f"СВОДКА:")
    print(f"  ПКО:          {len(pko_rows)}/{expected_count}")
    print(f"  РКО:          {len(rko_rows)}/{rows_with_commission}")
    print(f"  Приобретений: {len(priobr_rows)}/{rows_with_commission}")

    if all_errors:
        print(f"\nОШИБКИ ({len(all_errors)}):")
        for e in all_errors:
            print(f"  - {e}")
        print(f"\n[FAIL] Тест НЕ пройден.")
    else:
        print(f"\n[PASS] Все проверки пройдены!")

    print(f"\n[DONE] Тест завершён.")


if __name__ == "__main__":
    main()
