# -*- coding: utf-8 -*-
"""
Диагностика: найти ВСЕ места, где документ СписаниеБезналичныхДС 00000002640
ссылается на НеизвестныйПартнер — включая движения по регистрам.
"""
import win32com.client
import pythoncom
import sys

CONN_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
DOC_NUMBER = "00000002640"


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_STRING)
    print("[OK] Подключение к BaseERP")
    return conn


def query(conn, text, params=None):
    q = conn.NewObject("Запрос")
    q.Текст = text
    if params:
        for k, v in params.items():
            q.УстановитьПараметр(k, v)
    try:
        res = q.Выполнить()
    except Exception as e:
        return None  # запрос не подходит для этого регистра
    if res.Пустой():
        return []
    sel = res.Выбрать()
    rows = []
    while sel.Следующий():
        row = {}
        for i in range(res.Колонки.Количество()):
            col = res.Колонки.Получить(i).Имя
            row[col] = getattr(sel, col)
        rows.append(row)
    return rows


def main():
    conn = connect()

    # Находим ссылку документа
    rows = query(conn,
        'ВЫБРАТЬ Д.Ссылка КАК ДокСсылка, Д.Партнер.Наименование КАК ПартнерИмя, Д.Проведен '
        'ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Д '
        'ГДЕ Д.Номер = &Номер',
        {"Номер": DOC_NUMBER}
    )
    if not rows:
        print("Документ не найден!")
        return
    doc_ref = rows[0]["ДокСсылка"]
    print(f"Партнер шапки: {rows[0]['ПартнерИмя']}")
    print(f"Проведен: {rows[0]['Проведен']}")

    unknown_partner = conn.Справочники.Партнеры.НеизвестныйПартнер

    # ============ 1. Проверяем ТЧ РасшифровкаПлатежа ============
    print("\n" + "=" * 60)
    print("1. Табличная часть РасшифровкаПлатежа")
    print("=" * 60)
    tc_rows = query(conn,
        'ВЫБРАТЬ Р.НомерСтроки КАК НС, Р.Партнер.Наименование КАК ПартнерИмя '
        'ИЗ Документ.СписаниеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК Р '
        'ГДЕ Р.Ссылка = &Док',
        {"Док": doc_ref}
    )
    if tc_rows:
        for r in tc_rows:
            status = "[!!!]" if r['ПартнерИмя'] == "Неизвестный партнер" else "[OK]"
            print(f"  {status} Строка {r['НС']}: Партнер = {r['ПартнерИмя']}")
    else:
        print("  Нет строк")

    # ============ 2. Проверяем движения — перебираем через объект ============
    print("\n" + "=" * 60)
    print("2. Движения по регистрам (все наборы записей)")
    print("=" * 60)

    doc_obj = doc_ref.ПолучитьОбъект()
    movements = doc_obj.Движения

    # Перебираем все наборы движений через Количество/Получить
    total_regs = movements.Количество()
    print(f"Всего наборов движений: {total_regs}")

    regs_with_unknown = []

    for j in range(total_regs):
        reg_set = movements.Получить(j)
        table_name = str(reg_set)
        reg_set.Прочитать()
        cnt = reg_set.Количество()
        if cnt == 0:
            continue

        # Проверяем — есть ли поле Партнер в записях?
        has_partner = False
        try:
            test_rec = reg_set.Получить(0)
            _ = test_rec.Партнер
            has_partner = True
        except:
            pass

        if not has_partner:
            print(f"  {table_name}: {cnt} записей (нет поля Партнер)")
            continue

        # Есть Партнер — проверяем все записи
        unknown_count = 0
        for m in range(cnt):
            record = reg_set.Получить(m)
            if record.Партнер == unknown_partner:
                unknown_count += 1

        if unknown_count > 0:
            print(f"  [!!!] {table_name}: {cnt} записей, {unknown_count} с НеизвестнымПартнером")
            regs_with_unknown.append((table_name, unknown_count, cnt))
        else:
            print(f"  [OK] {table_name}: {cnt} записей, Партнер корректный")

    # ============ 3. НайтиПоСсылкам — полная проверка ============
    print("\n" + "=" * 60)
    print("3. НайтиПоСсылкам — ищем документ среди всех ссылок")
    print("=" * 60)

    arr = conn.NewObject("Массив")
    arr.Добавить(unknown_partner)
    found_table = conn.НайтиПоСсылкам(arr)
    total_refs = found_table.Количество()
    print(f"Всего ссылок на НеизвестныйПартнер: {total_refs}")

    # Ищем нашу ссылку среди результатов
    doc_found_in_refs = False
    for i in range(total_refs):
        row = found_table.Получить(i)
        data = row.Данные
        # Сравниваем ссылки напрямую
        if data == doc_ref:
            doc_found_in_refs = True
            print(f"  [!!!] Документ найден как Данные (шапка/ТЧ)")
            break

    if not doc_found_in_refs:
        print(f"  [OK] Документ НЕ найден как прямая ссылка")

    # ============ 4. Итог ============
    print("\n" + "=" * 60)
    print("ИТОГ")
    print("=" * 60)

    if rows[0]['ПартнерИмя'] == "Неизвестный партнер":
        print("[FAIL] Шапка: Партнер = НеизвестныйПартнер")
    else:
        print(f"[OK] Шапка: Партнер = {rows[0]['ПартнерИмя']}")

    if regs_with_unknown:
        print("[FAIL] Движения с НеизвестнымПартнером:")
        for name, unk, total in regs_with_unknown:
            print(f"  - {name}: {unk}/{total} записей")
        print("\n>>> НУЖНО ПЕРЕПРОВЕСТИ ДОКУМЕНТ, чтобы обновить движения!")
    else:
        print("[OK] Движения: НеизвестныйПартнер не найден")

    if doc_found_in_refs:
        print("[FAIL] НайтиПоСсылкам: документ ВСЁ ЕЩЁ ссылается")
    else:
        print("[OK] НайтиПоСсылкам: документ чист")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nОШИБКА: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        pythoncom.CoUninitialize()
