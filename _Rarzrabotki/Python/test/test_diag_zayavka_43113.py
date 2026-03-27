# -*- coding: utf-8 -*-
"""
Диагностика: почему ЗаявкаНаРасходованиеДС 00230043113 постоянно тянется
в обработку НеизвестногоПартнера. Шапка и ТЧ уже чистые — ищем в движениях.
"""
import win32com.client
import pythoncom
import sys

CONN_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
DOC_NUMBER = "00230043113"


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
    res = q.Выполнить()
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

    # Находим документ
    rows = query(conn,
        'ВЫБРАТЬ Д.Ссылка КАК ДокСсылка, Д.Партнер.Наименование КАК ПартнерИмя, Д.Проведен '
        'ИЗ Документ.ЗаявкаНаРасходованиеДенежныхСредств КАК Д '
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

    # ============ 1. НайтиПоСсылкам — ищем наш документ ============
    print("\n" + "=" * 60)
    print("1. НайтиПоСсылкам — ищем документ среди ссылок")
    print("=" * 60)

    arr = conn.NewObject("Массив")
    arr.Добавить(unknown_partner)
    found_table = conn.НайтиПоСсылкам(arr)
    total_refs = found_table.Количество()
    print(f"Всего ссылок на НеизвестныйПартнер: {total_refs}")

    doc_found_entries = []
    for i in range(total_refs):
        row = found_table.Получить(i)
        data = row.Данные
        # Прямое сравнение ссылки
        if data == doc_ref:
            meta = row.Метаданные
            doc_found_entries.append(str(meta))
            print(f"  [!!!] Найден! Метаданные: {str(meta)}")

    if not doc_found_entries:
        print("  Документ НЕ найден напрямую")
        print("  Проверяем — может найден через строковое представление...")
        for i in range(total_refs):
            row = found_table.Получить(i)
            data_str = str(row.Данные)
            if DOC_NUMBER in data_str or "43113" in data_str:
                meta = row.Метаданные
                print(f"  [!!!] Найден по номеру: Данные={data_str}, Метаданные={str(meta)}")
                doc_found_entries.append(str(meta))

    if not doc_found_entries:
        print("  [OK] Документ ВООБЩЕ НЕ найден — обработка не должна его тянуть")
        print("  >>> Возможно обработка кеширует старые результаты? Перезапустите.")
    else:
        print(f"\n  Найдено записей: {len(doc_found_entries)}")

    # ============ 2. Движения — перебираем ВСЕ наборы ============
    print("\n" + "=" * 60)
    print("2. Движения документа — ищем НеизвестныйПартнер")
    print("=" * 60)

    doc_obj = doc_ref.ПолучитьОбъект()
    movements = doc_obj.Движения
    total_regs = movements.Количество()
    print(f"Всего наборов движений: {total_regs}")

    regs_with_unknown = []

    for j in range(total_regs):
        reg_set = movements.Получить(j)
        reg_set.Прочитать()
        cnt = reg_set.Количество()
        if cnt == 0:
            continue

        # Проверяем каждое поле записи на наличие НеизвестныйПартнер
        # Пробуем Партнер
        has_partner = False
        try:
            _ = reg_set.Получить(0).Партнер
            has_partner = True
        except:
            pass

        if has_partner:
            unknown_count = 0
            for m in range(cnt):
                record = reg_set.Получить(m)
                if record.Партнер == unknown_partner:
                    unknown_count += 1
            if unknown_count > 0:
                # Пытаемся получить имя регистра
                reg_name = f"Регистр #{j}"
                try:
                    reg_name = str(reg_set)
                except:
                    pass
                print(f"  [!!!] {reg_name}: {unknown_count}/{cnt} записей с НеизвестнымПартнером")
                regs_with_unknown.append((reg_name, unknown_count, cnt))
            else:
                print(f"  [OK] Набор #{j}: {cnt} записей, Партнер корректный")
        else:
            # Нет поля Партнер — проверяем ВСЕ поля на ссылку НеизвестныйПартнер
            # Перебираем свойства первой записи
            found_in_other_field = False
            rec = reg_set.Получить(0)
            # Проверяем типичные поля, где может быть Партнер
            for field_name in ["Партнер", "ПартнерДт", "ПартнерКт", "Получатель", "Отправитель"]:
                try:
                    val = getattr(rec, field_name)
                    if val == unknown_partner:
                        found_in_other_field = True
                        print(f"  [!!!] Набор #{j}: поле {field_name} = НеизвестныйПартнер")
                        regs_with_unknown.append((f"Набор #{j}.{field_name}", 1, cnt))
                except:
                    pass

            if not found_in_other_field:
                pass  # нет полей с Партнером — пропускаем молча

    # ============ 3. Проверяем ВСЕ ТЧ документа ============
    print("\n" + "=" * 60)
    print("3. Все табличные части документа")
    print("=" * 60)

    # Перебираем ТЧ через метаданные
    doc_meta = doc_ref.Метаданные()
    tc_count = doc_meta.ТабличныеЧасти.Количество()
    print(f"Табличных частей: {tc_count}")

    for j in range(tc_count):
        tc_meta = doc_meta.ТабличныеЧасти.Получить(j)
        tc_name = tc_meta.Имя

        # Проверяем — есть ли поле Партнер
        has_partner_attr = False
        for k in range(tc_meta.Реквизиты.Количество()):
            attr = tc_meta.Реквизиты.Получить(k)
            if attr.Имя == "Партнер":
                has_partner_attr = True
                break

        if not has_partner_attr:
            continue

        tc_data = getattr(doc_obj, tc_name)
        tc_cnt = tc_data.Количество()
        if tc_cnt == 0:
            continue

        unknown_count = 0
        for m in range(tc_cnt):
            row = tc_data.Получить(m)
            if row.Партнер == unknown_partner:
                unknown_count += 1

        if unknown_count > 0:
            print(f"  [!!!] ТЧ {tc_name}: {unknown_count}/{tc_cnt} строк с НеизвестнымПартнером")
            regs_with_unknown.append((f"ТЧ.{tc_name}", unknown_count, tc_cnt))
        else:
            print(f"  [OK] ТЧ {tc_name}: {tc_cnt} строк, Партнер корректный")

    # ============ ИТОГ ============
    print("\n" + "=" * 60)
    print("ИТОГ")
    print("=" * 60)

    if doc_found_entries:
        print(f"[FAIL] НайтиПоСсылкам нашла документ в {len(doc_found_entries)} местах:")
        for e in doc_found_entries:
            print(f"  - {e}")
    else:
        print("[OK] НайтиПоСсылкам: документ чист")

    if regs_with_unknown:
        print("[FAIL] Найден НеизвестныйПартнер в:")
        for name, unk, total in regs_with_unknown:
            print(f"  - {name}: {unk}/{total}")
        print("\n>>> НУЖНО ПЕРЕПРОВЕСТИ ДОКУМЕНТ!")
    else:
        print("[OK] НеизвестныйПартнер нигде не найден")

    if not doc_found_entries and not regs_with_unknown:
        print("\n>>> ДОКУМЕНТ ПОЛНОСТЬЮ ЧИСТ")
        print(">>> Если обработка его тянет — проблема в кешировании, перезагрузите обработку")


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
