# -*- coding: utf-8 -*-
"""
Тест: замена НеизвестногоПартнера в документе СписаниеБезналичныхДС 00000002640
1. Подключаемся через COM
2. Находим документ, проверяем текущее состояние
3. Находим правильного партнера по контрагенту
4. Заменяем в шапке + РасшифровкаПлатежа
5. Проверяем, что НайтиПоСсылкам больше НЕ находит документ
"""
import win32com.client
import pythoncom
import sys

CONN_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
DOC_NUMBER = "00000002640"


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_STRING)
    print("[OK] Подключение к BaseERP")
    return conn


def v8str(conn, val):
    """Преобразовать значение 1С в строку через XMLСтрока или str."""
    try:
        return str(val)
    except:
        return repr(val)


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

    # ============ ШАГ 1: Текущее состояние ============
    print("\n" + "=" * 60)
    print("ШАГ 1: Текущее состояние документа")
    print("=" * 60)

    rows = query(conn,
        'ВЫБРАТЬ Д.Ссылка КАК ДокСсылка, Д.Партнер КАК Партнер, '
        'Д.Партнер.Наименование КАК ПартнерИмя, '
        'Д.Контрагент КАК Контрагент, '
        'Д.Контрагент.Наименование КАК КонтрагентИмя, '
        'Д.Проведен '
        'ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Д '
        'ГДЕ Д.Номер = &Номер',
        {"Номер": DOC_NUMBER}
    )
    if not rows:
        print("ОШИБКА: Документ не найден!")
        return
    doc_ref = rows[0]["ДокСсылка"]
    print(f"  Партнер шапки: {rows[0]['ПартнерИмя']}")
    print(f"  Контрагент: {rows[0]['КонтрагентИмя']}")
    print(f"  Проведен: {rows[0]['Проведен']}")

    # Табличная часть
    tc_rows = query(conn,
        'ВЫБРАТЬ Р.НомерСтроки КАК НС, '
        'Р.Партнер.Наименование КАК ПартнерИмя '
        'ИЗ Документ.СписаниеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК Р '
        'ГДЕ Р.Ссылка.Номер = &Номер',
        {"Номер": DOC_NUMBER}
    )
    print(f"  Строк в РасшифровкаПлатежа: {len(tc_rows)}")
    for r in tc_rows:
        print(f"    Строка {r['НС']}: Партнер = {r['ПартнерИмя']}")

    # ============ ШАГ 2: Найти правильного партнера ============
    print("\n" + "=" * 60)
    print("ШАГ 2: Определяем правильного партнера")
    print("=" * 60)

    kontragent_ref = rows[0]["Контрагент"]
    kontragent_name = rows[0]["КонтрагентИмя"]

    unknown_partner = conn.Справочники.Партнеры.НеизвестныйПартнер

    # Ищем партнера с именем контрагента
    partner_rows = query(conn,
        'ВЫБРАТЬ П.Ссылка КАК ПартнерСсылка, П.Наименование КАК Имя '
        'ИЗ Справочник.Партнеры КАК П '
        'ГДЕ П.Наименование = &Имя '
        'И П.Ссылка <> &НеизвестныйПартнер',
        {"Имя": kontragent_name,
         "НеизвестныйПартнер": unknown_partner}
    )

    if partner_rows:
        new_partner = partner_rows[0]["ПартнерСсылка"]
        print(f"  Найден партнер: {partner_rows[0]['Имя']}")
    else:
        print(f"  Партнер '{kontragent_name}' не найден — создаём")
        obj = conn.Справочники.Партнеры.СоздатьЭлемент()
        obj.Наименование = kontragent_name
        obj.НаименованиеПолное = kontragent_name
        try:
            obj.Клиент = True
            obj.Поставщик = True
        except:
            pass
        obj.Записать()
        new_partner = obj.Ссылка
        print(f"  Создан партнер: {kontragent_name}")

    # ============ ШАГ 3: Замена партнера ============
    print("\n" + "=" * 60)
    print("ШАГ 3: Заменяем партнера в документе")
    print("=" * 60)

    doc_obj = doc_ref.ПолучитьОбъект()

    # Флаги — чтобы ПередЗаписью не перезаписал
    try:
        doc_obj.А_Обработан = True
        print("  А_Обработан = Истина")
    except Exception as e:
        print(f"  [!] А_Обработан: {e}")
    try:
        doc_obj.А_ОбработанКазна = True
        print("  А_ОбработанКазна = Истина")
    except Exception as e:
        print(f"  [!] А_ОбработанКазна: {e}")

    # Шапка
    doc_obj.Партнер = new_partner
    print(f"  Шапка: Партнер -> {kontragent_name}")

    # Табличная часть РасшифровкаПлатежа
    tc = doc_obj.РасшифровкаПлатежа
    changed_rows = 0
    for i in range(tc.Количество()):
        row = tc.Получить(i)
        row_partner = row.Партнер
        # Сравниваем по ссылке с НеизвестныйПартнер или пустая
        if row_partner == unknown_partner:
            row.Партнер = new_partner
            changed_rows += 1
        else:
            # Проверяем — может пустая ссылка
            try:
                partner_name = query(conn,
                    'ВЫБРАТЬ П.Наименование КАК Имя ИЗ Справочник.Партнеры КАК П ГДЕ П.Ссылка = &Ссылка',
                    {"Ссылка": row_partner}
                )
                if not partner_name or partner_name[0]["Имя"] == "Неизвестный партнер":
                    row.Партнер = new_partner
                    changed_rows += 1
            except:
                pass
    print(f"  РасшифровкаПлатежа: заменено {changed_rows} из {tc.Количество()} строк")

    # Записываем
    doc_obj.ОбменДанными.Загрузка = True
    doc_obj.Записать(conn.РежимЗаписиДокумента.Запись)
    print("  [OK] Документ записан")

    # ============ ШАГ 4: Проверяем результат ============
    print("\n" + "=" * 60)
    print("ШАГ 4: Проверяем результат после записи")
    print("=" * 60)

    rows2 = query(conn,
        'ВЫБРАТЬ Д.Партнер.Наименование КАК ПартнерИмя '
        'ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Д '
        'ГДЕ Д.Номер = &Номер',
        {"Номер": DOC_NUMBER}
    )
    print(f"  Партнер шапки: {rows2[0]['ПартнерИмя']}")

    tc_rows2 = query(conn,
        'ВЫБРАТЬ Р.НомерСтроки КАК НС, '
        'Р.Партнер.Наименование КАК ПартнерИмя '
        'ИЗ Документ.СписаниеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК Р '
        'ГДЕ Р.Ссылка.Номер = &Номер',
        {"Номер": DOC_NUMBER}
    )
    for r in tc_rows2:
        print(f"    Строка {r['НС']}: Партнер = {r['ПартнерИмя']}")

    all_ok = True
    if rows2[0]['ПартнерИмя'] == "Неизвестный партнер":
        print("  [FAIL] Шапка всё ещё НеизвестныйПартнер!")
        all_ok = False
    for r in tc_rows2:
        if r['ПартнерИмя'] == "Неизвестный партнер":
            print(f"  [FAIL] Строка {r['НС']} всё ещё НеизвестныйПартнер!")
            all_ok = False

    # ============ ШАГ 5: НайтиПоСсылкам — находит ли ещё? ============
    print("\n" + "=" * 60)
    print("ШАГ 5: Проверяем через НайтиПоСсылкам")
    print("=" * 60)

    arr = conn.NewObject("Массив")
    arr.Добавить(unknown_partner)
    found_table = conn.НайтиПоСсылкам(arr)

    doc_still_found = False
    for i in range(found_table.Количество()):
        row = found_table.Получить(i)
        data = row.Данные
        # Сравниваем через строковое представление
        data_str = str(data)
        if DOC_NUMBER in data_str or "2640" in data_str:
            doc_still_found = True
            meta = row.Метаданные
            print(f"  [!] Документ ВСЁ ЕЩЁ найден: {data_str} (метаданные: {str(meta)})")

    if not doc_still_found:
        print(f"  [OK] Документ {DOC_NUMBER} НЕ найден через НайтиПоСсылкам")
    else:
        all_ok = False

    # Итог
    print("\n" + "=" * 60)
    if all_ok:
        print(">>> ТЕСТ ПРОЙДЕН: партнер заменён, обработка больше не найдёт документ")
    else:
        print(">>> ТЕСТ НЕ ПРОЙДЕН: см. ошибки выше")
    print("=" * 60)


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
