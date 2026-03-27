# -*- coding: utf-8 -*-
"""
Тест: Маппинг ФинАгентов между Казной (_Финагенты) и ERP (А_ФинАгенты)
"""
import win32com.client
import pythoncom

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_KAZNA = 'Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"'

def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    print("Подключение к ERP...")
    conn_erp = v8.Connect(CONN_ERP)
    print("OK")

    print("Подключение к Казне...")
    conn_kazna = v8.Connect(CONN_KAZNA)
    print("OK\n")

    # === 1. Получить все ФинАгенты ERP ===
    print("=== ФинАгенты ERP (А_ФинАгенты) ===")
    q_erp = conn_erp.NewObject("Запрос")
    q_erp.Текст = (
        "ВЫБРАТЬ ФА.Код, ФА.Наименование "
        "ИЗ Справочник.А_ФинАгенты КАК ФА "
        "ГДЕ НЕ ФА.ПометкаУдаления"
    )
    res_erp = q_erp.Выполнить().Выгрузить()

    erp_by_code = {}
    erp_by_name = {}
    for i in range(res_erp.Количество()):
        row = res_erp.Получить(i)
        code = str(row.Код).strip()
        name = str(row.Наименование).strip()
        erp_by_code[code] = name
        erp_by_name[name] = code
        print(f"  ERP: Код={code}, Наименование={name}")
    print(f"  Итого ERP: {res_erp.Количество()}\n")

    # === 2. Получить все ФинАгенты Казна ===
    print("=== ФинАгенты Казна (_Финагенты) ===")
    q_kazna = conn_kazna.NewObject("Запрос")
    q_kazna.Текст = (
        "ВЫБРАТЬ ФА.Код, ФА.Наименование "
        "ИЗ Справочник._Финагенты КАК ФА "
        "ГДЕ НЕ ФА.ПометкаУдаления"
    )
    res_kazna = q_kazna.Выполнить().Выгрузить()

    kazna_items = []
    for i in range(res_kazna.Количество()):
        row = res_kazna.Получить(i)
        code = str(row.Код).strip()
        name = str(row.Наименование).strip()
        kazna_items.append((code, name))
        print(f"  Казна: Код={code}, Наименование={name}")
    print(f"  Итого Казна: {len(kazna_items)}\n")

    # === 3. Тест маппинга ===
    print("=== Маппинг Казна -> ERP ===")
    matched_code = 0
    matched_name = 0
    unmatched = []

    for code, name in kazna_items:
        if code in erp_by_code:
            matched_code += 1
            print(f"  [КОД]  Казна '{name}' (код={code}) -> ERP '{erp_by_code[code]}'")
        elif name in erp_by_name:
            matched_name += 1
            print(f"  [ИМЯ]  Казна '{name}' (код={code}) -> ERP код={erp_by_name[name]}")
        else:
            unmatched.append((code, name))
            print(f"  [???]  Казна '{name}' (код={code}) -> НЕ НАЙДЕН в ERP!")

    total = len(kazna_items)
    print(f"\n=== РЕЗУЛЬТАТ МАППИНГА ===")
    print(f"  Всего Казна: {total}")
    print(f"  Совпали по коду: {matched_code}")
    print(f"  Совпали по имени: {matched_name}")
    print(f"  Не найдены: {len(unmatched)}")
    if total > 0:
        print(f"  Процент совпадений: {(matched_code + matched_name) / total * 100:.1f}%")

    # === 4. Маппинг Направлений ===
    print("\n=== Маппинг Направлений Казна -> ERP.НаправленияДеятельности ===")
    q_dir_kazna = conn_kazna.NewObject("Запрос")
    q_dir_kazna.Текст = (
        "ВЫБРАТЬ Н.Код, Н.Наименование "
        "ИЗ Справочник.Направления КАК Н "
        "ГДЕ НЕ Н.ПометкаУдаления"
    )
    res_dir_kazna = q_dir_kazna.Выполнить().Выгрузить()

    q_dir_erp = conn_erp.NewObject("Запрос")
    q_dir_erp.Текст = (
        "ВЫБРАТЬ НД.Код, НД.Наименование "
        "ИЗ Справочник.НаправленияДеятельности КАК НД "
        "ГДЕ НЕ НД.ПометкаУдаления"
    )
    res_dir_erp = q_dir_erp.Выполнить().Выгрузить()

    erp_dir_by_code = {}
    erp_dir_by_name = {}
    for i in range(res_dir_erp.Количество()):
        row = res_dir_erp.Получить(i)
        c = str(row.Код).strip()
        n = str(row.Наименование).strip()
        erp_dir_by_code[c] = n
        erp_dir_by_name[n] = c

    dir_matched = 0
    for i in range(res_dir_kazna.Количество()):
        row = res_dir_kazna.Получить(i)
        c = str(row.Код).strip()
        n = str(row.Наименование).strip()
        if c in erp_dir_by_code:
            dir_matched += 1
            print(f"  [КОД] Казна '{n}' -> ERP '{erp_dir_by_code[c]}'")
        elif n in erp_dir_by_name:
            dir_matched += 1
            print(f"  [ИМЯ] Казна '{n}' -> ERP код={erp_dir_by_name[n]}")
        else:
            print(f"  [???] Казна '{n}' (код={c}) -> НЕ НАЙДЕН!")

    print(f"  Совпали: {dir_matched} из {res_dir_kazna.Количество()}")

    print("\n=== ТЕСТ ЗАВЕРШЁН УСПЕШНО ===")

if __name__ == "__main__":
    main()
