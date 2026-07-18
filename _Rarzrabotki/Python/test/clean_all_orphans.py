# -*- coding: utf-8 -*-
"""
Очистка осиротевших записей во всех регистрах после удаления РегистраторовРасчетов.
"""
import win32com.client
import sys

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Подключение к BaseERP...")
conn = v8.Connect(CONN_ERP)
print("OK\n")

query = conn.NewObject("Запрос")

registers = [
    "РасчетыСПоставщикамиПланОплат",
    "РасчетыСКлиентамиПланОплат",
    "РасчетыСКлиентамиПланОтгрузок",
    "РасчетыСПоставщикамиПланПоставок",
]

for reg_name in registers:
    print(f"--- {reg_name} ---")

    # Подсчёт
    query.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрНакопления.{reg_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
    result = query.Execute().Выбрать()
    result.Следующий()
    orphan_count = result.Кол
    print(f"  Осиротевших: {orphan_count}")

    if orphan_count == 0:
        continue

    # Уникальные регистраторы
    query.Text = f"ВЫБРАТЬ РАЗЛИЧНЫЕ Регистратор КАК Рег ИЗ РегистрНакопления.{reg_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
    result = query.Execute().Выбрать()

    orphan_refs = []
    while result.Следующий():
        orphan_refs.append(result.Рег)

    print(f"  Уникальных регистраторов: {len(orphan_refs)}")

    # Очистка
    reg_obj = getattr(conn.РегистрыНакопления, reg_name)
    cleaned = 0
    errors = 0
    for ref in orphan_refs:
        try:
            record_set = reg_obj.СоздатьНаборЗаписей()
            record_set.Отбор.Регистратор.Установить(ref)
            record_set.Записать()
            cleaned += 1
            if cleaned % 100 == 0:
                print(f"  Очищено: {cleaned}...")
        except Exception as e:
            errors += 1
            if errors <= 3:
                print(f"  ОШИБКА: {str(e)[:100]}")

    print(f"  Очищено: {cleaned}, ошибок: {errors}")

    # Проверка
    query.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрНакопления.{reg_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
    result = query.Execute().Выбрать()
    result.Следующий()
    print(f"  Осталось: {result.Кол}")
    print()

print("=" * 50)
print("ГОТОВО")
