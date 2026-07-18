# -*- coding: utf-8 -*-
"""
Очистка очередей заданий старой архитектуры взаиморасчетов.
"""
import win32com.client
import sys

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("OK")

query = conn.NewObject("Запрос")

# Проверяем новую архитектуру
query.Text = "ВЫБРАТЬ Константы.НоваяАрхитектураВзаиморасчетов КАК Значение ИЗ Константы КАК Константы"
result = query.Execute().Выбрать()
result.Следующий()
if not result.Значение:
    print("НоваяАрхитектураВзаиморасчетов = ЛОЖЬ! Очистка отменена.")
    sys.exit(1)
print("НоваяАрхитектураВзаиморасчетов = ИСТИНА")

registers = [
    "ЗаданияКРаспределениюРасчетовСПоставщиками",
    "ЗаданияКРаспределениюРасчетовСКлиентами",
]

for reg_name in registers:
    query.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрСведений." + reg_name
    result = query.Execute().Выбрать()
    result.Следующий()
    count = result.Кол
    print(f"\n{reg_name}: {count} записей")

    if count == 0:
        continue

    # Очищаем через менеджер записей набора без отбора = полная очистка
    try:
        manager = getattr(conn.РегистрыСведений, reg_name)
        record_set = manager.СоздатьНаборЗаписей()
        record_set.Записать()
        print(f"  OK")
    except Exception as e:
        print(f"  Способ 1 не сработал: {e}")
        # Способ 2: через выборку и удаление по ключу
        try:
            query.Text = "ВЫБРАТЬ РАЗЛИЧНЫЕ Месяц ИЗ РегистрСведений." + reg_name
            months = query.Execute().Выбрать()
            deleted = 0
            while months.Следующий():
                rs = manager.СоздатьНаборЗаписей()
                rs.Отбор.Месяц.Установить(months.Месяц)
                rs.Записать()
                deleted += 1
            print(f"  OK (по месяцам: {deleted})")
        except Exception as e2:
            print(f"  Способ 2 тоже не сработал: {e2}")

# Проверка
print("\n=== Проверка ===")
for reg_name in registers:
    query.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрСведений." + reg_name
    result = query.Execute().Выбрать()
    result.Следующий()
    print(f"{reg_name}: {result.Кол}")
