# -*- coding: utf-8 -*-
"""
Очистка осиротевших записей в РасчетыСКлиентамиПоСрокам.
"""
import win32com.client
import sys

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Подключение к BaseERP...")
conn = v8.Connect(CONN_ERP)
print("OK")

query = conn.NewObject("Запрос")

# Подсчёт
query.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
result = query.Execute().Выбрать()
result.Следующий()
orphan_count = result.Кол
print(f"Осиротевших записей: {orphan_count}")

if orphan_count == 0:
    print("Нечего очищать!")
    sys.exit(0)

# Получить уникальные регистраторы
query.Text = "ВЫБРАТЬ РАЗЛИЧНЫЕ Регистратор КАК Рег ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
result = query.Execute().Выбрать()

orphan_refs = []
while result.Следующий():
    orphan_refs.append(result.Рег)

print(f"Уникальных регистраторов: {len(orphan_refs)}")

# Очистка
cleaned = 0
errors = 0
for ref in orphan_refs:
    try:
        record_set = conn.РегистрыНакопления.РасчетыСКлиентамиПоСрокам.СоздатьНаборЗаписей()
        record_set.Отбор.Регистратор.Установить(ref)
        record_set.Записать()
        cleaned += 1
        if cleaned % 50 == 0:
            print(f"  Очищено: {cleaned}...")
    except Exception as e:
        errors += 1
        if errors <= 3:
            print(f"  ОШИБКА: {str(e)[:100]}")

print(f"Очищено: {cleaned}, ошибок: {errors}")

# Проверка
query.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
result = query.Execute().Выбрать()
result.Следующий()
print(f"Осталось: {result.Кол} (было: {orphan_count})")

# Проверка ІБ00-000083
query.Text = """
    ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
    ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам
    ГДЕ ДокументРегистратор.Номер = "ІБ00-000083"
      И ДокументРегистратор.Дата = ДАТАВРЕМЯ(2025,12,1)
"""
try:
    result = query.Execute().Выбрать()
    result.Следующий()
    print(f"ІБ00-000083: {result.Кол} строк (было 2, должна быть 1)")
except:
    print("ІБ00-000083: ошибка запроса")

print("ГОТОВО")
