# -*- coding: utf-8 -*-
"""
Тест перепроведения документов по запасам (регистр ТоварыНаСкладах) за период.
Логика: выбираем уникальных регистраторов из РегистрНакопления.ТоварыНаСкладах,
для каждого делаем ОтменаПроведения + Проведение в одном цикле.
"""
import win32com.client
import pythoncom
import datetime
import sys

pythoncom.CoInitialize()

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Подключение к BaseERP...")
conn = v8.Connect(CONN_ERP)
print("OK")

# === ПАРАМЕТРЫ ТЕСТА ===
# Берём короткий период для теста — 1 день
ДАТА_НАЧАЛА = datetime.datetime(2026, 4, 1)
ДАТА_ОКОНЧАНИЯ = datetime.datetime(2026, 4, 1, 23, 59, 59)

print(f"Период: {ДАТА_НАЧАЛА.strftime('%d.%m.%Y')} - {ДАТА_ОКОНЧАНИЯ.strftime('%d.%m.%Y')}")

# === ШАГ 1: Запрос регистраторов ===
query = conn.NewObject("Запрос")
query.Text = (
    "ВЫБРАТЬ РАЗЛИЧНЫЕ "
    "   ТоварыНаСкладах.Регистратор КАК Регистратор, "
    "   ТоварыНаСкладах.Регистратор.Дата КАК Дата "
    "ИЗ "
    "   РегистрНакопления.ТоварыНаСкладах КАК ТоварыНаСкладах "
    "ГДЕ "
    "   ТоварыНаСкладах.Период МЕЖДУ &НачалоПериода И &ОкончаниеПериода "
    "УПОРЯДОЧИТЬ ПО "
    "   Дата"
)
query.SetParameter("НачалоПериода", ДАТА_НАЧАЛА)
query.SetParameter("ОкончаниеПериода", ДАТА_ОКОНЧАНИЯ)

print("Выполняю запрос...")
result = query.Execute().Выбрать()

# Собираем регистраторов в массив
docs = []
while result.Следующий():
    docs.append(result.Регистратор)

print(f"Найдено документов: {len(docs)}")

if len(docs) == 0:
    print("Нет документов для перепроведения.")
    pythoncom.CoUninitialize()
    sys.exit(0)

# === ШАГ 2: Перепроведение ===
mode_unpost = conn.РежимЗаписиДокумента.ОтменаПроведения
mode_post = conn.РежимЗаписиДокумента.Проведение

count_ok = 0
count_err = 0
errors = []

for i, doc_ref in enumerate(docs):
    try:
        doc_obj = doc_ref.ПолучитьОбъект()
        doc_obj.Записать(mode_unpost)
        doc_obj.Записать(mode_post)
        count_ok += 1
        if (i + 1) % 50 == 0:
            print(f"  Обработано: {i + 1} из {len(docs)}...")
    except Exception as e:
        count_err += 1
        err_msg = str(e)[:150]
        errors.append(f"  [{i+1}] {err_msg}")
        if count_err <= 5:
            print(f"  ОШИБКА [{i+1}]: {err_msg}")

# === ШАГ 3: Результаты ===
print(f"\n{'='*60}")
print(f"ИТОГО: успешно {count_ok}, ошибок {count_err}, всего {len(docs)}")

if errors:
    print(f"\nПервые 10 ошибок:")
    for e in errors[:10]:
        print(e)

pythoncom.CoUninitialize()
print("\nГотово.")
