# -*- coding: utf-8 -*-
"""
Полная пересборка РегистраторовРасчетов в BaseERP.

Алгоритм:
1. Удалить все документы РегистраторРасчетов (по аналогии с типовым ОчиститьОтПроведенных)
2. Последовательно перепровести все документы за период (хронологически)
3. При перепроведении 1С автоматически создаст новые РегистраторыРасчетов

ВНИМАНИЕ:
- Сделать бэкап базы перед запуском!
- Запускать в монопольном режиме (без пользователей)
- Может занять 10-30 минут в зависимости от объёма документов
"""

import win32com.client
import datetime
import sys
import traceback

# === Подключение ===
v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("=" * 70)
print("ПЕРЕСБОРКА РЕГИСТРАТОРОВ РАСЧЁТОВ")
print("=" * 70)

print("\nПодключение к BaseERP...")
try:
    conn = v8.Connect(CONN_ERP)
    print("OK")
except Exception as e:
    print(f"ОШИБКА подключения: {e}")
    sys.exit(1)

# === ШАГ 1: Подсчёт текущих РегистраторовРасчетов ===
print("\n--- ШАГ 1: Подсчёт РегистраторовРасчетов ---")

query = conn.NewObject("Запрос")
query.Text = """
    ВЫБРАТЬ КОЛИЧЕСТВО(Ссылка) КАК Кол
    ИЗ Документ.РегистраторРасчетов
"""
result = query.Execute().Выбрать()
result.Следующий()
total_registrators = result.Кол
print(f"Всего РегистраторовРасчетов: {total_registrators}")

if total_registrators == 0:
    print("Нечего удалять!")
    sys.exit(0)

# === ШАГ 2: Удаление всех РегистраторовРасчетов ===
print(f"\n--- ШАГ 2: Удаление {total_registrators} РегистраторовРасчетов ---")

query.Text = """
    ВЫБРАТЬ Ссылка ИЗ Документ.РегистраторРасчетов
"""
result = query.Execute().Выбрать()
deleted = 0
errors_delete = 0

while result.Следующий():
    try:
        doc_obj = result.Ссылка.ПолучитьОбъект()
        doc_obj.ОбменДанными.Загрузка = True
        doc_obj.Удалить()
        deleted += 1
        if deleted % 500 == 0:
            print(f"  Удалено: {deleted}/{total_registrators}")
    except Exception as e:
        errors_delete += 1
        if errors_delete <= 5:
            print(f"  ОШИБКА удаления: {e}")

print(f"Удалено: {deleted}, ошибок: {errors_delete}")

# === ШАГ 3: Перепроведение документов ===
print("\n--- ШАГ 3: Перепроведение документов ---")

# Период перепроведения
date_from = "ДАТАВРЕМЯ(2025,12,1)"
date_to = "ДАТАВРЕМЯ(2026,4,1)"

# Типы документов с расчётами (ВзаиморасчетыСервер)
doc_types = [
    "ПоступлениеБезналичныхДенежныхСредств",
    "СписаниеБезналичныхДенежныхСредств",
    "ПриходныйКассовыйОрдер",
    "РасходныйКассовыйОрдер",
    "АвансовыйОтчет",
    "ПриобретениеТоваровУслуг",
    "РеализацияТоваровУслуг",
    "ВозвратТоваровОтКлиента",
    "ВозвратТоваровПоставщику",
    "ВводОстатков",
]

# Собираем все проведённые документы хронологически
print(f"Запрос документов за период {date_from} - {date_to}...")

union_parts = []
for dt in doc_types:
    union_parts.append(f"""
        ВЫБРАТЬ Ссылка, Дата ИЗ Документ.{dt}
        ГДЕ Проведен = ИСТИНА И Дата >= {date_from} И Дата < {date_to}
    """)

query.Text = """
    ВЫБРАТЬ Ссылка, Дата ИЗ (
        {unions}
    ) КАК Все
    УПОРЯДОЧИТЬ ПО Дата
""".format(unions="\n        ОБЪЕДИНИТЬ ВСЕ\n    ".join(union_parts))

try:
    result = query.Execute().Выбрать()
except Exception as e:
    print(f"ОШИБКА запроса документов: {e}")
    # Попробуем по одному типу
    print("Пробую по отдельным типам документов...")
    result = None

if result:
    # Считаем общее количество
    total_docs = 0
    docs_list = []
    while result.Следующий():
        docs_list.append(result.Ссылка)
        total_docs += 1

    print(f"Найдено документов для перепроведения: {total_docs}")

    reposted = 0
    errors_repost = 0
    error_log = []

    for doc_ref in docs_list:
        try:
            doc_obj = doc_ref.ПолучитьОбъект()
            # Отмена проведения
            doc_obj.Записать(conn.NewObject("РежимЗаписиДокумента", "ОтменаПроведения"))
            # Перепроведение
            doc_obj.Записать(conn.NewObject("РежимЗаписиДокумента", "Проведение"))
            reposted += 1
            if reposted % 100 == 0:
                print(f"  Перепроведено: {reposted}/{total_docs}")
        except Exception as e:
            errors_repost += 1
            err_msg = f"{doc_ref}: {str(e)[:100]}"
            error_log.append(err_msg)
            if errors_repost <= 10:
                print(f"  ОШИБКА: {err_msg}")

    print(f"\nПерепроведено: {reposted}, ошибок: {errors_repost}")

    if error_log:
        print(f"\nПервые {min(len(error_log), 20)} ошибок:")
        for err in error_log[:20]:
            print(f"  - {err}")
else:
    # Fallback: перепроводим по типам
    total_reposted = 0
    total_errors = 0

    for dt in doc_types:
        print(f"\n  Тип: {dt}")
        query.Text = f"""
            ВЫБРАТЬ Ссылка ИЗ Документ.{dt}
            ГДЕ Проведен = ИСТИНА И Дата >= {date_from} И Дата < {date_to}
            УПОРЯДОЧИТЬ ПО Дата
        """
        try:
            result = query.Execute().Выбрать()
        except:
            print(f"    Тип не найден, пропускаю")
            continue

        count = 0
        errors = 0
        while result.Следующий():
            try:
                doc_obj = result.Ссылка.ПолучитьОбъект()
                doc_obj.Записать(conn.NewObject("РежимЗаписиДокумента", "ОтменаПроведения"))
                doc_obj.Записать(conn.NewObject("РежимЗаписиДокумента", "Проведение"))
                count += 1
            except Exception as e:
                errors += 1
                if errors <= 3:
                    print(f"    ОШИБКА: {str(e)[:80]}")

        print(f"    Перепроведено: {count}, ошибок: {errors}")
        total_reposted += count
        total_errors += errors

    print(f"\nИтого: перепроведено {total_reposted}, ошибок {total_errors}")

# === ШАГ 4: Проверка результата ===
print("\n--- ШАГ 4: Проверка ---")

query.Text = """
    ВЫБРАТЬ КОЛИЧЕСТВО(Ссылка) КАК Кол
    ИЗ Документ.РегистраторРасчетов
"""
result = query.Execute().Выбрать()
result.Следующий()
new_total = result.Кол
print(f"Новых РегистраторовРасчетов: {new_total} (было: {total_registrators})")

# Проверка 00DL-7185
query.Text = """
    ВЫБРАТЬ РАЗЛИЧНЫЕ Регистратор КАК Рег
    ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам
    ГДЕ ДокументРегистратор.Номер = "00DL-7185"
      И ДокументРегистратор.Дата = ДАТАВРЕМЯ(2025,12,10)
"""
try:
    result = query.Execute().Выбрать()
    count_7185 = 0
    while result.Следующий():
        count_7185 += 1
        print(f"  00DL-7185 -> {result.Рег}")
    print(f"  Регистраторов для 00DL-7185: {count_7185} (ожидаем 1-2, было 3)")
except Exception as e:
    print(f"  Ошибка проверки 00DL-7185: {e}")

print("\n" + "=" * 70)
print("ПЕРЕСБОРКА ЗАВЕРШЕНА")
print("=" * 70)
