# -*- coding: utf-8 -*-
"""
Создать/перепровести документы А_ПереносДвиженийИзКазны за январь, февраль, март 2026.
"""

import sys, io, datetime, time
import win32com.client
import pythoncom

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

MONTHS = [
    datetime.datetime(2026, 1, 1),
    datetime.datetime(2026, 2, 1),
    datetime.datetime(2026, 3, 1),
]

def last_day_of_month(d):
    next_month = d.replace(day=28) + datetime.timedelta(days=4)
    return next_month - datetime.timedelta(days=next_month.day)

print("=" * 70)
print("Загрузка А_ПереносДвиженийИзКазны за Q1 2026")
print("=" * 70)

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)
print("[OK] Подключение к BaseERP\n")

org = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")

results = []
for month in MONTHS:
    label = month.strftime('%Y-%m')
    print(f"--- {label} ---")

    q = erp.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1
        Док.Ссылка КАК Ссылка, Док.Проведен КАК Проведен
    ИЗ Документ.А_ПереносДвиженийИзКазны КАК Док
    ГДЕ Док.Месяц = &Месяц И НЕ Док.ПометкаУдаления
    """
    q.УстановитьПараметр("Месяц", month)
    res = q.Выполнить()

    if not res.Пустой():
        sel = res.Выбрать(); sel.Следующий()
        doc = sel.Ссылка.ПолучитьОбъект()
        if sel.Проведен:
            doc.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        print(f"  Найден: {sel.Ссылка}")
    else:
        doc = erp.Документы.А_ПереносДвиженийИзКазны.СоздатьДокумент()
        last = last_day_of_month(month)
        doc.Дата = datetime.datetime(last.year, last.month, last.day, 23, 59, 59)
        doc.Месяц = month
        doc.Организация = org
        doc.Записать()
        print(f"  Создан: {doc.Ссылка}")

    t = time.time()
    try:
        doc.ЗаполнитьИзКазны()
        t_fill = time.time() - t
    except Exception as e:
        print(f"  [ERR fill] {e}")
        results.append((label, "FILL FAILED", 0, 0))
        continue

    rows = doc.СтрокиДвижений.Количество()
    print(f"  Заполнено: {rows} строк за {t_fill:.1f}с")

    t = time.time()
    try:
        doc.Записать(erp.РежимЗаписиДокумента.Проведение)
        t_post = time.time() - t
    except Exception as e:
        print(f"  [ERR post] {e}")
        results.append((label, "POST FAILED", rows, t_fill))
        continue

    print(f"  Проведено за {t_post:.1f}с\n")
    results.append((label, "OK", rows, t_fill + t_post))

print("=" * 70)
print(f"{'Месяц':10s} {'Статус':15s} {'Строк':>8s} {'Время':>8s}")
print("-" * 70)
total_rows = 0
total_time = 0
for label, status, rows, secs in results:
    print(f"{label:10s} {status:15s} {rows:>8d} {secs:>7.1f}с")
    total_rows += rows
    total_time += secs
print("-" * 70)
print(f"{'ИТОГО':10s} {'':15s} {total_rows:>8d} {total_time:>7.1f}с")
print("=" * 70)
