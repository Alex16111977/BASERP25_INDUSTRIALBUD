# -*- coding: utf-8 -*-
"""
Backfill реквізиту СотрудникЕРП у довіднику А_СотрудникиКазна.

Стратегія: НЕ дублюємо логіку матчингу. Просто «торкаємось» кожного запису з
порожнім СотрудникЕРП через COM (ПолучитьОбъект + Записать) — спрацьовує штатний
обробник Catalog.А_СотрудникиКазна.ПередЗаписью, який сам ставить СотрудникЕРП
за алгоритмом документа А_ОтражениеЗПпоКазне (ИНН -> Сотрудники.А_КодПоДРФО,
fallback — код Казни -> ФізЛицо -> Сотрудник). LINK-ONLY: нічого не створюється.

Пише з ОбменДанными.Загрузка=Истина, щоб не реєструвати 2796 записів на обмін.
Ідемпотентно: повторний прогін торкається лише тих, що лишились порожніми.
"""
import sys
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)
print("Connected to BaseERP")


def count_empty():
    q = erp.NewObject("Запрос")
    q.Text = (
        'ВЫБРАТЬ КОЛИЧЕСТВО(С.Ссылка) КАК Кол '
        'ИЗ Справочник.А_СотрудникиКазна КАК С '
        'ГДЕ НЕ С.ЭтоГруппа '
        'И С.СотрудникЕРП = ЗНАЧЕНИЕ(Справочник.Сотрудники.ПустаяСсылка)'
    )
    sel = q.Execute().Выбрать()
    sel.Следующий()
    return int(sel.Кол)


before = count_empty()
print(f"Порожніх СотрудникЕРП ДО: {before}")

# Вибірка посилань на порожні записи
q = erp.NewObject("Запрос")
q.Text = (
    'ВЫБРАТЬ С.Ссылка КАК Ссылка '
    'ИЗ Справочник.А_СотрудникиКазна КАК С '
    'ГДЕ НЕ С.ЭтоГруппа '
    'И С.СотрудникЕРП = ЗНАЧЕНИЕ(Справочник.Сотрудники.ПустаяСсылка)'
)
tz = q.Execute().Выгрузить()
total = tz.Количество()
print(f"До обробки: {total}")

written = 0
errors = 0
first_errors = []
for i in range(total):
    ref = tz.Получить(i).Ссылка
    try:
        obj = ref.ПолучитьОбъект()
        if obj is None:
            errors += 1
            continue
        obj.ОбменДанными.Загрузка = True
        obj.Записать()
        written += 1
    except Exception as e:
        errors += 1
        if len(first_errors) < 5:
            msg = e.excepinfo[2] if getattr(e, "excepinfo", None) else str(e)
            first_errors.append(msg)

print(f"Записано (torknuto): {written}, помилок: {errors}")
for m in first_errors:
    print("  ERR:", m)

after = count_empty()
filled = before - after
print(f"Порожніх СотрудникЕРП ПІСЛЯ: {after}")
print(f"Заповнено цим прогоном: {filled}")

erp = None
v8 = None
print("DONE")
