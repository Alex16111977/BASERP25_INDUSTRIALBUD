# -*- coding: utf-8 -*-
"""Verify: колонка Аналитика1 заполнена строковым представлением:
- для ЕРП/Каса строк → СокрЛП(Аналитика)
- для PL строк → СокрЛП(Комментарий)
"""
import sys, io, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Вызовем сам отчёт ПриКомпоновкеРезультата — тут таблица создаётся внутри.
# Проще: повторим основной шаг вручную с пост-процессингом.
отчёт = conn.Отчеты.А_ОтчетPL.Создать()
НМ = datetime.datetime(2025,12,1)
КМ = datetime.datetime(2025,12,31,23,59,59)
тз = отчёт.ПолучитьОбъединенныеДанные(НМ, КМ, False, True, True, True)

# Колонки Аналитика1 нет в результате ПолучитьОбъединенныеДанные(), она добавляется в ПриКомпоновкеРезультата.
# Эмулируем пост-процессинг для проверки логики:
print(f"Всего строк: {тз.Количество()}\n")

if "Аналитика1" not in [c.Имя for c in тз.Колонки]:
    тз.Колонки.Добавить("Аналитика1", conn.NewObject("ОписаниеТипов", "Строка"))

erp_count = 0
pl_count = 0
empty_count = 0
samples_erp = []
samples_pl = []
for r in тз:
    is_erp_or_kazna = (float(r.СуммаЕРП) != 0 or float(r.СуммаДДСизКазныРасход) != 0 or float(r.СуммаДДСизКазныПриход) != 0)
    if is_erp_or_kazna:
        val = conn.String(r.Аналитика) if r.Аналитика is not None else ""
        r.Аналитика1 = val.strip()
        erp_count += 1
        if len(samples_erp) < 5 and r.Аналитика1:
            samples_erp.append(r.Аналитика1)
    else:
        val = conn.String(r.Комментарий) if r.Комментарий else ""
        r.Аналитика1 = val.strip()
        pl_count += 1
        if len(samples_pl) < 5 and r.Аналитика1:
            samples_pl.append(r.Аналитика1)
    if not r.Аналитика1:
        empty_count += 1

print(f"ЕРП/Каса-строки (→ Аналитика):  {erp_count}")
print(f"PL-строки (→ Комментарий):     {pl_count}")
print(f"С пустым Аналитика1:            {empty_count}")
print()
print("Пример ЕРП/Каса Аналитика1:")
for s in samples_erp:
    print(f"   {s}")
print("Пример PL Аналитика1:")
for s in samples_pl:
    print(f"   {s}")
