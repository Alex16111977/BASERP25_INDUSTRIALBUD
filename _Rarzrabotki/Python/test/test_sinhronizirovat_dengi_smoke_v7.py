# -*- coding: utf-8 -*-
"""
Smoke v7: новый API объекта (только чтение / безопасные вызовы).
1. СравнитьОстатки + АнализироватьДокументы() без аргументов — как раньше (совместимость).
2. АнализироватьДокументы(0) — точечный анализ первой строки расхождений.
3. АнализироватьДокументы(9999) — протухший индекс -> «Невірний номер рядка».
4. ВыполнитьСинхронизациюСтроки(-1) — «Невірний номер рядка» (без мутаций).
"""
import sys
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

FAILS = 0

obr = erp.ВнешниеОбработки.Создать(EPF, False)
obr.НачалоПериода = datetime(2026, 6, 1, 0, 0, 0)
obr.ОкончаниеПериода = datetime(2026, 6, 7, 0, 0, 0)

rez1 = obr.СравнитьОстатки()
n_rash = obr.ТаблицаРасхождений.Количество()
print(f"Фаза 1: {erp.String(rez1)} (рядків={n_rash})")
assert n_rash > 0

# 2. Точечный анализ строки 0
rez_point = obr.АнализироватьДокументы(0)
n_dok_point = obr.ТаблицаДокументов.Количество()
print(f"Точковий аналіз [0]: {erp.String(rez_point)} (рядків={n_dok_point})")
if n_dok_point == 0:
    print("ПРИМІТКА: документів немає (можливо рахунок без рухів за тиждень) — не FAIL")

# 3. Протухший индекс
rez_bad = str(erp.String(obr.АнализироватьДокументы(9999)))
print(f"Протухлий індекс: '{rez_bad}'")
if "Невірний номер рядка" not in rez_bad:
    print("FAIL: очікував «Невірний номер рядка»")
    FAILS += 1

# 4. Построчная синхронизация с невалидным индексом — безопасно
rez_sync_bad = str(erp.String(obr.ВыполнитьСинхронизациюСтроки(-1)))
print(f"Синхронізація [-1]: '{rez_sync_bad}'")
if "Невірний номер рядка" not in rez_sync_bad:
    print("FAIL: очікував «Невірний номер рядка»")
    FAILS += 1

# 1 (контроль). Совместимость: вызов без аргументов
for i in range(n_rash):
    obr.ТаблицаРасхождений.Получить(i).Синхронизировать = (i == 0)
rez_compat = obr.АнализироватьДокументы()
print(f"Виклик без аргументів: {erp.String(rez_compat)}")

print("РЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
