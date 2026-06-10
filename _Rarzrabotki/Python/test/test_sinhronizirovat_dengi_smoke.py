# -*- coding: utf-8 -*-
"""
Smoke-тест обработки СинхронизироватьДеньги (.epf) — Фаза 1 + Фаза 2 (только чтение).
Фаза 3 (синхронизация) НЕ выполняется — боевые действия за пользователем.
"""
import sys
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

obr = erp.ВнешниеОбработки.Создать(EPF, False)
obr.НачалоПериода = datetime(2026, 6, 1, 0, 0, 0)
obr.ОкончаниеПериода = datetime(2026, 6, 7, 0, 0, 0)

# Фаза 1
rez1 = obr.СравнитьОстатки()
n_rash = obr.ТаблицаРасхождений.Количество()
print(f"Фаза 1: {erp.String(rez1)}")
print(f"ТаблицаРасхождений: {n_rash} рядків")
for i in range(min(n_rash, 5)):
    s = obr.ТаблицаРасхождений.Получить(i)
    print(f"  [{i}] {erp.String(s.БанковскийСчет)}: ЕРП={s.ОстатокЕРП} Бух={s.ОстатокБух} Δ={s.Разница}")

# Фаза 2 (чтение): только первая строка расхождений, чтобы smoke был быстрым
if n_rash > 0:
    for i in range(n_rash):
        obr.ТаблицаРасхождений.Получить(i).Синхронизировать = (i == 0)
    rez2 = obr.АнализироватьДокументы()
    n_dok = obr.ТаблицаДокументов.Количество()
    print(f"Фаза 2: {erp.String(rez2)}")
    for i in range(min(n_dok, 10)):
        s = obr.ТаблицаДокументов.Получить(i)
        print(f"  [{i}] ЕРП='{erp.String(s.ДокументЕРП)}' Бух='{erp.String(s.ДокументБух)}'"
              f" СумЕРП={s.СуммаЕРП} СумБух={s.СуммаБух} Дія='{erp.String(s.Действие)}' Статус='{erp.String(s.Статус)}'")
else:
    print("Фаза 2: пропущено (немає розбіжностей)")

print("SMOKE OK")
