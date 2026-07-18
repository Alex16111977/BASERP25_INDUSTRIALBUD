# -*- coding: utf-8 -*-
"""
Smoke-тест обработки СинхронизироватьДеньги (.epf) — Фаза 1 + Фаза 2 (только чтение).
v2: + проверка сортировки по |Різниці| убыв и фильтра по банковскому счёту.
Фаза 3 (синхронизация) НЕ выполняется — боевые действия за пользователем.
"""
import sys
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

obr = erp.ВнешниеОбработки.Создать(EPF, False)
obr.НачалоПериода = datetime(2026, 6, 1, 0, 0, 0)
obr.ОкончаниеПериода = datetime(2026, 6, 7, 0, 0, 0)

# --- Фаза 1 без фильтра ---
rez1 = obr.СравнитьОстатки()
n_rash = obr.ТаблицаРасхождений.Количество()
print(f"Фаза 1: {erp.String(rez1)}")
prev_abs = None
sort_ok = True
for i in range(n_rash):
    s = obr.ТаблицаРасхождений.Получить(i)
    cur_abs = abs(float(s.Разница))
    if prev_abs is not None and cur_abs > prev_abs + 0.005:
        sort_ok = False
    prev_abs = cur_abs
    if i < 5:
        print(f"  [{i}] {erp.String(s.БанковскийСчет)}: ЕРП={s.ОстатокЕРП} Бух={s.ОстатокБух} Δ={s.Разница}")
print(f"Сортировка |Різниця| убыв: {'OK' if sort_ok else 'FAIL'}")
assert sort_ok

# --- Фаза 1 с фильтром по счёту первой строки ---
first_uid = str(erp.String(obr.ТаблицаРасхождений.Получить(0).БанковскийСчетUID)).strip()
acc_ref = erp.Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(
    erp.NewObject("УникальныйИдентификатор", first_uid))
obr.ФильтрБанковскийСчет = acc_ref
rez1f = obr.СравнитьОстатки()
n_filt = obr.ТаблицаРасхождений.Количество()
print(f"Фаза 1 (фільтр {erp.String(acc_ref)}): рядків={n_filt}")
assert n_filt == 1, f"с фильтром ожидалась 1 строка, получено {n_filt}"

# --- Фаза 2 по отфильтрованной строке (чтение) ---
obr.ТаблицаРасхождений.Получить(0).Синхронизировать = True
rez2 = obr.АнализироватьДокументы()
n_dok = obr.ТаблицаДокументов.Количество()
print(f"Фаза 2: {erp.String(rez2)}")
for i in range(min(n_dok, 8)):
    s = obr.ТаблицаДокументов.Получить(i)
    print(f"  [{i}] ЕРП='{erp.String(s.ДокументЕРП)}' Бух='{erp.String(s.ДокументБух)}'"
          f" СумЕРП={s.СуммаЕРП} СумБух={s.СуммаБух} Дія='{erp.String(s.Действие)}' Статус='{erp.String(s.Статус)}'")

print("SMOKE OK")
