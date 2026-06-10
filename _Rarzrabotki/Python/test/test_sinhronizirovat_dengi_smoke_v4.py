# -*- coding: utf-8 -*-
"""
Smoke v4: счёт ТАСдепозит (на нём v3 падал на МетаДанные(671) и давал «Бух = 0»).
Проверяет: анализ доходит до конца, переводы получают пер-счётный поток (не 0),
распределение статусов. Только чтение.
"""
import sys
from collections import Counter
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"
IBAN_DEPOZIT = "UA693395002610201537072000001"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 БСО.Ссылка КАК Ссылка
ИЗ Справочник.БанковскиеСчетаОрганизаций КАК БСО
ГДЕ БСО.НомерСчета = &НомерСчета
"""
q.SetParameter("НомерСчета", IBAN_DEPOZIT)
r = q.Execute().Выгрузить()
assert r.Количество() == 1, "депозитный счёт не найден в ERP"

obr = erp.ВнешниеОбработки.Создать(EPF, False)
obr.ФильтрБанковскийСчет = r.Получить(0).Ссылка
obr.НачалоПериода = datetime(2026, 1, 1, 0, 0, 0)
obr.ОкончаниеПериода = datetime(2026, 6, 10, 0, 0, 0)

obr.СравнитьОстатки()
n_rash = obr.ТаблицаРасхождений.Количество()
print(f"Фаза 1: розбіжностей={n_rash}")
assert n_rash >= 1

for i in range(n_rash):
    obr.ТаблицаРасхождений.Получить(i).Синхронизировать = True
obr.АнализироватьДокументы()
n_dok = obr.ТаблицаДокументов.Количество()
print(f"Фаза 2: документів={n_dok}")

FAILS = 0
statusy = Counter()
buh0 = 0
rash_err = 0
for i in range(n_dok):
    s = obr.ТаблицаДокументов.Получить(i)
    status = str(erp.String(s.Статус))
    key = status.split("(")[0].strip()
    statusy[key] += 1
    # признак расхождения: Ложь только для спокойных статусов
    spok = status in ("ОК", "Переказ: синхронно")
    if bool(s.ЕстьРасхождение) == spok:
        rash_err += 1
        if rash_err <= 3:
            print(f"  FAIL ЕстьРасхождение={s.ЕстьРасхождение} при статусе '{status}'")
    if status.startswith("Переказ: розбіжність суми") and abs(float(s.СуммаБух)) < 0.005 \
            and str(erp.String(s.ДокументБух)).strip():
        buh0 += 1
        if buh0 <= 3:
            print(f"  ПІДОЗРА Бух=0: ЕРП='{erp.String(s.ДокументЕРП)}' Бух='{erp.String(s.ДокументБух)}'"
                  f" СумЕРП={s.СуммаЕРП}")

print("Розподіл статусів:")
for k, v in sorted(statusy.items(), key=lambda x: -x[1]):
    print(f"  {v:5d}  {k}")

if statusy.get("Переказ: синхронно", 0) == 0:
    print("FAIL: жодного «Переказ: синхронно» на депозитному рахунку")
    FAILS += 1
if buh0 > 0:
    print(f"FAIL: {buh0} переказів з документом Бух, але потоком 0 — рахунок BuhBud не зрезолвлено")
    FAILS += 1
if rash_err > 0:
    print(f"FAIL: {rash_err} рядків з неверным ЕстьРасхождение")
    FAILS += 1

print("РЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
