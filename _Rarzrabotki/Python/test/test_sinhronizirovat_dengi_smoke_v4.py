# -*- coding: utf-8 -*-
"""
Smoke: депозитный счёт ТАСдепозит (v6 API — анализ напрямую по счёту).
Проверяет: анализ доходит до конца, пер-счётные потоки (не «Бух=0»),
корректность признака ЕстьРасхождение. Только чтение.
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
mas = erp.NewObject("Массив")
mas.Добавить(r.Получить(0).Ссылка)
rez = obr.АнализироватьДокументыПоСчетам(
    mas, datetime(2026, 1, 1, 0, 0, 0), datetime(2026, 6, 10, 23, 59, 59))
n_dok = obr.ТаблицаДокументов.Количество()
print(f"Фаза 2: {erp.String(rez)} (рядків={n_dok})")
assert n_dok > 0

FAILS = 0
statusy = Counter()
buh0 = 0
rash_err = 0
for i in range(n_dok):
    s = obr.ТаблицаДокументов.Получить(i)
    status = str(erp.String(s.Статус))
    key = status.split("(")[0].strip()
    statusy[key] += 1
    spok = status in ("ОК", "Переказ: синхронно")
    if bool(s.ЕстьРасхождение) == spok:
        rash_err += 1
    if status.startswith("Переказ: розбіжність суми") and abs(float(s.СуммаБух)) < 0.005 \
            and str(erp.String(s.ДокументБух)).strip():
        buh0 += 1

print("Розподіл статусів:")
for k, v in sorted(statusy.items(), key=lambda x: -x[1]):
    print(f"  {v:5d}  {k}")

if statusy.get("Переказ: синхронно", 0) == 0:
    print("FAIL: жодного «Переказ: синхронно»")
    FAILS += 1
if buh0 > 0:
    print(f"FAIL: {buh0} переказів з документом Бух і потоком 0")
    FAILS += 1
if rash_err > 0:
    print(f"FAIL: {rash_err} рядків з неверным ЕстьРасхождение")
    FAILS += 1

print("РЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
