# -*- coding: utf-8 -*-
"""Smoke (dry-run): headless загрузка А_НачальнаяЗадолженностьПоЗарплатеСозданнаяПоВыплатам.epf,
шаги 1-2 (ЗаполнитьВедомости + СформироватьПревью) БЕЗ ЗаполнитьДокумент.
Эталон pretest: 7 ведомостей, 192 строки превью, Σ к заполнению = 2 931 876,16."""
import os
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "Обработки", "А_НачальнаяЗадолженностьПоЗарплатеСозданнаяПоВыплатам.epf"))
EXP_VED = 7
EXP_ROWS = 192
EXP_SUM = 2931876.16

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

obr = erp.ВнешниеОбработки.Создать(EPF, False)

# Середина месяца устойчива к TZ-сдвигу datetime->COM; обработка сама берёт НачалоМесяца/КонецМесяца
import datetime
obr.МесяцОстатков = datetime.datetime(2025, 11, 15, 12, 0, 0)
obr.МесяцВедомостей = datetime.datetime(2025, 12, 15, 12, 0, 0)

obr.ЗаполнитьВедомости()
n_ved = obr.Ведомости.Количество()
print(f"Ведомостей: {n_ved} (ожид. {EXP_VED})")

obr.СформироватьПревью()
n_rows = obr.Превью.Количество()
s_fill = 0.0
s_ved = 0.0
statuses = {}
for i in range(n_rows):
    r = obr.Превью.Получить(i)
    s_fill += float(r.СуммаКЗаполнению or 0)
    s_ved += float(r.СуммаПоВедомостям or 0)
    st = str(r.Статус)
    statuses[st] = statuses.get(st, 0) + 1
print(f"Строк превью: {n_rows} (ожид. {EXP_ROWS})")
print(f"Σ по ведомостям = {s_ved:,.2f}; Σ к заполнению = {s_fill:,.2f} (ожид. {EXP_SUM:,.2f})")
print(f"Статусы: {statuses}")

ok = (n_ved == EXP_VED and n_rows == EXP_ROWS and abs(s_fill - EXP_SUM) <= 0.01)
print("SMOKE: " + ("PASS" if ok else "FAIL"))
sys.exit(0 if ok else 1)
