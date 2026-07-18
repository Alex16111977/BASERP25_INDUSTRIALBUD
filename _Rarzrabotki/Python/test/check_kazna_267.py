# -*- coding: utf-8 -*-
import win32com.client, sys, json

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
S = conn.String

q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 30
    БДДС.Сотрудник.ИНН КАК ИНН,
    БДДС.Сотрудник.Наименование КАК ФИО,
    БДДС.Подразделение.Код КАК ПодрКод,
    СУММА(БДДС.Сумма) КАК Сумма
ИЗ
    РегистрНакопления.БДДС КАК БДДС
ГДЕ
    БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
    И БДДС.Регистратор.Номер = "000000267"
    И БДДС.Период МЕЖДУ &Н И &К
СГРУППИРОВАТЬ ПО
    БДДС.Сотрудник.ИНН,
    БДДС.Сотрудник.Наименование,
    БДДС.Подразделение.Код
УПОРЯДОЧИТЬ ПО ФИО"""

import datetime
q.SetParameter("Н", datetime.datetime(2025, 12, 1))
q.SetParameter("К", datetime.datetime(2025, 12, 31, 23, 59, 59))

res = q.Execute()
sel = res.Choose()

print("БДДС records for РаспределениеЗП 000000267 (Dec 2025):")
print("-" * 80)
cnt = 0
empty_inn = 0
while sel.Next():
    inn = S(sel.ИНН).strip()
    fio = S(sel.ФИО).strip()
    dept = S(sel.ПодрКод).strip()
    summ = sel.Сумма
    if not inn:
        empty_inn += 1
        tag = "!NO_INN!"
    else:
        tag = inn
    print(f"  INN={tag:12s} FIO={fio:35s} Dept={dept:6s} Sum={summ}")
    cnt += 1

print(f"\nTotal: {cnt} rows, empty INN: {empty_inn}")
