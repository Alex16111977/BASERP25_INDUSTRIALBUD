# -*- coding: utf-8 -*-
"""Счета Казны (IRC 2026 / ООН Админ) без приложенного скана: заказ ЕРП, оплачено и поступило на дату. Только чтение.

    python scheta_bez_faila.py [ГГГГ-ММ-ДД]   # по умолчанию 2026-10-08; xlsx — рядом с scheta.json
"""
import datetime as dt
import json
import sys

import openpyxl
import win32com.client

from sopostavit import Sopostavitel

sys.stdout.reconfigure(encoding='utf-8')
DIR = r"C:\Razrabotki\tovary_schetov_kazna"
d = dt.datetime.strptime(sys.argv[1] if len(sys.argv) > 1 else "2026-10-08", "%Y-%m-%d") + dt.timedelta(hours=23, minutes=59, seconds=59)

erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
sch = [s for s in json.load(open(DIR + r"\scheta.json", encoding="utf-8")) if not s["files"]]
sp = Sopostavitel(erp)

z = erp.NewObject("Запрос")
z.Text = """ВЫБРАТЬ Р.ОбъектРасчетов.Объект КАК Заказ, СУММА(Р.Сумма) КАК Сумма
ИЗ (ВЫБРАТЬ Р.ОбъектРасчетов КАК ОбъектРасчетов, Р.Сумма КАК Сумма
	ИЗ Документ.СписаниеБезналичныхДенежныхСредств.РасшифровкаПлатежа КАК Р
	ГДЕ Р.Ссылка.Проведен И Р.Ссылка.Дата <= &М И ТИПЗНАЧЕНИЯ(Р.ОбъектРасчетов.Объект) = ТИП(Документ.ЗаказПоставщику)
	ОБЪЕДИНИТЬ ВСЕ
	ВЫБРАТЬ Р.ОбъектРасчетов, Р.Сумма
	ИЗ Документ.РасходныйКассовыйОрдер.РасшифровкаПлатежа КАК Р
	ГДЕ Р.Ссылка.Проведен И Р.Ссылка.Дата <= &М И ТИПЗНАЧЕНИЯ(Р.ОбъектРасчетов.Объект) = ТИП(Документ.ЗаказПоставщику)) КАК Р
СГРУППИРОВАТЬ ПО Р.ОбъектРасчетов.Объект
;
ВЫБРАТЬ Т.Ссылка.ЗаказПоставщику КАК Заказ, СУММА(Т.СуммаСНДС) КАК Сумма
ИЗ Документ.ПриобретениеТоваровУслуг.Товары КАК Т
ГДЕ Т.Ссылка.Проведен И Т.Ссылка.Дата <= &М
СГРУППИРОВАТЬ ПО Т.Ссылка.ЗаказПоставщику"""
z.SetParameter("М", d)
res = z.ExecuteBatch()


def to_map(r):
    t = r.Выгрузить()
    m = {}
    for i in range(t.Количество()):
        x = t.Получить(i)
        if S(x.Получить(0)):
            m[S(x.Получить(0).УникальныйИдентификатор())] = float(x.Получить(1) or 0)
    return m


opl, pol = to_map(res.Получить(0)), to_map(res.Получить(1))
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Без скану"
ws.append(["Підрозділ", "Рахунок №", "Дата рахунку", "Контрагент", "ЄДРПОУ", "Сума рахунку", "Замовлення ЕРП",
           f"Оплачено на {d:%d.%m.%Y}", "Надійшло (ПТУ)", "Примітка"])
tot = {}
for s in sorted(sch, key=lambda s: (s["podr"], s["kontragent"], s["data"])):
    ref, uid, warn = sp.zakaz(s)
    o, p = opl.get(uid, 0.0), pol.get(uid, 0.0)
    proj = "МД IRC 2026" if "IRC" in s["podr"] else "МД ООН Админ"
    t = tot.setdefault(proj, [0, 0.0, 0.0, 0, 0.0])
    t[0] += 1
    t[1] += s["summa_scheta"] or s["summa_doc"]
    t[2] += o
    if o:
        t[3] += 1
        t[4] += o
    ws.append([s["podr"], s["nomer"], s["data_vh"], s["kontragent"], s["edrpou"], s["summa_scheta"] or s["summa_doc"],
               S(ref) if ref else "", round(o, 2), round(p, 2), warn])
for c, w in zip("ABCDEFGHIJ", (34, 16, 12, 40, 11, 14, 44, 16, 16, 34)):
    ws.column_dimensions[c].width = w
for row in ws.iter_rows(min_row=2, min_col=6, max_col=9):
    for c in row:
        c.number_format = "# ##0.00"
out = DIR + rf"\Рахунки без скану на {d:%d.%m.%Y}.xlsx"
wb.save(out)
for k, (n, sm, o, no, so) in tot.items():
    print(f"{k}: счетов {n} на {sm:,.2f}; оплачено {o:,.2f} (по {no} счетам)".replace(",", " "))
print(out)
