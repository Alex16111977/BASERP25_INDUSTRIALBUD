# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
nach = datetime.datetime(2025,12,1,0,0,0); kon = datetime.datetime(2025,12,31,23,59,59)

def typ_po_schetu(code):
    code = code.strip()
    e = erp.Перечисления.ТипыНалогов
    if code=="6411": return e.НДФЛ
    if code=="642":  return e.ВоенныйСбор
    if code=="651":  return e.НачисленныйЕСВ
    if code=="6412": return e.НДС
    if code=="6413": return e.НалогНаПрибыль
    if code=="6414": return e.ЕдиныйНалог
    return e.ДругиеНалоги

# ERP side keyed
names = ["НДФЛ","НДФЛДоходыКонтрагентов","НФДЛДивиденды","НФДЛДивидендыСотрудникам",
         "НДФЛДоначисленныйПоРезультатамПроверки","НДФЛПередачаЗадолженностиВНалоговыйОрган",
         "НДФЛПрочиеРасчетыСПерсоналом"]
arr = erp.NewObject("Массив")
for nm in names: arr.Добавить(getattr(erp.Перечисления.ТипыНалогов, nm))
qe = erp.NewObject("Запрос")
qe.УстановитьПараметр("НачалоПериода", nach); qe.УстановитьПараметр("КонецПериода", kon); qe.УстановитьПараметр("СписокНДФЛ", arr)
qe.Текст = open(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\_review_nalogi_erp.py",encoding="utf-8").read().split('q.Текст = """')[1].split('"""')[0]
te = qe.Выполнить().Выгрузить()
erp_keys = {}
for i in range(te.Количество()):
    r = te.Получить(i)
    k = str(r.ЕДРПОУ).strip() + "|" + erp.XMLСтрока(r.ВидНалога)
    erp_keys[k] = float(r.КонечныйОстаток)

# BuhBud side -> map to tax, aggregate, key
qo = erp.NewObject("Запрос")
qo.Текст = 'ВЫБРАТЬ Организации.КодПоЕДРПОУ КАК КодПоЕДРПОУ ИЗ Справочник.Организации КАК Организации ГДЕ Организации.А_ВБалансе И Организации.КодПоЕДРПОУ <> ""'
masEDR = buh.NewObject("Массив")
so = qo.Выполнить().Выбрать()
while so.Следующий(): masEDR.Добавить(so.КодПоЕДРПОУ)
masSch = buh.NewObject("Массив")
for c in "6411,6412,6413,6414,6415,6417,642,651".split(","): masSch.Добавить(c.strip())
qb = buh.NewObject("Запрос")
qb.УстановитьПараметр("НачалоПериода", nach); qb.УстановитьПараметр("КонецПериода", kon)
qb.УстановитьПараметр("масЕДРПОУ", masEDR); qb.УстановитьПараметр("Счета", masSch)
qb.Текст = """ВЫБРАТЬ Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ, Ост.Счет.Код КАК КодСчета,
СУММА(Ост.СуммаКонечныйОстатокКт - Ост.СуммаКонечныйОстатокДт) КАК КонечныйОстаток
ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(&НачалоПериода, &КонецПериода, , , Счет.Код В (&Счета), , ) КАК Ост
ГДЕ Ост.Организация.КодПоЕДРПОУ В (&масЕДРПОУ)
СГРУППИРОВАТЬ ПО Ост.Организация.КодПоЕДРПОУ, Ост.Счет.Код"""
tb = qb.Выполнить().Выгрузить()
s = buh.ЗначениеВСтрокуВнутр(tb); tb2 = erp.ЗначениеИзСтрокиВнутр(s)
buh_keys = {}
for i in range(tb2.Количество()):
    r = tb2.Получить(i)
    vn = typ_po_schetu(str(r.КодСчета))
    k = str(r.ЕДРПОУ).strip() + "|" + erp.XMLСтрока(vn)
    buh_keys[k] = buh_keys.get(k,0)+float(r.КонечныйОстаток)

allk = set(erp_keys)|set(buh_keys)
print("ERP keys:", len(erp_keys), "Buh keys:", len(buh_keys), "merged:", len(allk))
both=only_e=only_b=0
for k in allk:
    if k in erp_keys and k in buh_keys: both+=1
    elif k in erp_keys: only_e+=1
    else: only_b+=1
print(f"both={both} only_ERP={only_e} only_Buh={only_b}")
