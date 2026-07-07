# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

nach = datetime.datetime(2025,12,1,0,0,0)
kon  = datetime.datetime(2025,12,31,23,59,59)

# ЕДРПОУ list from ERP (А_ВБалансе)
qo = erp.NewObject("Запрос")
qo.Текст = 'ВЫБРАТЬ Организации.КодПоЕДРПОУ КАК КодПоЕДРПОУ ИЗ Справочник.Организации КАК Организации ГДЕ Организации.А_ВБалансе И Организации.КодПоЕДРПОУ <> ""'
masEDR = buh.NewObject("Массив")
so = qo.Выполнить().Выбрать()
while so.Следующий():
    masEDR.Добавить(so.КодПоЕДРПОУ)
print("masEDRPOU:", masEDR.Количество())

masSch = buh.NewObject("Массив")
for c in "6411,6412,6413,6414,6415,6417,642,651".split(","):
    masSch.Добавить(c.strip())

qb = buh.NewObject("Запрос")
qb.УстановитьПараметр("НачалоПериода", nach)
qb.УстановитьПараметр("КонецПериода", kon)
qb.УстановитьПараметр("масЕДРПОУ", masEDR)
qb.УстановитьПараметр("Счета", masSch)
qb.Текст = """
ВЫБРАТЬ
	Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
	Ост.Счет.Код КАК КодСчета,
	СУММА(Ост.СуммаНачальныйОстатокКт - Ост.СуммаНачальныйОстатокДт) КАК НачальныйОстаток,
	СУММА(Ост.СуммаОборотКт) КАК Начисления,
	СУММА(Ост.СуммаОборотДт) КАК Выплаты,
	СУММА(Ост.СуммаКонечныйОстатокКт - Ост.СуммаКонечныйОстатокДт) КАК КонечныйОстаток
ИЗ
	РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(&НачалоПериода, &КонецПериода, , , Счет.Код В (&Счета), , ) КАК Ост
ГДЕ
	Ост.Организация.КодПоЕДРПОУ В (&масЕДРПОУ)
СГРУППИРОВАТЬ ПО
	Ост.Организация.КодПоЕДРПОУ,
	Ост.Счет.Код
"""
try:
    tz = qb.Выполнить().Выгрузить()
    print("BuhBud OK rows:", tz.Количество())
    # serialize/deserialize roundtrip (mirror BSL)
    s = buh.ЗначениеВСтрокуВнутр(tz)
    tz2 = erp.ЗначениеИзСтрокиВнутр(s)
    print("roundtrip rows:", tz2.Количество())
    # 6412 total for 40645273
    tot = {}
    for i in range(tz2.Количество()):
        r = tz2.Получить(i)
        if str(r.ЕДРПОУ).strip() == "40645273":
            tot[str(r.КодСчета).strip()] = tot.get(str(r.КодСчета).strip(),0)+float(r.КонечныйОстаток)
    print("40645273 КонОст by счет:")
    for k,v in sorted(tot.items()):
        print(f"   {k}: {v:.2f}")
except Exception as e:
    info = getattr(e,'excepinfo',None)
    print("BuhBud FAIL:", info[2] if info else e)
