# -*- coding: utf-8 -*-
# Follow-up: субсчета 661/663 + сходимость в разрезе ТОЛЬКО ФЛ (атрибуция vs реальный разрыв)
import win32com.client as wc
import sys
from datetime import datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def fail(e):
    if hasattr(e, 'excepinfo') and e.excepinfo:
        return e.excepinfo[2]
    return str(e)

v8 = wc.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
nach = datetime(2025, 1, 1)
kon  = datetime(2026, 12, 31, 23, 59, 59)

# --- T2'. Список счетов 661*/663* (есть ли субсчета) ---
print("=== T2'. Счета 661*/663* в BuhBud ===")
qs = buh.NewObject("Запрос")
qs.Текст = ('ВЫБРАТЬ Сч.Код КАК Код, Сч.Наименование КАК Имя, Сч.Предопределенный КАК Предоп '
            'ИЗ ПланСчетов.Хозрасчетный КАК Сч '
            'ГДЕ Сч.Код ПОДОБНО "661%" ИЛИ Сч.Код ПОДОБНО "663%" ИЛИ Сч.Код = "66" '
            'УПОРЯДОЧИТЬ ПО Сч.Код')
sel = qs.Выполнить().Выбрать()
while sel.Следующий():
    print(f"   {str(sel.Код).strip():8} {str(sel.Имя)}")

# --- ЕДРПОУ оргов в балансе ---
qo = erp.NewObject("Запрос")
qo.Текст = ('ВЫБРАТЬ Орг.КодПоЕДРПОУ КАК Е ИЗ Справочник.Организации КАК Орг '
            'ГДЕ Орг.А_ВБалансе = ИСТИНА И Орг.КодПоЕДРПОУ <> ""')
edrpou = []
s = qo.Выполнить().Выбрать()
while s.Следующий():
    edrpou.append(str(s.Е).strip())
mas = buh.NewObject("Массив")
for e in edrpou:
    mas.Добавить(e)

# --- ЕРП per ФЛ ---
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Н", nach); q.УстановитьПараметр("К", kon)
q.Текст = """
ВЫБРАТЬ Ост.ФизическоеЛицо.А_ИдКод КАК ИдКод,
    СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КонОст
ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(&Н, &К, , ,
        ФормаPL = ЗНАЧЕНИЕ(Перечисление.А_ФормыPL.Форма1)) КАК Ост
ГДЕ Ост.ФизическоеЛицо.А_ИдКод <> ""
СГРУППИРОВАТЬ ПО Ост.ФизическоеЛицо.А_ИдКод
"""
erp_fl = {}
s = q.Выполнить().Выбрать()
while s.Следующий():
    erp_fl[str(s.ИдКод).strip()] = erp_fl.get(str(s.ИдКод).strip(), 0.0) + float(s.КонОст)

# --- BuhBud per ФЛ ---
qb = buh.NewObject("Запрос")
qb.УстановитьПараметр("Н", nach); qb.УстановитьПараметр("К", kon); qb.УстановитьПараметр("мас", mas)
qb.Текст = """
ВЫБРАТЬ ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод КАК ИдКод,
    СУММА(Ост.СуммаКонечныйОстатокКт) - СУММА(Ост.СуммаКонечныйОстатокДт) КАК Сальдо
ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(&Н, &К) КАК Ост
ГДЕ Ост.Счет.Код В ("661","663") И Ост.Субконто1 ССЫЛКА Справочник.ФизическиеЛица
    И Ост.Организация.КодПоЕДРПОУ В (&мас)
СГРУППИРОВАТЬ ПО ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод
"""
buh_fl = {}
s = qb.Выполнить().Выбрать()
while s.Следующий():
    buh_fl[str(s.ИдКод).strip()] = buh_fl.get(str(s.ИдКод).strip(), 0.0) + float(s.Сальдо)

keys = set(erp_fl) | set(buh_fl)
te = sum(erp_fl.values()); tb = sum(buh_fl.values())
absdiff = sum(abs(erp_fl.get(k,0.0) - buh_fl.get(k,0.0)) for k in keys)
nbig = sum(1 for k in keys if abs(erp_fl.get(k,0.0) - buh_fl.get(k,0.0)) > 0.01)
only_erp = [k for k in keys if k in erp_fl and k not in buh_fl and abs(erp_fl[k])>0.01]
only_buh = [k for k in keys if k in buh_fl and k not in erp_fl and abs(buh_fl[k])>0.01]
print("\n=== Сходимость в разрезе ТОЛЬКО ФЛ (атрибуция по юрлицам убрана) ===")
print(f"Σ ЕРП={te:.2f}  Σ BuhBud={tb:.2f}  Σ|разн|={absdiff:.2f}")
print(f"ФЛ всего={len(keys)}, с |разн|>0.01: {nbig}, только ЕРП={len(only_erp)}, только BuhBud={len(only_buh)}")
big = sorted(((k, erp_fl.get(k,0.0), buh_fl.get(k,0.0)) for k in keys),
            key=lambda t: -abs(t[1]-t[2]))
print("Топ-10 реальных расхождений per ФЛ:")
for k,e,b in big[:10]:
    print(f"   {k}: ЕРП={e:.2f} BuhBud={b:.2f} разн={e-b:.2f}")
print("\nDONE")
