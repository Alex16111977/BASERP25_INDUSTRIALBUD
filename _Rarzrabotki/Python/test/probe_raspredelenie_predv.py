# -*- coding: utf-8 -*-
"""
Rule #-1 БЕЗ записи в базу: вызвать РаспределитьРасходыНаФинансовыйРезультат в
ПРЕДВАРИТЕЛЬНОМ режиме (ДокументРаспределения=279) — он исполняет пакет в МВТ и
возвращается БЕЗ записи движений (стр. 1116-1121). Затем читаем РезультатРаспределения
из МВТ и проверяем разрез ФормаPL (ожид. Ф1=342280.10 / Ф2=1032673.36).
"""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def forma(f):
    if not erp.ЗначениеЗаполнено(f):
        return "(пусто)"
    x = erp.XMLСтрока(f)
    return "Форма1" if "Форма1" in x else ("Форма2" if "Форма2" in x else x)


# ref 279
q = erp.NewObject("Запрос")
q.Текст = ('ВЫБРАТЬ Р.Ссылка КАК С, Р.Организация КАК Орг ИЗ Документ.РаспределениеПрочихЗатрат КАК Р '
           'ГДЕ Р.Номер="00000000279" И Р.Дата=ДАТАВРЕМЯ(2025,12,31,23,59,59)')
s = q.Выполнить().Выбрать(); s.Следующий()
doc = s.С
org = s.Орг
print("doc279 =", erp.XMLСтрока(doc), "| орг =", org)

период = erp.NewObject("Date", "20251231000000") if False else None
# period as datetime
import datetime
период = datetime.datetime(2025, 12, 31)

орги = erp.NewObject("Array")
орги.Добавить(org)

import pythoncom
from win32com.client import VARIANT
# МенеджерВременныхТаблиц — [out]-параметр (присваивается на стр.1119); ловим через VARIANT byref
mvt_var = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
print("Вызов РаспределитьРасходыНаФинансовыйРезультат (предв. режим)...")
res = erp.Документы.РаспределениеПрочихЗатрат.РаспределитьРасходыНаФинансовыйРезультат(
    период, орги, mvt_var, doc)
print("Возврат:", res, "| byref value type:", type(mvt_var.value))
мвт = mvt_var.value


def read_result(tm, label):
    q2 = erp.NewObject("Запрос")
    q2.МенеджерВременныхТаблиц = tm
    q2.Текст = ("ВЫБРАТЬ Т.ФормаPL КАК ФормаPL, КОЛИЧЕСТВО(*) КАК К, СУММА(Т.Сумма) КАК Сумма "
                "ИЗ РезультатРаспределения КАК Т "
                "СГРУППИРОВАТЬ ПО Т.ФормаPL")
    try:
        tbl = q2.Выполнить().Выгрузить()
    except Exception as e:
        ei = getattr(e, 'excepinfo', None)
        print("  [%s] ошибка чтения РезультатРаспределения: %s" % (label, ei[2] if ei else e))
        return
    print("  [%s] строк РезультатРаспределения=%d:" % (label, tbl.Количество()))
    for i in range(tbl.Количество()):
        r = tbl.Получить(i)
        print("     %-8s К=%-3d Сумма=%15.2f" % (forma(r.ФормаPL), r.К, float(r.Сумма)))


def dump(tm, table, sumcol="Сумма"):
    q2 = erp.NewObject("Запрос")
    q2.МенеджерВременныхТаблиц = tm
    q2.Текст = ("ВЫБРАТЬ Т.ФормаPL КАК ФормаPL, КОЛИЧЕСТВО(*) КАК К, СУММА(Т.%s) КАК С "
                "ИЗ %s КАК Т СГРУППИРОВАТЬ ПО Т.ФормаPL" % (sumcol, table))
    has_form = True
    try:
        t = q2.Выполнить().Выгрузить()
    except Exception:
        has_form = False
        q2.Текст = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, СУММА(Т.%s) КАК С ИЗ %s КАК Т" % (sumcol, table)
        try:
            t = q2.Выполнить().Выгрузить()
        except Exception as e:
            ei = getattr(e, 'excepinfo', None)
            print("  [%s] нет таблицы/ошибка: %s" % (table, ei[2] if ei else e))
            return
    print("  %s:" % table)
    tot = 0.0
    for i in range(t.Количество()):
        r = t.Получить(i)
        val = float(r.С) if r.С is not None else 0.0
        tot += val
        if has_form:
            print("     %-8s К=%-4d Σ=%15.2f" % (forma(r.ФормаPL), r.К, val))
        else:
            print("     [без формы] К=%-4d Σ=%15.2f" % (r.К, val))
    print("     ИТОГО Σ=%.2f" % tot)


for tbl in ["ПредварительныеОстаткиРасходов", "ОстаткиРасходов", "РасходыКРаспределению", "РезультатРаспределения"]:
    dump(мвт, tbl)
