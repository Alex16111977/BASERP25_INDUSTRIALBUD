# -*- coding: utf-8 -*-
"""Кросс-месячная диагностика незакрытия по форме (статья по argv или 'Зарплата ИТР').
Для каждого месяца: Приход по форме vs записанный Расход распределения vs что предв.режим
ХОЧЕТ распределить. Если 'предв.хочет' == Приход, а 'записано' < Приход → лечится пере-закрытием.
Без записи.
Запуск: python diag_zarplata_itr_predv.py ["Статья"]
"""
import sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client, pythoncom
from win32com.client import VARIANT

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
STAT = sys.argv[1] if len(sys.argv) > 1 else "Зарплата ИТР"
MONTHS = [(2025, 12), (2026, 1), (2026, 3), (2026, 5)]


def forma(f):
    if not erp.ЗначениеЗаполнено(f):
        return "(пусто)"
    x = erp.XMLСтрока(f)
    return "Ф1" if "Форма1" in x else ("Ф2" if "Форма2" in x else x)


def sums_by_form(extra_where, ref_date):
    # границы месяца СЕРВЕРНО (memory feedback_balans_etalon_period_serverside)
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Стат", STAT)
    q.УстановитьПараметр("Д", ref_date)
    q.Текст = ("ВЫБРАТЬ П.ФормаPL КАК Ф, СУММА(П.Сумма) КАК С ИЗ РегистрНакопления.ПрочиеРасходы КАК П "
               "ГДЕ П.СтатьяРасходов.Наименование=&Стат И П.Активность "
               "И П.Период>=НАЧАЛОПЕРИОДА(&Д,МЕСЯЦ) И П.Период<=КОНЕЦПЕРИОДА(&Д,МЕСЯЦ) " + extra_where +
               " СГРУППИРОВАТЬ ПО П.ФормаPL")
    t = q.Выполнить().Выгрузить()
    d = {}
    for i in range(t.Количество()):
        r = t.Получить(i); d[forma(r.Ф)] = d.get(forma(r.Ф), 0.0) + float(r.С or 0)
    return d


print("Статья: '%s'\n" % STAT)
print("%-10s %-6s %14s %14s %14s  %s" % ("Месяц", "Форма", "Приход", "Записано", "Предв.хочет", "Вердикт"))
for y, m in MONTHS:
    ref = datetime.datetime(y, m, 15)  # любая дата внутри месяца; границы берём серверно
    wp = " И П.ВидДвижения=ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)"
    wr = (" И П.ВидДвижения=ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) "
          "И П.Регистратор ССЫЛКА Документ.РаспределениеПрочихЗатрат")
    prihod = sums_by_form(wp, ref)
    rashod = sums_by_form(wr, ref)

    # предв.режим по всем распределениям месяца, двигающим статью
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Стат", STAT); q.УстановитьПараметр("Д", ref)
    q.Текст = ("ВЫБРАТЬ РАЗЛИЧНЫЕ П.Регистратор КАК Док, П.Регистратор.Организация КАК Орг "
               "ИЗ РегистрНакопления.ПрочиеРасходы КАК П ГДЕ П.Регистратор ССЫЛКА Документ.РаспределениеПрочихЗатрат "
               "И П.СтатьяРасходов.Наименование=&Стат И П.Активность "
               "И П.Период>=НАЧАЛОПЕРИОДА(&Д,МЕСЯЦ) И П.Период<=КОНЕЦПЕРИОДА(&Д,МЕСЯЦ)")
    docs = q.Выполнить().Выгрузить()
    predv = {}
    for i in range(docs.Количество()):
        row = docs.Получить(i)
        орги = erp.NewObject("Array"); орги.Добавить(row.Орг)
        mvt = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
        try:
            erp.Документы.РаспределениеПрочихЗатрат.РаспределитьРасходыНаФинансовыйРезультат(
                datetime.datetime(y, m, 28), орги, mvt, row.Док)
            qr = erp.NewObject("Запрос"); qr.МенеджерВременныхТаблиц = mvt.value
            qr.УстановитьПараметр("Стат", STAT)
            qr.Текст = ("ВЫБРАТЬ Т.ФормаPL КАК Ф, СУММА(Т.Сумма) КАК С ИЗ РезультатРаспределения КАК Т "
                        "ГДЕ Т.СтатьяРасходов.Наименование=&Стат СГРУППИРОВАТЬ ПО Т.ФормаPL")
            tr = qr.Выполнить().Выгрузить()
            for j in range(tr.Количество()):
                r = tr.Получить(j); predv[forma(r.Ф)] = predv.get(forma(r.Ф), 0.0) + float(r.С or 0)
        except Exception as e:
            ei = getattr(e, 'excepinfo', None)
            print("   предв.ОШИБКА %d-%02d: %s" % (y, m, ei[2] if ei else e))
    for f in ("Ф1", "Ф2"):
        p, rec, pv = prihod.get(f, 0.0), rashod.get(f, 0.0), predv.get(f, 0.0)
        if abs(p) < 0.01 and abs(rec) < 0.01 and abs(pv) < 0.01:
            continue
        verd = "OK (закрыто)" if abs(p - rec) < 0.01 else (
            "ЛЕЧИТСЯ пере-закрытием" if abs(pv - p) < 0.01 else "??? предв≠приход")
        print("%-10s %-6s %14.2f %14.2f %14.2f  %s" % ("%d-%02d" % (y, m), f, p, rec, pv, verd))
