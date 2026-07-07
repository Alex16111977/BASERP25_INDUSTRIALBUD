# -*- coding: utf-8 -*-
# Smoke: новые разрезы Договор + СтатьяДвижения в отчёте А_РасшифровкаЗадолженностиПодразделений.
import sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_РасшифровкаЗадолженностиПодразделений.erf"
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

def err(e):
    if hasattr(e, 'excepinfo') and e.excepinfo:
        return e.excepinfo[2]
    return str(e)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)
print("Connected OK")

report = erp.ВнешниеОтчеты.Создать(ERF, False)
print("Report loaded OK")

params = erp.NewObject("Структура")
params.Insert("НачалоПериода",         datetime.datetime(2025, 12, 1, 12, 0, 0))
params.Insert("ОкончаниеПериода",      datetime.datetime(2025, 12, 31, 12, 0, 0))
params.Insert("Организация",           erp.Справочники.Организации.ПустаяСсылка())
params.Insert("ТолькоРасхождения",     False)
params.Insert("ВключатьНепроведенные", False)

try:
    tz = report.ПолучитьТаблицуРасшифровки(params)
except Exception as e:
    print("FAIL ПолучитьТаблицуРасшифровки:", err(e))
    sys.exit(1)

n = tz.Count()
print("Rows TЗ:", n)
if n == 0:
    print("FAIL: пустая ТЗ")
    sys.exit(1)

dogovor_filled = 0
prihod = 0.0
rashod = 0.0
suti = {}
for i in range(n):
    r = tz.Get(i)
    if erp.ЗначениеЗаполнено(r.Договор):
        dogovor_filled += 1
    s = r.СтатьяДвижения
    suti[s] = suti.get(s, 0) + 1
    prihod += float(r.Приход)
    rashod += float(r.Расход)

print("Договор заполнен в строках: %d / %d" % (dogovor_filled, n))
print("Сумма Приход = %.2f ; Расход = %.2f" % (prihod, rashod))
print("Различных 'статей движения / сути': %d" % len(suti))
print("--- ТОП-15 сутей (по числу строк) ---")
for s, c in sorted(suti.items(), key=lambda kv: -kv[1])[:15]:
    label = s if s else "<ПУСТО>"
    print("  %5d  %s" % (c, label))

empty_suti = suti.get("", 0)
print("Строк с ПУСТОЙ сутью: %d (ожидается 0 — fallback на синоним)" % empty_suti)
print("SMOKE DONE")
