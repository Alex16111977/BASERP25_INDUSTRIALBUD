# -*- coding: utf-8 -*-
# Функциональный smoke Фазы 2: СравнитьОстатки -> АнализироватьДокументы(0) на одном контрагенте.
import win32com.client as w, sys, datetime, pywintypes
sys.stdout.reconfigure(encoding='utf-8')
v8 = w.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
path = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты.epf"
обр = erp.ВнешниеОбработки.Создать(path, False)
обр.НачалоПериода    = pywintypes.Time(datetime.datetime(2026, 1, 1, 12, 0, 0))
обр.ОкончаниеПериода = pywintypes.Time(datetime.datetime(2026, 1, 31, 12, 0, 0))
обр.СравнитьОстатки()
n = обр.ТаблицаКонтрагентов.Количество()
print("Контрагентов:", n)
if n == 0:
    print("[SKIP] нет контрагентов"); sys.exit(0)
# найти контрагента с расхождением В ОБОРОТЕ (Приход/Расход) — тогда есть документы периода
целевойКлюч = None
тр = обр.ТаблицаРасхождений
for i in range(тр.Количество()):
    r = тр.Получить(i)
    if abs(float(r.ПриходЕРП) - float(r.ПриходБух)) > 0.01 or abs(float(r.РасходЕРП) - float(r.РасходБух)) > 0.01:
        целевойКлюч = erp.String(r.КонтрагентКлюч); break
idx = 0
if целевойКлюч is not None:
    for i in range(n):
        if erp.String(обр.ТаблицаКонтрагентов.Получить(i).КонтрагентКлюч) == целевойКлюч:
            idx = i; break
к0 = обр.ТаблицаКонтрагентов.Получить(idx)
print("Анализ контрагента[%d]:" % idx, erp.String(к0.Контрагент), "| договоров:", int(к0.Договоров),
      "| оборотное расхождение:", целевойКлюч is not None)
try:
    res = обр.АнализироватьДокументы(idx)
    print("[OK] АнализироватьДокументы ->", erp.String(res))
    nd = обр.ТаблицаДокументов.Количество()
    print("    Документов:", nd)
    # распечатать первые 8 строк (тип/действие/статус/суммы)
    for i in range(min(nd, 8)):
        d = обр.ТаблицаДокументов.Получить(i)
        print("    -", erp.String(d.ТипДокумента), "|", erp.String(d.Действие), "|", erp.String(d.Статус),
              "| ЕРП=", round(float(d.СуммаЕРП), 2), " Бух=", round(float(d.СуммаБух), 2))
    print("[OK] PHASE2 SMOKE PASS")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print("[FAIL]", info[2] if info else e)
