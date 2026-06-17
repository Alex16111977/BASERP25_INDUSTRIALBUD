# -*- coding: utf-8 -*-
# Функциональный smoke Фазы 1: прогон СравнитьОстатки() через COM-объект обработки на BaseERP.
import win32com.client as w, sys, datetime, pywintypes
sys.stdout.reconfigure(encoding='utf-8')
v8 = w.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
path = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты.epf"
обр = erp.ВнешниеОбработки.Создать(path, False)
# Дата полднем — pywintypes сдвигает день, ставим полдень (memory win32com_epf_smoke_gotchas)
обр.НачалоПериода    = pywintypes.Time(datetime.datetime(2026, 1, 1, 12, 0, 0))
обр.ОкончаниеПериода = pywintypes.Time(datetime.datetime(2026, 1, 31, 12, 0, 0))
try:
    res = обр.СравнитьОстатки()
    print("[OK] СравнитьОстатки ->", erp.String(res))
    print("    Контрагентов:", обр.ТаблицаКонтрагентов.Количество())
    print("    Позиций(расхождений):", обр.ТаблицаРасхождений.Количество())
    # сумма |Разница| по своду
    s = 0
    for i in range(обр.ТаблицаКонтрагентов.Количество()):
        s += float(обр.ТаблицаКонтрагентов.Получить(i).СуммаРазницы)
    print("    Σ|Разница| по контрагентам:", round(s, 2))
    print("[OK] PHASE1 SMOKE PASS")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print("[FAIL]", info[2] if info else e)
