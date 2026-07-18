# -*- coding: utf-8 -*-
# COM-smoke: открыть собранную обработку на реальной BaseERP (компиляция модулей против боевого конфига)
import win32com.client as w, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = w.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
path = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты.epf"
try:
    обр = erp.ВнешниеОбработки.Создать(path, False)
    print("[OK] обработка создана:", обр)
    # проверим, что члены объекта на месте
    мета = обр.Метаданные()
    print("    Имя метаданных:", erp.String(мета.Имя))
    тч = [erp.String(t.Имя) for t in мета.ТабличныеЧасти]
    print("    ТЧ:", тч)
    print("[OK] SMOKE PASS")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print("[FAIL]", info[2] if info else e)
