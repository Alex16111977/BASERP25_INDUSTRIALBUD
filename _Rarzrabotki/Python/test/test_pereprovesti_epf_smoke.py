# -*- coding: utf-8 -*-
"""COM-smoke: rebuilt .epf loads into live BaseERP (types resolve, modules compile).
Does NOT run reposting — user reposts history in UI."""
import sys
import win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PATH = (r"C:\Configuration_downloads\BASERP25\.claude\worktrees\pedantic-greider-ffa593"
        r"\_Rarzrabotki\Обработки\Перепровести Расчеты по контрагентам с выбором за период.epf")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

try:
    obj = erp.ВнешниеОбработки.Создать(PATH, False)  # False = небезопасный режим
    print("SMOKE OK: .epf загружен в BaseERP, объект создан ->", obj is not None)
    try:
        мета = obj.Метаданные()
        print("  Имя обработки:", erp.String(мета.Имя))
        # есть ли ТЧ ТабличнаяЧастьКонтрагенты
        есть_тч = False
        for тч in мета.ТабличныеЧасти:
            if erp.String(тч.Имя) == "ТабличнаяЧастьКонтрагенты":
                есть_тч = True
                кол = [erp.String(r.Имя) for r in тч.Реквизиты]
                print("  ТЧ ТабличнаяЧастьКонтрагенты, реквизиты:", кол)
        print("  ТЧ найдена:", есть_тч)
    except Exception as e2:
        print("  (метаданные не прочитаны:", e2, ")")
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("SMOKE FAIL:", info[2] if info else e)
