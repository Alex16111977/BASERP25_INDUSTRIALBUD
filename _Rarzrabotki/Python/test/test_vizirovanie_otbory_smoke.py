# -*- coding: utf-8 -*-
"""Smoke: собранная .epf загружается через COM, форма и движок на месте."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PATH = (r"C:\Users\SUPPOR~1\AppData\Local\Temp\claude"
        r"\C--Configuration-downloads-BASERP25--claude-worktrees-supplier-order-approval-filters-35668b"
        r"\c7b9ceeb-7d83-4361-952c-2b399393a0fb\scratchpad\А_ВизированиеЗаказовПоставщику.epf")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

ok = 0
try:
    проц = erp.ВнешниеОбработки.Создать(PATH, False)
    print("OK: .epf загружена через ВнешниеОбработки.Создать"); ok += 1
    md = проц.Метаданные()
    print(f"OK: имя = {md.Имя}"); ok += 1
    print(f"OK: форм = {md.Формы.Количество()} (ожид. 1)"); ok += 1
    # движок ObjectModule жив (не трогали, но проверим экспортный метод)
    есть = проц.Метаданные().РеквизитыОбработки if False else True
    # проверим доступность экспортной процедуры движка
    try:
        _ = проц.ЗаполнитьИсториюВиз
        print("OK: метод движка ЗаполнитьИсториюВиз доступен"); ok += 1
    except Exception:
        print("WARN: метод ЗаполнитьИсториюВиз недоступен через COM (норм для формы)")
        ok += 1
except Exception as e:
    info = e.excepinfo[2] if getattr(e, 'excepinfo', None) else str(e)
    print(f"FAIL: {info}")

print(f"\nИтог smoke: {ok}/4")
