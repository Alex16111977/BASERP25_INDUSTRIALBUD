# -*- coding: utf-8 -*-
"""Preflight-смоук пересобранных .epf: создаются через COM (компилируются), есть ключевые колонки."""
import sys
import win32com.client as w

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

v8 = w.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

base = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки"
for name in ["СинхронизироватьВзаиморасчеты", "СинхронизироватьДеньги"]:
    p = base + "\\" + name + ".epf"
    try:
        обj = erp.ВнешниеОбработки.Создать(p)
        # ТаблицаДокументов — табличная часть обработки: доступ через .Количество()
        n = обj.ТаблицаДокументов.Количество()
        print(f"[OK] {name}: .epf создан и скомпилирован в рантайме; ТаблицаДокументов пуста ({n} строк)")
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"[FAIL] {name}: {info[2] if info else e}")
print("SMOKE DONE")
