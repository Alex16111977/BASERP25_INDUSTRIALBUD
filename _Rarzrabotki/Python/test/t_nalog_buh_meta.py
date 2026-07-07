import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch('V83.COMConnector')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

rb = buh.Метаданные.РегистрыБухгалтерии.Хозрасчетный
print("=== Измерения РегистрБухгалтерии.Хозрасчетный (BuhBud) ===")
for m in rb.Измерения:
    print("  ИЗМ:", m.Имя)
print("=== Ресурсы ===")
for r in rb.Ресурсы:
    print("  РЕС:", r.Имя)
print("=== Реквизиты ===")
for a in rb.Реквизиты:
    print("  РЕКВ:", a.Имя)
