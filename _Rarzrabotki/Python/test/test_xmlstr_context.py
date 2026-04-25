# -*- coding: utf-8 -*-
# Перевірка: XMLСтрока у ERP-контексті на COM-UUID з Казни vs у COM-контексті

import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
kazna = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
erp   = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Беремо будь-який документ з Казни
q = kazna.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РаспределениеЗаработнойПлаты КАК Д"
sel = q.Execute().Choose()
sel.Next()
uuid_obj = sel.Ссылка.УникальныйИдентификатор()
print(f"Тип COM об'єкта UUID: {type(uuid_obj)}")

# Варіант 1: COM-контекст Казни (kazna.XMLСтрока)
try:
    r1 = kazna.XMLСтрока(uuid_obj)
    print(f"kazna.XMLСтрока(UUID) = '{r1}'  (довжина {len(r1)})")
except Exception as e:
    print(f"kazna.XMLСтрока ERROR: {e}")

# Варіант 2: COM-контекст ERP (erp.XMLСтрока) — як якби XMLСтрока викликалась у ObjectModule
try:
    r2 = erp.XMLСтрока(uuid_obj)
    print(f"erp.XMLСтрока(UUID)   = '{r2}'  (довжина {len(r2)})")
except Exception as e:
    print(f"erp.XMLСтрока ERROR: {e}")

# Варіант 3: Просто string()
try:
    r3 = kazna.String(uuid_obj)
    print(f"kazna.String(UUID)    = '{r3}'")
except Exception as e:
    print(f"kazna.String ERROR: {e}")

try:
    r4 = erp.String(uuid_obj)
    print(f"erp.String(UUID)      = '{r4}'")
except Exception as e:
    print(f"erp.String ERROR: {e}")

# Можна ли створити УникальныйИдентификатор з того, що повернула XMLСтрока?
for ctx, name in [(kazna, "kazna"), (erp, "erp")]:
    try:
        s = ctx.XMLСтрока(uuid_obj)
        uid = erp.NewObject("УникальныйИдентификатор", s)
        print(f"  {name}.XMLСтрока -> ПолучитьСсылку: {erp.String(erp.Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(uid))}")
    except Exception as e:
        print(f"  {name}.XMLСтрока -> ОШИБКА: {e}")
