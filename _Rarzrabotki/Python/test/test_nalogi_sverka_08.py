# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

md = erp.Метаданные.РегистрыНакопления.ПрочиеАктивыПассивы
print("=== Типы измерений ПрочиеАктивыПассивы ===")
for изм in md.Измерения:
    типы = изм.Тип.Типы()
    имена = []
    for i in range(типы.Количество()):
        т = типы.Получить(i)
        try:
            мдт = erp.Метаданные.НайтиПоТипу(т)
            имена.append(мдт.ПолноеИмя() if мдт else "примитив")
        except Exception:
            имена.append("?")
    print(f"  {изм.Имя}: {'; '.join(имена)}")

# Уникальные Статья ПОДОБНО Налог
print("\n=== Уникальные Статья в движениях ПрочиеАктивыПассивы (ПОДОБНО Налог) ===")
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ РАЗЛИЧНЫЕ
    Дв.Статья КАК Статья,
    ПРЕДСТАВЛЕНИЕ(Дв.Статья) КАК СтатьяПредст
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Дв
ГДЕ ПРЕДСТАВЛЕНИЕ(Дв.Статья) ПОДОБНО "%алог%"
"""
try:
    r = q.Execute().Выгрузить()
    print(f"  рядкiв={r.Количество()}")
    for s in r:
        тип = s.Статья.Метаданные().ПолноеИмя() if erp.ЗначениеЗаполнено(s.Статья) else "пусто"
        print(f"  Статья='{s.СтатьяПредст}' тип={тип}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"  FAIL: {e.excepinfo[2]}")
    else:
        print(f"  FAIL: {e}")
