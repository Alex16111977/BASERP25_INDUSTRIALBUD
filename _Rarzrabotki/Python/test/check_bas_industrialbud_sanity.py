# -*- coding: utf-8 -*-
"""Sanity-check живой базы SQLSERVER/bas_industrialbud перед созданием документа РасчетКомплектаций.
Read-only: подключение + наличие ключевых объектов метаданных + пробный запрос остатков."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
try:
    buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
    print("CONNECT OK: SQLSERVER/bas_industrialbud")
except Exception as e:
    print(f"CONNECT FAIL: {e}")
    sys.exit(1)

md = buh.Метаданные
checks = [
    ("Справочники", "СтруктураСебестоимости"),
    ("Справочники", "ОбщиеНазванияНоменклатуры"),
    ("Справочники", "ЭтапыРабот"),
    ("Справочники", "НазначенияИспользования"),
    ("Справочники", "СпособыОтраженияРасходовПоАмортизации"),
    ("Документы", "КомплектацияНоменклатуры"),
    ("Документы", "ПередачаМалоценныхАктивовВЭксплуатацию"),
    ("Документы", "РасчетКомплектаций"),  # ожидаем НЕТ (ещё не создан)
]
for coll, name in checks:
    obj = getattr(md, coll).Найти(name)
    print(f"{'OK ' if obj is not None else 'НЕТ'} {coll}.{name}")

q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1
    Ост.КоличествоОстаток
ИЗ
    РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет В ИЕРАРХИИ (&Счета), , ) КАК Ост
"""
sch = buh.ПланыСчетов.Хозрасчетный.НайтиПоКоду("20")
arr = buh.NewObject("Массив")
arr.Добавить(sch)
q.SetParameter("Счета", arr)
try:
    r = q.Execute().Выгрузить()
    print(f"QUERY OK, строк={r.Количество()}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"QUERY FAIL: {e.excepinfo[2]}")
    else:
        print(f"QUERY FAIL: {e}")
