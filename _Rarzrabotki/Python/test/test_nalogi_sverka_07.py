# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Структура регистра ПрочиеАктивыПассивы — измерения
print("=== Измерения РН ПрочиеАктивыПассивы ===")
md = erp.Метаданные.РегистрыНакопления.ПрочиеАктивыПассивы
for изм in md.Измерения:
    типы = изм.Тип.Типы()
    тстр = ", ".join([str(т) for т in типы][:3])
    print(f"  Изм: {изм.Имя} | {тстр}")
for рес in md.Ресурсы:
    print(f"  Рес: {рес.Имя}")

# Найдём статью Налоги
print("\n=== Поиск статьи 'Налоги' ===")
qst = erp.NewObject("Запрос")
qst.Text = """
ВЫБРАТЬ Ст.Ссылка КАК Ссылка, Ст.Наименование КАК Наим
ИЗ Справочник.А_СтатьиПрочихАктивовПассивов КАК Ст
ГДЕ Ст.Наименование ПОДОБНО "%алог%"
"""
try:
    r = qst.Execute().Выгрузить()
    for s in r:
        print(f"  Статья: {s.Наим}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"  FAIL статья: {e.excepinfo[2]}")
    else:
        print(f"  FAIL статья: {e}")
