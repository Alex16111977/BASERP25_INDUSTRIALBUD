# -*- coding: utf-8 -*-
# Discovery: чи Этап функціонально визначений ОбщееНазвание у спец (1 еталон = 1 етап)?
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
spec = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000005").Ссылка
q = buh.NewObject("Запрос")
q.УстановитьПараметр("С", spec)
q.Текст = """
ВЫБРАТЬ РАЗЛИЧНЫЕ К.ОбщееНазвание КАК Эталон, К.Этап КАК Этап
ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К
ГДЕ К.Ссылка = &С
"""
t = q.Выполнить().Выгрузить()
по_эталону = {}
for i in range(t.Количество()):
    r = t.Получить(i)
    э = buh.String(r.Эталон)
    эт = buh.String(r.Этап) if (r.Этап is not None and not r.Этап.Пустая()) else "<ПУСТО>"
    по_эталону.setdefault(э, set()).add(эт)
mult = {э: s for э, s in по_эталону.items() if len(s) > 1}
print(f"еталонів={len(по_эталону)}; з >1 етапом={len(mult)}")
for э, s in list(mult.items())[:15]:
    print(f"  '{э}' -> етапи: {s}")
# приклад мапінгу етап по еталону
print("--- приклади (еталон -> етап) ---")
for э, s in list(по_эталону.items())[:10]:
    print(f"  '{э}' -> {list(s)[0]}")
