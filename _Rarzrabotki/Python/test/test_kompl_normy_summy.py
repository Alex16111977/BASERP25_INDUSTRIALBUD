# -*- coding: utf-8 -*-
# Task A2 (Фаза 6): запит норм +СУММА(К.Сумма) КАК НормаСум.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
spec = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000005").Ссылка
q = buh.NewObject("Запрос")
q.УстановитьПараметр("Спецификация", spec)
q.Текст = """
ВЫБРАТЬ
	К.ОбщееНазвание КАК ОбщееНазвание,
	СУММА(К.Количество) КАК Норма,
	СУММА(К.Сумма) КАК НормаСум
ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К
ГДЕ К.Ссылка = &Спецификация
СГРУППИРОВАТЬ ПО К.ОбщееНазвание
"""
try:
    t = q.Выполнить().Выгрузить()
    sN = sum(float(t.Получить(i).Норма) for i in range(t.Количество()))
    sS = sum(float(t.Получить(i).НормаСум) for i in range(t.Количество()))
    print(f"OK еталонів={t.Количество()}, ΣНорма={sN:.3f}, ΣНормаСум={sS:.2f}")
    assert t.Количество() > 0
    print("PASS")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print("FAIL:", info[2] if info else e)
