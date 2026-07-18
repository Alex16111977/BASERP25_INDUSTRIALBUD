# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

# Найдём налоговые счета в плане счетов BuhBud
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Сч.Код КАК Код,
    Сч.Наименование КАК Наим,
    Сч.Вид КАК Вид,
    Сч.Забалансовый КАК Забаланс
ИЗ
    ПланСчетов.Хозрасчетный КАК Сч
ГДЕ
    Сч.Код ПОДОБНО "641%"
    ИЛИ Сч.Код ПОДОБНО "642%"
    ИЛИ Сч.Код ПОДОБНО "651%"
УПОРЯДОЧИТЬ ПО Код
"""
try:
    r = q.Execute().Выгрузить()
    print(f"OK BuhBud план счетов, рядкiв={r.Количество()}")
    for s in r:
        print(f"  {s.Код:8} | {s.Наим[:55]:55} | Вид={buh.XMLСтрока(s.Вид)} | Забал={s.Забаланс}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL: {e.excepinfo[2]}")
    else:
        print(f"FAIL: {e}")
