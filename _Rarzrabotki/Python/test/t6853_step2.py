# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Сч.Код КАК Код,
    Сч.Наименование КАК Наим,
    Сч.Ссылка КАК Ссылка,
    Сч.Забалансовый КАК Заб
ИЗ ПланСчетов.Хозрасчетный КАК Сч
ГДЕ Сч.Код ПОДОБНО "6853%"
"""
try:
    r = q.Execute().Выгрузить()
    print(f"BuhBud счета 6853*, рядкiв={r.Количество()}")
    for i in range(r.Количество()):
        s = r.Получить(i)
        print(f"  Код={s.Код} | {s.Наим} | Заб={s.Заб}")
        obj = s.Ссылка.ПолучитьОбъект()
        for vс in obj.ВидыСубконто:
            try:
                tname = ""
                tt = vс.ВидСубконто.ВидСубконто
                tname = str(tt)
            except Exception:
                tname = "?"
            print(f"     Субконто: {vс.ВидСубконто.Наименование} | ТипЗначения={vс.ВидСубконто.ТипЗначения}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL: {e.excepinfo[2]}")
    else:
        print(f"FAIL: {e}")
