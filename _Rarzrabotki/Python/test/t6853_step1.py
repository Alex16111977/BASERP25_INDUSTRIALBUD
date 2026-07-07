# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Орг.Наименование КАК Наим,
    Орг.КодПоЕДРПОУ КАК ЕДРПОУ,
    Орг.А_ВБалансе КАК ВБалансе
ИЗ Справочник.Организации КАК Орг
ГДЕ Орг.А_ВБалансе = ИСТИНА
"""
try:
    r = q.Execute().Выгрузить()
    print(f"OK орг с А_ВБалансе, рядкiв={r.Количество()}")
    edrpou_list = []
    for i in range(r.Количество()):
        s = r.Получить(i)
        print(f"  {s.Наим} | ЕДРПОУ={s.ЕДРПОУ}")
        if s.ЕДРПОУ:
            edrpou_list.append(s.ЕДРПОУ)
    print("ЕДРПОУ:", edrpou_list)
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL: {e.excepinfo[2]}")
    else:
        print(f"FAIL: {e}")
