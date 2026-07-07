# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

# Субконто налоговых счетов BuhBud
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Сч.Код КАК Код,
    Сч.Наименование КАК Наим,
    ВидыСубконто.НомерСтроки КАК НС,
    ВидыСубконто.ВидСубконто.Наименование КАК ВидСубконто,
    ВидыСубконто.Оборотный КАК Оборотный
ИЗ
    ПланСчетов.Хозрасчетный.ВидыСубконто КАК ВидыСубконто
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ ПланСчетов.Хозрасчетный КАК Сч
    ПО ВидыСубконто.Ссылка = Сч.Ссылка
ГДЕ
    Сч.Код ПОДОБНО "641%"
    ИЛИ Сч.Код ПОДОБНО "642%"
    ИЛИ Сч.Код ПОДОБНО "651%"
УПОРЯДОЧИТЬ ПО Код, НС
"""
try:
    r = q.Execute().Выгрузить()
    print(f"=== СУБКОНТО налоговых счетов BuhBud, рядкiв={r.Количество()} ===")
    for s in r:
        vs = s.ВидСубконто if s.ВидСубконто else "(нет субконто)"
        print(f"  {s.Код:6} | стр{s.НС} | {vs:35} | Оборотный={s.Оборотный}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL subkonto: {e.excepinfo[2]}")
    else:
        print(f"FAIL subkonto: {e}")

print()

# Организации BuhBud: реквизиты
q2 = buh.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ
    Орг.Ссылка КАК Ссылка,
    Орг.Наименование КАК Наим,
    Орг.КодПоЕДРПОУ КАК ЕДРПОУ
ИЗ
    Справочник.Организации КАК Орг
ГДЕ
    НЕ Орг.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Наим
"""
try:
    r2 = q2.Execute().Выгрузить()
    print(f"=== ОРГАНИЗАЦИИ BuhBud, рядкiв={r2.Количество()} ===")
    for s in r2:
        print(f"  ЕДРПОУ={str(s.ЕДРПОУ):12} | {s.Наим}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL org: {e.excepinfo[2]}")
    else:
        print(f"FAIL org: {e}")
