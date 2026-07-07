# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

EDRPOU = ['40645273', '41597184', '44590697']
EDRPOU_arr = buh.NewObject("Массив")
for e in EDRPOU:
    EDRPOU_arr.Добавить(e)
# КонецДня(30.11.2025) = 30.11.2025 23:59:59
konec = datetime.datetime(2025, 11, 30, 23, 59, 59)

# Список организаций BuhBud по ЕДРПОУ
qo = buh.NewObject("Запрос")
qo.Text = """
ВЫБРАТЬ Орг.Ссылка КАК Ссылка, Орг.Наименование КАК Наим, Орг.КодПоЕДРПОУ КАК ЕДРПОУ
ИЗ Справочник.Организации КАК Орг
ГДЕ Орг.КодПоЕДРПОУ В (&Список)
"""
qo.SetParameter("Список", EDRPOU_arr)
ro = qo.Execute().Выгрузить()
orgs = []
orgs_arr = buh.NewObject("Массив")
print("Орг BuhBud по ЕДРПОУ:")
for i in range(ro.Количество()):
    s = ro.Получить(i)
    print(f"  {s.Наим} | ЕДРПОУ={s.ЕДРПОУ}")
    orgs.append(s.Ссылка)
    orgs_arr.Добавить(s.Ссылка)

# Счёт 6853
qs = buh.NewObject("Запрос")
qs.Text = """ВЫБРАТЬ Сч.Ссылка КАК Ссылка ИЗ ПланСчетов.Хозрасчетный КАК Сч ГДЕ Сч.Код = "6853" """
счет = qs.Execute().Выгрузить().Получить(0).Ссылка

# Остатки по субконто на конец дня
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Ост.Организация КАК Организация,
    Ост.Субконто1 КАК Получатель,
    Ост.Субконто2 КАК Договор,
    Ост.СуммаОстатокДт КАК СальдоДт,
    Ост.СуммаОстатокКт КАК СальдоКт
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Дата, Счет = &Счет, , Организация В (&Орги)) КАК Ост
"""
q.SetParameter("Дата", konec)
q.SetParameter("Счет", счет)
q.SetParameter("Орги", orgs_arr)
try:
    r = q.Execute().Выгрузить()
    print(f"\nОстатки 6853 на 30.11.2025, рядкiв={r.Количество()}")
    total_dt = 0
    total_kt = 0
    types_seen = {}
    for i in range(r.Количество()):
        s = r.Получить(i)
        dt = s.СальдоДт or 0
        kt = s.СальдоКт or 0
        total_dt += dt
        total_kt += kt
        # тип субконто1
        tname = "Пусто"
        if buh.ЗначениеЗаполнено(s.Получатель):
            tname = s.Получатель.Метаданные().Имя
            types_seen[tname] = types_seen.get(tname, 0) + 1
        pol = ""
        try:
            pol = str(s.Получатель)
        except Exception:
            pol = "?"
        if i < 30:
            print(f"  Орг={s.Организация} | Получатель[{tname}]={pol} | Дт={dt} Кт={kt}")
    print(f"\nИТОГО Дт={total_dt} Кт={total_kt} | Сальдо(Кт-Дт)={total_kt-total_dt}")
    print(f"Типы субконто1: {types_seen}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL: {e.excepinfo[2]}")
    else:
        print(f"FAIL: {e}")
