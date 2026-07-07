# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

qs = buh.NewObject("Запрос")
qs.Text = """ВЫБРАТЬ Сч.Ссылка КАК Ссылка ИЗ ПланСчетов.Хозрасчетный КАК Сч ГДЕ Сч.Код = "6853" """
счет = qs.Execute().Выгрузить().Получить(0).Ссылка

def ostatki(dt):
    q = buh.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
        Ост.Субконто1 КАК Получатель,
        Ост.СуммаОстатокДт КАК Дт,
        Ост.СуммаОстатокКт КАК Кт
    ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Дата, Счет = &Счет) КАК Ост
    """
    q.SetParameter("Дата", dt)
    q.SetParameter("Счет", счет)
    r = q.Execute().Выгрузить()
    tdt=tkt=0
    rows=[]
    for i in range(r.Количество()):
        s=r.Получить(i)
        tdt+=s.Дт or 0; tkt+=s.Кт or 0
        tname = s.Получатель.Метаданные().Имя if buh.ЗначениеЗаполнено(s.Получатель) else "Пусто"
        rows.append((s.ЕДРПОУ, tname, str(s.Получатель), s.Дт or 0, s.Кт or 0))
    return r.Количество(), tdt, tkt, rows

for label, d in [
    ("01.01.2025 нач", datetime.datetime(2025,1,1,0,0,0)),
    ("30.06.2025", datetime.datetime(2025,6,30,23,59,59)),
    ("31.10.2025", datetime.datetime(2025,10,31,23,59,59)),
    ("30.11.2025 КОНЕЦ", datetime.datetime(2025,11,30,23,59,59)),
    ("сейчас 2026-06-20", datetime.datetime(2026,6,20,23,59,59)),
]:
    n,tdt,tkt,rows = ostatki(d)
    print(f"=== {label}: строк={n} Дт={tdt} Кт={tkt} Сальдо(Кт-Дт)={tkt-tdt}")
    for r in rows[:15]:
        print(f"    ЕДРПОУ={r[0]} Пол[{r[1]}]={r[2]} Дт={r[3]} Кт={r[4]}")
