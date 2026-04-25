# -*- coding: utf-8 -*-
import win32com.client, datetime, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")

conn = v8.Connect('Srvr="SQLSERVER";Ref="zup_2";Usr="cfo";Pwd="2442"')
for month in range(3, 0, -1):
    ps = datetime.datetime(2026, month, 1)
    pe = datetime.datetime(2026, month, 28, 23, 59, 59)
    q = conn.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК КолВо ИЗ Документ.НачислениеЗарплатыРаботникам.Начисления КАК Н ГДЕ Н.Ссылка.Дата МЕЖДУ &Н И &К И Н.Ссылка.Проведен = ИСТИНА"
    q.SetParameter("Н", ps)
    q.SetParameter("К", pe)
    r = q.Execute().Choose()
    r.Next()
    print(f"zup_2 {month:02d}/2026: {r.КолВо}", flush=True)
conn = None

conn1 = v8.Connect('Srvr="SQLSERVER";Ref="zup_1";Usr="cfo";Pwd="2442"')
for month in range(3, 0, -1):
    ps = datetime.datetime(2026, month, 1)
    q = conn1.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК КолВо ИЗ РегистрРасчета.ОсновныеНачисленияРаботниковОрганизаций КАК Н ГДЕ НАЧАЛОПЕРИОДА(Н.ПериодРегистрации, МЕСЯЦ) = &П И Н.Результат <> 0"
    q.SetParameter("П", ps)
    r = q.Execute().Choose()
    r.Next()
    print(f"zup_1 {month:02d}/2026: {r.КолВо}", flush=True)
conn1 = None
