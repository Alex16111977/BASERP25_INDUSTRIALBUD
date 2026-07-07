import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch('V83.COMConnector')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Счет.Код КАК Код,
    Счет.Наименование КАК Наим,
    Счет.Родитель.Код КАК РодКод,
    Счет.Забалансовый КАК Заб
ИЗ
    ПланСчетов.Хозрасчетный КАК Счет
ГДЕ
    Счет.Код ПОДОБНО "64%"
    ИЛИ Счет.Код ПОДОБНО "65%"
УПОРЯДОЧИТЬ ПО Код
"""
r = q.Execute().Выгрузить()
print("BuhBud: налоговые счета 64x/65x, rows:", r.Количество())
for i in range(r.Количество()):
    s = r.Получить(i)
    print(f"{s.Код:<10} | род={str(s.РодКод):<8} | {s.Наим}")
