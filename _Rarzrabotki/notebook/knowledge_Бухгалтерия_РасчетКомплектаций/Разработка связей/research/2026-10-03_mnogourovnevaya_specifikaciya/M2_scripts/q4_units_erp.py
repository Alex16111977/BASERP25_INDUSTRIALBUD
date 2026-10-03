import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
e = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
z = e.NewObject("Запрос")
z.Text = """
ВЫБРАТЬ У.Код КАК Код, У.Наименование КАК Наименование, ПРЕДСТАВЛЕНИЕ(У.Владелец) КАК Владелец, У.ПометкаУдаления КАК Пометка
ИЗ Справочник.УпаковкиЕдиницыИзмерения КАК У
ГДЕ У.Код В (&Коды)
"""
a = e.NewObject("Массив")
for c in ["0306", "0301", "2009", "0101", "0123", "0134", "2938", "0303", "1   ", "7   ", "1", "7", "0102", "5   "]:
    a.Добавить(c)
z.SetParameter("Коды", a)
try:
    t = z.Execute().Выгрузить()
    print("OK", t.Количество())
    for i in range(t.Количество()):
        r = t.Получить(i); print(" ", repr(r.Код), repr(r.Наименование), r.Владелец, r.Пометка)
except Exception as ex:
    print("FAIL:", ex.excepinfo[2] if getattr(ex, 'excepinfo', None) else ex)
