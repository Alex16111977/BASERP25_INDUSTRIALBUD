# M2: сверка UUID BuhBud -> ERP (только чтение)
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
b = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
e = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def q(conn, text):
    z = conn.NewObject("Запрос"); z.Text = text
    try:
        return z.Execute().Выгрузить()
    except Exception as ex:
        print("FAIL:", ex.excepinfo[2] if getattr(ex, 'excepinfo', None) else ex); return None

t = q(b, """
ВЫБРАТЬ П.Ссылка КАК Ссылка, П.Наименование КАК Наименование
ИЗ Справочник.Подразделения КАК П
ГДЕ П.Наименование ПОДОБНО "МД IRC 2026 30 м%"
""")
print("BuhBud домов 30 м:", t.Количество())
ok = 0
for i in range(t.Количество()):
    r = t.Получить(i)
    u = b.String(r.Ссылка.УникальныйИдентификатор())
    ref = e.Справочники.СтруктураПредприятия.ПолучитьСсылку(e.NewObject("УникальныйИдентификатор", u))
    nm = e.String(ref.Наименование) if ref.ПолучитьОбъект() is not None else None
    print(" ", repr(r.Наименование), u, "-> ERP:", repr(nm))
    ok += nm is not None
print("совпало UUID:", ok)

t = q(b, """ВЫБРАТЬ С.Ссылка КАК Ссылка, С.Код КАК Код ИЗ Справочник.СтруктураСебестоимости КАК С ГДЕ С.Код В ("000000008", "000000009")""")
for i in range(t.Количество()):
    r = t.Получить(i)
    u = b.String(r.Ссылка.УникальныйИдентификатор())
    ref = e.Справочники.А_СтруктураСебестоимости.ПолучитьСсылку(e.NewObject("УникальныйИдентификатор", u))
    o = ref.ПолучитьОбъект()
    print("СС", r.Код, u, "-> ERP:", (e.String(o.Наименование), o.Комплектующие.Количество()) if o is not None else None)

t = q(b, """ВЫБРАТЬ Е.Ссылка КАК Ссылка, Е.Код КАК Код, Е.Наименование КАК Наименование ИЗ Справочник.КлассификаторЕдиницИзмерения КАК Е""")
ok = 0; miss = []
for i in range(t.Количество()):
    r = t.Получить(i)
    u = b.String(r.Ссылка.УникальныйИдентификатор())
    ref = e.Справочники.УпаковкиЕдиницыИзмерения.ПолучитьСсылку(e.NewObject("УникальныйИдентификатор", u))
    if ref.ПолучитьОбъект() is not None: ok += 1
    else: miss.append(r.Наименование)
print("единиц BuhBud:", t.Количество(), "найдено в ERP по UUID:", ok, "нет:", miss)
