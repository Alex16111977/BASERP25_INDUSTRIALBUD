"""Перепроведение РКО ІБ00-000006 через COM (после загрузки фикса)."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
XS = erp.XMLСтрока

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ Д.Ссылка КАК Ссылка, Д.ХозяйственнаяОперация КАК ХозОп
ИЗ Документ.РасходныйКассовыйОрдер КАК Д
ГДЕ Д.Номер = &Ном
"""
q.SetParameter("Ном", "ІБ00-000006")
sel = q.Execute().Выгрузить()
ref = None
for r in sel:
    if XS(r.ХозОп) == "ВыплатаЗарплатыРаботнику":
        ref = r.Ссылка
assert ref is not None, "РКО не найден"

obj = ref.ПолучитьОбъект()
try:
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print(f"OK: РКО ІБ00-000006 перепроведён, Проведен={obj.Проведен}")
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
    print(f"FAIL repost: {msg}")
    sys.exit(2)
