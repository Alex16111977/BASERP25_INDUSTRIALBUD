"""Знайти всі підрозділи з 'MAN' в найменуванні."""
import win32com.client
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Ссылка,
    Наименование,
    Код
ИЗ Справочник.СтруктураПредприятия
ГДЕ Наименование ПОДОБНО "%MAN%"
    ИЛИ Наименование ПОДОБНО "%KA5697%"
    ИЛИ Наименование ПОДОБНО "%КА5697%"
"""
r = q.Execute().Выгрузить()
print(f"Знайдено: {r.Количество()}")
for i in range(r.Количество()):
    row = r.Получить(i)
    name = str(row.Наименование)
    print(f"  [{i+1}] '{name}' (len={len(name)}) "
          f"hex={name.encode('utf-8').hex()[:80]}... "
          f"Код={row.Код} "
          f"Ref={row.Ссылка}")
