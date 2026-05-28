"""Проверка значения константы ИспользоватьНачислениеЗарплаты в БД."""
import sys, win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Прямой получение константы
val = erp.Константы.ИспользоватьНачислениеЗарплаты.Получить()
print(f"Константа ИспользоватьНачислениеЗарплаты = {val}")
print(f"Тип значения: {type(val).__name__}")

# Также через запрос
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Константы.ИспользоватьНачислениеЗарплаты КАК Знач"
r = q.Execute().Выгрузить()
for row in r:
    print(f"Через запрос: {row.Знач}")
