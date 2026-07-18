import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch('V83.COMConnector')
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Узнаем реальные имена полей таблицы записей через пустую выборку
q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 * ИЗ РегистрБухгалтерии.Хозрасчетный КАК Дв"""
сел = q.Execute().Выбрать()
кол = q.Execute().Выгрузить().Колонки
print("=== Колонки таблицы записей РегистрБухгалтерии.Хозрасчетный (ERP) ===")
for k in кол:
    print("  ", k.Имя)
