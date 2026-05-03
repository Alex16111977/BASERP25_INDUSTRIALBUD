import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch('V83.COMConnector')
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject('Запрос')
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1
    Д.Ссылка
ИЗ
    Документ.А_ОказаниеУслугМеждуПодразделениями КАК Д
ГДЕ
    НЕ Д.ПометкаУдаления
"""
r = q.Execute().Выгрузить()
print(f"Знайдено: {r.Количество()}")

if r.Количество() > 0:
    ref = r.Получить(0).Ссылка
    obj = ref.ПолучитьОбъект()
    print(f"Документ: {erp.String(ref)}")

    all_methods = sorted([m for m in dir(obj) if not m.startswith('_')])
    print(f"\nВсього методів: {len(all_methods)}")
    print("\n--- Методи зі 'Создать' / 'Реализац' / 'Услуг' ---")
    for m in all_methods:
        lo = m.lower()
        if 'создат' in lo or 'реализац' in lo or 'услуг' in lo:
            print(f"  >>> {m}")
    print("\n--- Всі методи ---")
    for m in all_methods:
        print(f"  {m}")
else:
    print("Документів не знайдено")
