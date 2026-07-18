# -*- coding: utf-8 -*-
# Перевіряємо конкретний UUID з debug_a_otrazhenie.txt
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

UUID_STR = "d4d9b293-dfcc-11f0-8104-00155dce3d04"
print(f"Перевіряємо UUID: {UUID_STR}")

try:
    uid = erp.NewObject("УникальныйИдентификатор", UUID_STR)
    print(f"  Новый УникальныйИдентификатор: OK, type={type(uid)}")
except Exception as e:
    print(f"  ERROR створення UUID: {e}")
    sys.exit(1)

try:
    ref = erp.Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(uid)
    print(f"  ПолучитьСсылку: {erp.String(ref)}")
except Exception as e:
    print(f"  ERROR ПолучитьСсылку: {e}")
    sys.exit(1)

try:
    obj = ref.ПолучитьОбъект()
    if obj is None:
        print(f"  ПолучитьОбъект: None (бите посилання)")
    else:
        print(f"  ПолучитьОбъект: OK, Дата={erp.String(obj.Дата)}, Номер={erp.String(obj.Номер)}")
except Exception as e:
    print(f"  ERROR ПолучитьОбъект: {e}")

# Також перевіряємо чи Ссылка — це Документы.А_РаспределениеЗаработнойПлаты
print(f"\n  erp.ЗначениеЗаполнено(ref): {erp.ЗначениеЗаполнено(ref)}")

# Спробуємо знайти документ через запит за UUID
try:
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ Д.Ссылка, Д.Дата, Д.Номер ИЗ Документ.А_РаспределениеЗаработнойПлаты КАК Д ГДЕ Д.Ссылка = &С"
    q.SetParameter("С", ref)
    r = q.Execute().Выгрузить()
    print(f"  Запит за Ссылкой: {r.Количество()} рядків")
    if r.Количество() > 0:
        row = r.Получить(0)
        print(f"    → {erp.String(row.Ссылка)}, Дата={erp.String(row.Дата)}")
except Exception as e:
    print(f"  ERROR запиту: {e}")
