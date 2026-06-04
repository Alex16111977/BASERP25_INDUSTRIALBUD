# -*- coding: utf-8 -*-
"""Update Месяц = НачалоМесяца(Дата) for all А_ФинРез_DDS via COM (no posting)."""
import sys
from datetime import datetime
import win32com.client
import pythoncom

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def first_of_month(dt):
    return datetime(dt.year, dt.month, 1, 0, 0, 0)


# List all docs
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ Д.Ссылка КАК Ссылка, Д.Номер КАК Номер, Д.Дата КАК Дата, Д.Месяц КАК Месяц, Д.Проведен КАК Проведен
ИЗ Документ.А_ФинРез_DDS КАК Д
ГДЕ НЕ Д.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Д.Дата
"""
res = q.Execute().Выгрузить()
print(f"Found {res.Количество()} docs")
print()

updated = 0
unchanged = 0
errors = 0

for i in range(res.Количество()):
    r = res.Получить(i)
    target = first_of_month(r.Дата)
    if r.Месяц == target:
        unchanged += 1
        print(f"  [{i+1:>2}] {r.Номер} от {r.Дата.strftime('%d.%m.%Y')} → Месяц={r.Месяц.strftime('%d.%m.%Y')} ✓ already correct")
        continue

    # Update
    try:
        obj = r.Ссылка.ПолучитьОбъект()
        if obj is None:
            print(f"  [{i+1:>2}] {r.Номер} от {r.Дата.strftime('%d.%m.%Y')} ✗ getObject=None")
            errors += 1
            continue

        was_posted = r.Проведен
        old_month = r.Месяц
        obj.Месяц = target
        obj.Записать()  # write without changing posting status

        updated += 1
        print(f"  [{i+1:>2}] {r.Номер} от {r.Дата.strftime('%d.%m.%Y')} → Месяц {old_month.strftime('%d.%m.%Y') if old_month else '(empty)'} → {target.strftime('%d.%m.%Y')} ✓")
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        print(f"  [{i+1:>2}] {r.Номер} от {r.Дата.strftime('%d.%m.%Y')} ✗ {msg[:100]}")
        errors += 1

print()
print(f"Updated: {updated}, Unchanged: {unchanged}, Errors: {errors}")
