# -*- coding: utf-8 -*-
import win32com.client, sys, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q.SetParameter("Н", "000Ц-000001")
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

# BEFORE
empty_sotr = sum(1 for i in range(obj.А_РаспределениеКазна.Count()) if not VZ(obj.А_РаспределениеКазна.Get(i).Сотрудник))
empty_inn = sum(1 for i in range(obj.А_РаспределениеКазна.Count()) if not S(obj.А_РаспределениеКазна.Get(i).ИНН).strip())
print(f"BEFORE: Kazna={obj.А_РаспределениеКазна.Count()}, empty Sotr={empty_sotr}, empty INN={empty_inn}")

# CALL
print("Calling procedure...")
obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
print("OK")

# AFTER
empty_sotr2 = sum(1 for i in range(obj.А_РаспределениеКазна.Count()) if not VZ(obj.А_РаспределениеКазна.Get(i).Сотрудник))
empty_inn2 = sum(1 for i in range(obj.А_РаспределениеКазна.Count()) if not S(obj.А_РаспределениеКазна.Get(i).ИНН).strip())
empty_fl = sum(1 for i in range(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()) if not VZ(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i).ФизическоеЛицо))
empty_dept = sum(1 for i in range(obj.НачисленнаяЗарплатаИВзносы.Count()) if not VZ(obj.НачисленнаяЗарплатаИВзносы.Get(i).ПодразделениеПредприятия))

print(f"AFTER: empty Sotr={empty_sotr2}, empty INN={empty_inn2}")
print(f"  NachFiz={obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()}, empty FL={empty_fl}")
print(f"  Nach={obj.НачисленнаяЗарплатаИВзносы.Count()}, empty Dept={empty_dept}")
print(f"  NDFL={obj.НачисленныйНДФЛ.Count()}")

# Check Безсмертній
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    fio = S(row.ФИО).strip()
    if "Безсмертн" in fio:
        print(f"\n  Безсмертній: Сотр={S(row.Сотрудник)[:30]}, ІНН={S(row.ИНН)}, Код={S(row.КодСотрудникаКазна)}")
        break
