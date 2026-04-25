# -*- coding: utf-8 -*-
import win32com.client, sys
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

print(f"Doc: {S(sel.Ссылка)}")
print(f"BEFORE: Kazna={obj.А_РаспределениеКазна.Count()}")

# Call
print("Calling...")
obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
print("OK")

# Results
afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
anach = obj.НачисленнаяЗарплатаИВзносы.Count()
empty_fl = sum(1 for i in range(afiz) if not VZ(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i).ФизическоеЛицо))
empty_dept = sum(1 for i in range(anach) if not VZ(obj.НачисленнаяЗарплатаИВзносы.Get(i).ПодразделениеПредприятия))
empty_sotr = sum(1 for i in range(obj.А_РаспределениеКазна.Count()) if not VZ(obj.А_РаспределениеКазна.Get(i).Сотрудник))

print(f"\nRESULTS:")
print(f"  NachFiz={afiz}, empty ФизЛицо={empty_fl}")
print(f"  Nach={anach}, empty Подразделение={empty_dept}")
print(f"  Empty Сотрудник in Kazna={empty_sotr}")

if empty_fl == 0 and empty_dept == 0:
    print("\n*** ALL OK ***")
else:
    if empty_fl > 0:
        print(f"\nEmpty ФизЛицо rows:")
        for i in range(afiz):
            r = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i)
            if not VZ(r.ФизическоеЛицо):
                print(f"  [{i}] Dept={S(r.ПодразделениеПредприятия)}, Sum={r.Сумма}, Op={S(r.ВидОперации)}")
                if i > 5:
                    print("  ...")
                    break
