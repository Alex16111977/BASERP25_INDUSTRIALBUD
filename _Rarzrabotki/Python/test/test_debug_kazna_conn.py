# -*- coding: utf-8 -*-
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

# 1. Check if КодСотрудникаКазна field exists in А_РаспределениеКазна
q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q.SetParameter("Н", "000Ц-000001")
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

row0 = obj.А_РаспределениеКазна.Get(0)
print("=== Field check ===")
try:
    val = S(row0.КодСотрудникаКазна)
    print(f"КодСотрудникаКазна field exists: YES, value=[{val}]")
except Exception as e:
    print(f"КодСотрудникаКазна field: ERROR - {e}")

try:
    val = S(row0.ИНН)
    print(f"ИНН field exists: YES, value=[{val}]")
except Exception as e:
    print(f"ИНН field: ERROR - {e}")

# 2. Test COM connection to Kazna from ERP context
print("\n=== Kazna COM test ===")
try:
    v8k = conn.NewObject("COMОбъект", "V83.COMConnector")
    print("COMОбъект created: OK")
except Exception as e:
    print(f"COMОбъект via 1C: ERROR - {e}")
    # Try direct
    try:
        v8k = win32com.client.Dispatch("V83.COMConnector")
        conn_kazn = v8k.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
        print("Direct COM to Kazna: OK")
        Sk = conn_kazn.String

        # Test query
        qk = conn_kazn.NewObject("Запрос")
        qk.Text = "ВЫБРАТЬ ПЕРВЫЕ 3 БДДС.Сотрудник.Код КАК КодС, БДДС.Сотрудник.ИНН КАК ИНН, БДДС.Сотрудник.Наименование КАК ФИО ИЗ РегистрНакопления.БДДС КАК БДДС"
        rk = qk.Execute().Choose()
        while rk.Next():
            print(f"  Kazna: Код={Sk(rk.КодС)}, ИНН={Sk(rk.ИНН)}, ФИО={Sk(rk.ФИО)[:30]}")
    except Exception as e2:
        print(f"Direct COM to Kazna: ERROR - {e2}")

# 3. Check if ExternalConnection context has COMОбъект
print("\n=== ExternalConnection COMОбъект test ===")
try:
    result = conn.Вычислить('ТипЗнч(Новый COMОбъект("V83.COMConnector"))')
    print(f"COMОбъект in ERP context: {S(result)}")
except Exception as e:
    print(f"COMОбъект in ERP context: ERROR - {e}")
    print("This means the procedure CANNOT connect to Kazna from External Connection!")
    print("COMОбъект is NOT available in COM/External Connection mode")
