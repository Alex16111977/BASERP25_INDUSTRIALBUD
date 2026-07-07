# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# 1. План счетов ERP — есть ли налоговые счета
print("=== ERP план счетов Хозрасчетный (налоговые) ===")
for код in ["6411","6412","6413","6414","6415","6416","6417","642","651"]:
    try:
        сч = erp.ПланыСчетов.Хозрасчетный.НайтиПоКоду(код)
        if сч.Пустая():
            print(f"  {код}: НЕ НАЙДЕН")
        else:
            об = сч.ПолучитьОбъект()
            субы = []
            for вс in об.ВидыСубконто:
                try:
                    об_флаг = " [ТолькоОбороты]" if вс.ТолькоОбороты else ""
                except Exception:
                    об_флаг = ""
                субы.append(вс.ВидСубконто.Наименование + об_флаг)
            стр = "; ".join(субы) if субы else "(нет субконто)"
            print(f"  {код:6} | {об.Наименование[:40]:40} | {стр}")
    except Exception as e:
        print(f"  {код}: EXC {e}")

# 2. Реквизиты Справочник.Организации в ERP — А_ВБалансе и КодПоЕДРПОУ
print("\n=== ERP Организации: А_ВБалансе + ЕДРПОУ ===")
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Орг.Наименование КАК Наим,
    Орг.КодПоЕДРПОУ КАК ЕДРПОУ,
    Орг.А_ВБалансе КАК ВБалансе
ИЗ Справочник.Организации КАК Орг
ГДЕ НЕ Орг.ПометкаУдаления
УПОРЯДОЧИТЬ ПО ВБалансе УБЫВ, Наим
"""
try:
    r = q.Execute().Выгрузить()
    print(f"  рядкiв={r.Количество()}")
    for s in r:
        print(f"  ВБал={str(s.ВБалансе):5} | ЕДРПОУ={str(s.ЕДРПОУ):12} | {s.Наим}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"  FAIL: {e.excepinfo[2]}")
    else:
        print(f"  FAIL: {e}")
