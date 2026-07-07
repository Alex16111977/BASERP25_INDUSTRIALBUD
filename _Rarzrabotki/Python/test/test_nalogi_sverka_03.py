# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

# Субконто через объектную модель (без запроса)
коды = ["6411","6412","6413","6414","6415","6417","642","651"]
print("=== СУБКОНТО налоговых счетов BuhBud (объектно) ===")
for код in коды:
    сч = buh.ПланыСчетов.Хозрасчетный.НайтиПоКоду(код)
    if сч.Пустая():
        print(f"  {код}: счёт не найден")
        continue
    об = сч.ПолучитьОбъект()
    субы = []
    for вс in об.ВидыСубконто:
        try:
            об_флаг = " [ТолькоОбороты]" if вс.ТолькоОбороты else ""
        except Exception:
            об_флаг = ""
        субы.append(вс.ВидСубконто.Наименование + об_флаг)
    стр = "; ".join(субы) if субы else "(нет субконто)"
    print(f"  {код:6} | {об.Наименование[:45]:45} | {стр}")
