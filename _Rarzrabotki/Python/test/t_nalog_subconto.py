import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch('V83.COMConnector')
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

pc = buh.ПланыСчетов.Хозрасчетный
codes = ["641","6411","6412","6413","6414","6415","6417","642","651","65"]
for code in codes:
    acc = pc.НайтиПоКоду(code)
    if acc.Пустая():
        print(f"{code}: NOT FOUND")
        continue
    acc_o = acc.ПолучитьОбъект()
    vt = acc_o.ВидыСубконто
    subs = []
    for j in range(vt.Количество()):
        row = vt.Получить(j)
        # ВидСубконто -> ссылка на ПВХ; имя берём через Наименование
        subs.append(str(row.ВидСубконто.Наименование))
    print(f"{code:<6} | субконто: {subs}")
