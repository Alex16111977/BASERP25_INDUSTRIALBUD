# -*- coding: utf-8 -*-
"""Предпосылки А_НастройкиПодключенияБаз:
1) все 4 базы доступны через Srvr=localhost;
2) БСП ОбщегоНазначения.УстановитьВнешнееСоединениеСБазой работает из контекста BaseERP.
"""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")

BASES = ["zup", "zup2", "bas_industrialbud", "kazna"]
fails = 0
for ref in BASES:
    try:
        conn = v8.Connect(f'Srvr="localhost";Ref="{ref}";Usr="cfo";Pwd="2442"')
        print(f"OK localhost {ref}: {conn.Metadata.Synonym}")
        conn = None
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else e
        print(f"FAIL localhost {ref}: {info}")
        fails += 1

# БСП-функция из контекста ERP
try:
    erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    p = erp.ОбщегоНазначенияКлиентСервер.СтруктураПараметровДляУстановкиВнешнегоСоединения()
    p.ВариантРаботыИнформационнойБазы = 1
    p.ИмяСервера1СПредприятия = "localhost"
    p.ИмяИнформационнойБазыНаСервере1СПредприятия = "bas_industrialbud"
    p.ИмяПользователя = "cfo"
    p.ПарольПользователя = "2442"
    r = erp.ОбщегоНазначения.УстановитьВнешнееСоединениеСБазой(p)
    if r.Соединение is not None:
        print("OK БСП УстановитьВнешнееСоединениеСБазой: соединение получено")
    else:
        print(f"FAIL БСП: {r.ПодробноеОписаниеОшибки}")
        fails += 1
except Exception as e:
    info = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else e
    print(f"FAIL БСП (исключение): {info}")
    fails += 1

print("ИТОГ:", "OK" if fails == 0 else f"FAIL ({fails})")
