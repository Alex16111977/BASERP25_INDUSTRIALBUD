# -*- coding: utf-8 -*-
import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8=win32com.client.Dispatch("V83.COMConnector")
buh=v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
try:
    f=buh.Документы.РасчетКомплектаций.ПолучитьФорму("ФормаДокумента")
    print("ФОРМА ПОЛУЧЕНА OK — реквизиты формы связаны без ошибок")
except Exception as e:
    info=getattr(e,'excepinfo',None)
    print("FAIL:", info[2] if info and len(info)>2 else e)
