# -*- coding: utf-8 -*-
"""PL-факт/план по подразделению МД IRS 2026 через А_ОтчетPL.ПолучитьОбъединенныеДанные()."""
import win32com.client, sys, json, datetime
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
def f(v):
    try: return round(float(v),2)
    except: return 0.0
rep = erp.Отчеты.А_ОтчетPL.Создать()
nach = datetime.datetime(2026,7,1,0,0,0); kon = datetime.datetime(2026,10,31,23,59,59)
try:
    tz = rep.ПолучитьОбъединенныеДанные(nach, kon, False, True, True, True, False, True)
except Exception as e:
    print("FAIL:", e.excepinfo[2] if hasattr(e,'excepinfo') and e.excepinfo else e); sys.exit(1)
cols=[S(tz.Колонки.Получить(i).Имя) for i in range(tz.Колонки.Количество())]
print("КОЛОНКИ:", cols)
print("строк всего:", tz.Количество())
rows=[]
for i in range(tz.Количество()):
    r=tz.Получить(i)
    d={c:(f(getattr(r,c)) if any(x in c for x in ("Сумма","Разница","Итог","Ф1","Ф2")) else S(getattr(r,c))) for c in cols}
    rows.append(d)
podr_col = "Подразделение" if "Подразделение" in cols else cols[0]
irs=[d for d in rows if "IRS" in str(d.get(podr_col,""))]
print(f"\nстрок по IRS: {len(irs)}")
uniq=sorted(set(str(d.get(podr_col)) for d in rows))
print("Подразделения (первые 15):", uniq[:15])
json.dump({"cols":cols,"rows":rows}, open("data_pl_raw.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("-> data_pl_raw.json")
