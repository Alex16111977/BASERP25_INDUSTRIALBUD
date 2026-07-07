# -*- coding: utf-8 -*-
"""Финальная сверка после перепроведения дек/фев/мар + Refresh Power BI.
Слои: 1С регистр А_ОтчетБаланс_Свод  и  SQL OlapBASERP.Fact_Balance.
Статьи: Оплата труда, Денежные средства (наличные). Эталоны = штатный Управленческий баланс."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def f(x):
    try: return float(x)
    except: return 0.0

MONTHS = [("дек2025","2025-12","ДАТАВРЕМЯ(2025,12,1)"),
          ("янв2026","2026-01","ДАТАВРЕМЯ(2026,1,1)"),
          ("фев2026","2026-02","ДАТАВРЕМЯ(2026,2,1)"),
          ("мар2026","2026-03","ДАТАВРЕМЯ(2026,3,1)")]
ET_OT  = {"дек2025":1913101.80,"янв2026":246187.13,"фев2026":248195.93,"мар2026":-125654.87}
ET_KAS = {"дек2025":21783747.00,"янв2026":23489702.23,"фев2026":18441213.58,"мар2026":10304321.88}
TOL = 0.05
fails = []

def reg(statya, nm_lit):
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Н", statya)
    q.Текст = f"""ВЫБРАТЬ СУММА(Рег.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Рег
    ГДЕ Рег.Статья.Наименование = &Н И Рег.Регистратор.Месяц = {nm_lit}"""
    return f(q.Выполнить().Выгрузить()[0].КО)

def reg_sum(nm_lit):
    q = erp.NewObject("Запрос")
    q.Текст = f"""ВЫБРАТЬ СУММА(Рег.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Рег ГДЕ Рег.Регистратор.Месяц = {nm_lit}"""
    return f(q.Выполнить().Выгрузить()[0].КО)

print("=== СЛОЙ 1: 1С регистр А_ОтчетБаланс_Свод ===")
print(f"{'мес':<8} | {'ОТ регистр':>14} | {'эталон':>14} | {'касса регистр':>15} | {'эталон':>15} | {'Σ КО':>8}")
for label, ym, nm in MONTHS:
    ot = reg("Оплата труда", nm); kas = reg("Денежные средства (наличные)", nm); s = reg_sum(nm)
    ok_ot = abs(ot-ET_OT[label])<=TOL; ok_kas = abs(kas-ET_KAS[label])<=TOL; ok_s = abs(s)<=0.5
    if not ok_ot: fails.append(f"1C ОТ {label}")
    if not ok_kas: fails.append(f"1C касса {label}")
    if not ok_s: fails.append(f"1C Σ {label}")
    print(f"{label:<8} | {ot:>14,.2f} | {ET_OT[label]:>14,.2f} | {kas:>15,.2f} | {ET_KAS[label]:>15,.2f} | {s:>8,.2f} {'✓' if ok_ot and ok_kas and ok_s else 'FAIL'}")

print("\n=== СЛОЙ 2: SQL OlapBASERP.Fact_Balance (то, что читает Power BI) ===")
try:
    import pyodbc
    cn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!", timeout=10)
    cur = cn.cursor()
    def sql_art(ym, art):
        cur.execute("""SELECT SUM(f.Sum_Close) FROM Fact_Balance f JOIN Dim_PAP_Articles d ON f.PAP_Article_ID=d.PAP_Article_ID
            WHERE CONVERT(varchar(7),f.Period,120)=? AND d.PAP_Article_Name=?""", ym, art)
        r = cur.fetchone()[0]; return f(r) if r is not None else 0.0
    def sql_sum(ym):
        cur.execute("SELECT SUM(Sum_Close) FROM Fact_Balance WHERE CONVERT(varchar(7),Period,120)=?", ym)
        r = cur.fetchone()[0]; return f(r) if r is not None else 0.0
    print(f"{'мес':<8} | {'ОТ SQL':>14} | {'эталон':>14} | {'касса SQL':>15} | {'эталон':>15} | {'Σ Close':>9}")
    for label, ym, nm in MONTHS:
        ot = sql_art(ym,"Оплата труда"); kas = sql_art(ym,"Денежные средства (наличные)"); s = sql_sum(ym)
        ok_ot = abs(ot-ET_OT[label])<=TOL; ok_kas = abs(kas-ET_KAS[label])<=TOL; ok_s = abs(s)<=0.5
        if not ok_ot: fails.append(f"SQL ОТ {label}")
        if not ok_kas: fails.append(f"SQL касса {label}")
        if not ok_s: fails.append(f"SQL Σ {label}")
        print(f"{label:<8} | {ot:>14,.2f} | {ET_OT[label]:>14,.2f} | {kas:>15,.2f} | {ET_KAS[label]:>15,.2f} | {s:>9,.2f} {'✓' if ok_ot and ok_kas and ok_s else 'FAIL'}")
except Exception as e:
    print(f"  [pyodbc: {e}]"); fails.append("SQL unreachable")

print("\n" + "="*64)
print("ВСЁ СОШЛОСЬ — 1С регистр и SQL Fact_Balance == штатный баланс по всем месяцам." if not fails
      else f"ЕСТЬ РАСХОЖДЕНИЯ ({len(fails)}): {fails}")
