# -*- coding: utf-8 -*-
"""Кошторис МАТЕРИАЛИ.xlsx (листы IRS 15 / IRS 30): позиции + график закупок по месяцам."""
import sys, json, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\А_ПланФактныйПроизводствоПолный\Бюджет\МАТЕРИАЛИ.xlsx"
wb = openpyxl.load_workbook(P, data_only=True, read_only=True)
def num(v): return float(v) if isinstance(v,(int,float)) else 0.0
out={}
for sheet, tag in (("IRS 15","15"),("IRS 30","30")):
    ws = wb[sheet]; rows=[]; sched={"липень":[0.0,0.0],"серпень":[0.0,0.0],"вересень":[0.0,0.0]}
    domov=0
    for r_i,row in enumerate(ws.iter_rows(min_row=5, max_row=ws.max_row, max_col=29, values_only=True),5):
        name = row[1]                       # B
        if not isinstance(name,str) or not name.strip(): continue
        ed,kol_dom,kol_all,dom,cena,summa = row[2],num(row[3]),num(row[4]),num(row[5]),num(row[6]),num(row[7])
        if dom: domov = int(dom)
        li_k,li_s = num(row[8]), num(row[9])     # I,J липень
        se_k,se_s = num(row[10]),num(row[11])    # K,L серпень
        ve_k,ve_s = num(row[12]),num(row[13])    # M,N вересень
        rows.append({"Назва":name.strip(),"Ед":(ed or "").strip(),"КолДом":kol_dom,"КолВсього":kol_all,
                     "Домов":int(dom or 0),"Цена":cena,"СумаДом":summa,
                     "ЛипеньКол":li_k,"ЛипеньСума":li_s,"СерпеньКол":se_k,"СерпеньСума":se_s,
                     "ВересеньКол":ve_k,"ВересеньСума":ve_s})
        sched["липень"][0]+=li_k; sched["липень"][1]+=li_s
        sched["серпень"][0]+=se_k; sched["серпень"][1]+=se_s
        sched["вересень"][0]+=ve_k; sched["вересень"][1]+=ve_s
    sum_dom = round(sum(r["СумаДом"] for r in rows),2)
    out[tag]={"лист":sheet,"домов":domov,"позиций":len(rows),"СумаДом":sum_dom,
              "СумаВсього":round(sum_dom*domov,2),
              "график":{k:[round(v[0],3),round(v[1],2)] for k,v in sched.items()},"rows":rows}
    g=out[tag]["график"]
    print(f"{sheet}: поз={len(rows)}, домів={domov}, Σ/дім={sum_dom:,.2f}, Σ всього={sum_dom*domov:,.2f}")
    print(f"   графік: липень={g['липень'][1]:,.2f}  серпень={g['серпень'][1]:,.2f}  вересень={g['вересень'][1]:,.2f}"
          f"  (Σ={sum(g[k][1] for k in g):,.2f})")
tot = sum(out[t]["СумаВсього"] for t in out)
print(f"\nExcel кошторис разом = {tot:,.2f}   (1С СС = 11 296 725,24; Δ = {tot-11296725.24:,.2f})")
json.dump(out, open("data_excel_koshtoris.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("-> data_excel_koshtoris.json")
