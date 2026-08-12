# -*- coding: utf-8 -*-
"""Бюджет_Серпень26-Жовтень26.xlsx, лист «МД IRS 2026» → статьи ДДС."""
import sys, json, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\А_ПланФактныйПроизводствоПолный\Бюджет\Бюджет_Серпень26-Жовтень26.xlsx"
wb = openpyxl.load_workbook(P, data_only=True, read_only=True)
ws = wb["МД IRS 2026"]
def n(v): return round(float(v),2) if isinstance(v,(int,float)) else 0.0
rows=[]
for r_i,row in enumerate(ws.iter_rows(min_row=6, max_row=201, max_col=25, values_only=True),6):
    A=row[0]
    if not isinstance(A,str) or not A.strip(): continue
    d={"row":r_i,"Стаття":A.strip(),
       "Серпень":n(row[3]),"Вересень":n(row[7]),"Жовтень":n(row[11]),
       "п15":n(row[14]),"п30":n(row[15]),"Разом":n(row[16]),
       "Факт":n(row[17]),"Бюджет":n(row[18]),"БюджетФакт":n(row[19]),"Відхилення":n(row[20])}
    if any(abs(d[k])>0.004 for k in ("Серпень","Вересень","Жовтень","п15","п30","Факт","Бюджет")):
        rows.append(d)
print(f"значимых строк: {len(rows)}\n")
print(f"{'Стаття':<50}{'Серпень':>14}{'Вересень':>14}{'Жовтень':>13}{'Бюджет':>15}{'Факт':>14}")
for d in rows:
    print(f"{d['Стаття'][:48]:<50}{d['Серпень']:>14,.2f}{d['Вересень']:>14,.2f}{d['Жовтень']:>13,.2f}{d['Бюджет']:>15,.2f}{d['Факт']:>14,.2f}")
json.dump(rows, open("data_budget.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n-> data_budget.json")
