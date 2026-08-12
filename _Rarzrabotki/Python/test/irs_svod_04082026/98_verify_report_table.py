# -*- coding: utf-8 -*-
"""Звірка по СКЛАДЕНОМУ ключу (ЗН | Аналітика | Документ) з вивантаженням 1С."""
import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

def num(v): return round(float(v),2) if isinstance(v,(int,float)) else 0.0

# --- вивантаження 1С: рівень з outlineLevel ---
ws = openpyxl.load_workbook(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\А_ПланФактныйПроизводствоПолный\1С\очет в 1С.xlsx").worksheets[0]
S={}; path=[None,None,None,None]
for r in range(11, ws.max_row+1):
    c=ws.cell(r,1); n=c.value
    if not isinstance(n,str) or not n.strip(): continue
    if n.strip()=="Итого": continue
    lv=ws.row_dimensions[r].outlineLevel
    path[lv]=n.strip()
    for j in range(lv+1,4): path[j]=None
    key=tuple(path[1:4])
    v=(num(ws.cell(r,6).value),num(ws.cell(r,8).value),num(ws.cell(r,9).value),num(ws.cell(r,11).value))
    if v==(0,0,0,0): continue
    S[key]=tuple(round(x+y,2) for x,y in zip(S.get(key,(0,0,0,0)),v))

# --- моя таблиця: рівень з outlineLevel рядка ---
wm = openpyxl.load_workbook(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\А_ПланФактныйПроизводствоПолный\1С\План на факт з графіком серпня.xlsx").worksheets[0]
M={}; path=[None,None,None,None]
for r in range(6, wm.max_row+1):
    n=wm.cell(r,1).value
    if not isinstance(n,str) or not n.strip(): continue
    lv=wm.row_dimensions[r].outlineLevel
    path[lv]=n.strip()
    for j in range(lv+1,4): path[j]=None
    key=tuple(path[1:4])
    v=(num(wm.cell(r,3).value),num(wm.cell(r,5).value),num(wm.cell(r,6).value),num(wm.cell(r,8).value))
    if v==(0,0,0,0): continue
    M[key]=tuple(round(x+y,2) for x,y in zip(M.get(key,(0,0,0,0)),v))

common=set(S)&set(M)
bad=[k for k in common if any(abs(S[k][i]-M[k][i])>0.02 for i in range(4))]
print(f"ключів у 1С: {len(S)} | у моїй таблиці: {len(M)} | спільних: {len(common)}")
print(f"РОЗБІЖНОСТЕЙ ПО ЗНАЧЕННЯХ: {len(bad)}")
for k in bad[:12]: print("   ", k, "1С=",S[k]," моє=",M[k])
os_=sorted(set(S)-set(M)); om=sorted(set(M)-set(S))
print(f"тільки в 1С: {len(os_)}", os_[:4])
print(f"тільки в мене: {len(om)}", om[:4])
print("\n" + ("ЗБІГ 1:1 — ТАБЛИЦЯ ІДЕНТИЧНА ЗВІТУ" if not bad and not os_ and not om else "!!! Є РІЗНИЦЯ !!!"))
