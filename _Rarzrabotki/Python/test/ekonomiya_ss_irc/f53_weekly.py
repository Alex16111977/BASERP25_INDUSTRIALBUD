# -*- coding: utf-8 -*-
"""Ф3. Динамика отчёта директору по неделям: итоговая строка «МД IRC 2026» из xlsx «План-факт виконання СС на …».
Числа — ровно из файлов (без пересчёта)."""
import os, sys, glob
sys.path.insert(0, os.path.dirname(__file__))
from common import *
B = os.path.join(ROOT, r"_Rarzrabotki\Письмо Директору по отчетам на производство")
res = []
for p in sorted(glob.glob(os.path.join(B, "*", "IRC 2026", "План-факт виконання СС на *.xlsx"))):
    wb = load_xlsx(p); ws = wb.worksheets[0]
    hdr = None; tot = None
    for row in ws.iter_rows(values_only=True):
        cells = [c for c in row]
        if hdr is None and any(isinstance(c, str) and "План" in c for c in cells) and any(isinstance(c, str) and "Факт" in c for c in cells):
            hdr = cells
        if any(isinstance(c, str) and c.strip() in ("МД IRC 2026", "МД IRС 2026") for c in cells) and tot is None:
            tot = cells
    res.append((os.path.basename(p), hdr, tot))
for name, hdr, tot in res:
    print(name)
    print("   H:", [h for h in (hdr or []) if h is not None])
    print("   T:", [t for t in (tot or []) if t is not None])
