# -*- coding: utf-8 -*-
"""Ф0. Дамп листов «МД IRC 2026» и «СВОД_Производство» принятой книги бюджета (только значения)."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from common import *
import openpyxl
P = os.path.join(ROOT, r"_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\Принятый бюджет\Бюджет_Жовтень26-Грудень26.xlsx")
wb = openpyxl.load_workbook(P, read_only=True, data_only=True)
print([ws.title for ws in wb.worksheets])
for name in sys.argv[1:] or ["МД IRC 2026"]:
    ws = wb[name]
    out = [f"=== ЛИСТ «{name}»"]
    for r in ws.iter_rows():
        vals = [(c.coordinate, c.value) for c in r if getattr(c, "value", None) not in (None, "")]
        if vals:
            out.append(" | ".join(f"{k}={v}" for k, v in vals))
    p = os.path.join(OUT, f"dump_budget_{name}.txt")
    open(p, "w", encoding="utf-8").write("\n".join(out)); print(p, len(out))
