# -*- coding: utf-8 -*-
"""Ф0. Дамп файлов-источников в текст: Списання IRC.xlsx, МАТЕРИАЛИ.xlsx, книга бюджета, план-факт 30.09."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from common import *

B = os.path.join(ROOT, "_Rarzrabotki")
FILES = [
    (os.path.join(B, r"Письмо Директору по отчетам на производство\Отчет\Списання IRC.xlsx"), None),
    (os.path.join(B, r"Бюджет прогноз\Бюджет октябрь 2026\МАТЕРИАЛИ.xlsx"), None),
]
only = sys.argv[1] if len(sys.argv) > 1 else None

for path, sheets in FILES:
    if only and only not in path:
        continue
    wb = load_xlsx(path)
    out = []
    out.append(f"ФАЙЛ {path}")
    for ws in wb.worksheets:
        out.append(f"=== ЛИСТ «{ws.title}» {ws.max_row}x{ws.max_column}")
        for r in ws.iter_rows():
            vals = [(c.coordinate, c.value) for c in r if c.value not in (None, "")]
            if vals:
                out.append(" | ".join(f"{k}={v}" for k, v in vals))
    name = os.path.splitext(os.path.basename(path))[0] + ".txt"
    p = os.path.join(OUT, "dump_" + name)
    open(p, "w", encoding="utf-8").write("\n".join(out))
    print(p, len(out))
