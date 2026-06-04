# -*- coding: utf-8 -*-
"""READ-ONLY: какие из _Code/_ParentIDRRef/_Folder/_Description/_Marked есть в _Reference502 (Склады)."""
import sys, io, os, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_baserp_sql

with get_baserp_sql() as c:
    cur = c.cursor()
    cur.execute("""SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                   WHERE TABLE_NAME='_Reference502'
                     AND COLUMN_NAME IN ('_IDRRef','_Code','_Description','_ParentIDRRef','_Folder','_Marked')
                   ORDER BY COLUMN_NAME""")
    have = {r[0] for r in cur.fetchall()}
print("_Reference502 имеет:", sorted(have))
for col in ("_IDRRef", "_Description", "_Marked", "_ParentIDRRef", "_Folder", "_Code"):
    print(f"  {col:16} {'ЕСТЬ' if col in have else 'НЕТ'}")
