# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pyodbc

conn = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=OlapBASERP;UID=sa;PWD=Brw739182465!")
cur = conn.cursor()
cur.execute("SELECT Type_Statya, COUNT(*) FROM Dim_PL_Articles WHERE Marked_For_Deletion=0 GROUP BY Type_Statya ORDER BY 2 DESC")
print("Распределение Type_Statya в Dim_PL_Articles (без удалённых):")
for r in cur.fetchall():
    print(f"  {r[0]!r:30}  {r[1]}")
