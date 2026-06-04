# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pyodbc
c = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=OlapBASERP;UID=sa;PWD=Brw739182465!")
cur = c.cursor()
for t in ("Dim_PL_Articles", "Dim_PL_ArticleGroups"):
    cur.execute(f"SELECT name, system_type_id, max_length FROM sys.columns WHERE object_id=OBJECT_ID('dbo.{t}')")
    cols = cur.fetchall()
    print(f"\n=== {t} ===")
    for r in cols:
        print(f"  {r[0]}")
    # Sample data
    cur.execute(f"SELECT TOP 5 * FROM dbo.{t}")
    names = [d[0] for d in cur.description]
    print(f"  Sample 3 rows:")
    for r in cur.fetchmany(5):
        d = dict(zip(names, r))
        print(f"    {d}")
