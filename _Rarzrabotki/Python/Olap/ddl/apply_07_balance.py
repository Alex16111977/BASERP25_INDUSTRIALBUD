# -*- coding: utf-8 -*-
"""Apply 07_balance.sql to OlapBASERP + verify (idempotent)."""
import sys, io, pyodbc, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
CONN = ("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
        "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;")
sql = pathlib.Path(__file__).with_name("07_balance.sql").read_text(encoding="utf-8")
cx = pyodbc.connect(CONN, autocommit=True)
cu = cx.cursor()
for batch in [b.strip() for b in sql.split("\nGO") if b.strip()]:
    cu.execute(batch)
for t in ("Fact_Balance", "Dim_PAP_Articles"):
    n = cu.execute(
        "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME=?", t
    ).fetchval()
    print(f"{t}: {'OK' if n == 1 else 'MISSING'}")
    assert n == 1, f"FAIL: {t} not created"
cx.close()
print("PASS Task 1 DDL")
