# -*- coding: utf-8 -*-
"""SQL ALTER: добавить Type_Statya nvarchar(50) в Dim_PL_Articles."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pyodbc

conn = pyodbc.connect(
    "Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=OlapBASERP;UID=sa;PWD=Brw739182465!"
)
cur = conn.cursor()
cur.execute("""
    IF COL_LENGTH('dbo.Dim_PL_Articles','Type_Statya') IS NULL
    BEGIN
        ALTER TABLE dbo.Dim_PL_Articles ADD Type_Statya nvarchar(50) NULL;
        PRINT 'Type_Statya добавлена';
    END
    ELSE PRINT 'Type_Statya уже существует';
""")
conn.commit()
# Проверка
cur.execute("SELECT name FROM sys.columns WHERE object_id=OBJECT_ID('dbo.Dim_PL_Articles')")
cols = [r[0] for r in cur.fetchall()]
print(f"Колонки Dim_PL_Articles: {cols}")
assert "Type_Statya" in cols, "Type_Statya НЕ создана!"
print("[OK]")
