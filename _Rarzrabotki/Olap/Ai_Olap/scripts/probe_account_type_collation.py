# -*- coding: utf-8 -*-
# Read-only probe: коллация БД OlapBASERP, тип/коллация колонки
# Dim_DenezhnyeSredstva.Account_Type, тест кириллицы в varchar(30).
import sys, io, os, pyodbc
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parents[1] / ".env")
conn = pyodbc.connect(os.environ["OLAP_SQL_DSN"], timeout=15, autocommit=True)
cur = conn.cursor()
cur.execute("SELECT DATABASEPROPERTYEX('OlapBASERP','Collation') AS C")
print(f"DB collation = {cur.fetchone()[0]}")
cur.execute(
    "SELECT DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, COLLATION_NAME "
    "FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Dim_DenezhnyeSredstva' "
    "  AND COLUMN_NAME='Account_Type'")
r = cur.fetchone()
print(f"Account_Type column: type={r[0]} len={r[1]} collation={r[2]}")
cur.execute(
    "SELECT TOP 50 Account_Type, COUNT(*) AS N "
    "FROM Dim_DenezhnyeSredstva GROUP BY Account_Type ORDER BY N DESC")
print("\nТекущие Account_Type значения:")
for ac, n in cur.fetchall():
    print(f"  [{ac}] N={n}")
# тест кириллицы roundtrip через CAST AS varchar(30) с дефолтной коллацией
print("\nТест: SELECT CAST(N'БанковскийСчет' AS varchar(30)):")
cur.execute("SELECT CAST(N'БанковскийСчет' AS varchar(30)) AS V")
v = cur.fetchone()[0]
print(f"  [{v}]  (если '???????' — varchar НЕ держит кириллицу без COLLATE)")
cur.execute(
    "SELECT CAST(N'БанковскийСчет' COLLATE Cyrillic_General_CI_AS AS varchar(30)) AS V")
v2 = cur.fetchone()[0]
print(f"\nТест с COLLATE Cyrillic_General_CI_AS: [{v2}]")
conn.close()
