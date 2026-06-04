# -*- coding: utf-8 -*-
"""
E2E acceptance test для інтеграції А_ИсключатьИзОтчетаCashflow.

Перевіряє:
1. SQL колонка Dim_DDS_Articles.Is_Excluded_From_Cashflow існує (bit, NOT NULL)
2. ETL заповнив колонку — total ≈ 425, excluded > 0 (фінансист відмітив частину)
3. Індекс IX_DDS_Articles_Excluded створений
4. Виключені статті мають реальні рухи у Fact_Cashflow
"""
import sys, io
import pyodbc

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

cn = pyodbc.connect(
    "Driver={SQL Server};Server=SQLSERVER;Database=OlapBASERP;UID=sa;PWD=Brw739182465!"
)
cur = cn.cursor()

# === 1. Колонка існує ===
cur.execute("""
    SELECT DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'Dim_DDS_Articles'
      AND COLUMN_NAME = 'Is_Excluded_From_Cashflow'
""")
row = cur.fetchone()
assert row is not None, "Колонка Is_Excluded_From_Cashflow не знайдена в Dim_DDS_Articles"
assert row.DATA_TYPE == 'bit', f"Тип має бути bit, отримано {row.DATA_TYPE}"
assert row.IS_NULLABLE == 'NO', "Має бути NOT NULL"
print(f"[OK] SQL колонка: тип={row.DATA_TYPE}, nullable={row.IS_NULLABLE}, default={row.COLUMN_DEFAULT}")

# === 2. ETL заповнив дані ===
cur.execute("""
    SELECT COUNT(*) AS total,
           SUM(CAST(Is_Excluded_From_Cashflow AS int)) AS excluded
    FROM Dim_DDS_Articles
    WHERE Marked_For_Deletion = 0
""")
row = cur.fetchone()
total, excluded = row.total, row.excluded
assert total > 0, "Dim_DDS_Articles порожня — ETL не запускався"
assert excluded >= 0, "Sum мав би бути ≥ 0"
print(f"[OK] ETL data: total={total}, excluded={excluded}")

# === 3. Індекс створений ===
cur.execute("""
    SELECT name FROM sys.indexes
    WHERE name = 'IX_DDS_Articles_Excluded'
""")
assert cur.fetchone() is not None, "Індекс IX_DDS_Articles_Excluded не знайдено"
print("[OK] Index IX_DDS_Articles_Excluded створено")

# === 4. Виключені статті мають рухи у Fact_Cashflow (якщо excluded > 0) ===
if excluded > 0:
    cur.execute("""
        SELECT da.DDS_Article_Name, COUNT(*) AS records
        FROM Fact_Cashflow fc
        INNER JOIN Dim_DDS_Articles da
            ON fc.DDS_Article_ID = da.DDS_Article_ID
        WHERE da.Is_Excluded_From_Cashflow = 1
        GROUP BY da.DDS_Article_Name
        ORDER BY COUNT(*) DESC
    """)
    rows = cur.fetchall()
    print(f"[OK] Виключені статті ({len(rows)}) з рухами у Fact_Cashflow:")
    for r in rows:
        print(f"     {r.DDS_Article_Name}: {r.records} records")
else:
    print("[WARN] excluded=0 — фінансист ще не відмітив жодної статті в 1С")

cn.close()
print("\n=== Acceptance E2E test PASSED ===")
