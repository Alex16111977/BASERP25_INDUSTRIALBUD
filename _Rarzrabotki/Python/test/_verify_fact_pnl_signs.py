# -*- coding: utf-8 -*-
"""Acceptance: Fact_PnL знаки соответствуют Dim_PL_Articles.Type_Statya."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pyodbc

conn = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=OlapBASERP;UID=sa;PWD=Brw739182465!")
cur = conn.cursor()

# Узнать имена колонок Fact_PnL
cur.execute("SELECT TOP 0 * FROM Fact_PnL")
cols = [c[0] for c in cur.description]
print(f"Колонки Fact_PnL ({len(cols)}): {cols}")

# Подобрать sum колонку
sum_col = next((c for c in cols if 'sum' in c.lower() and 'plan' not in c.lower() and 'fact' not in c.lower()), None)
# Если есть Sum_Fact, Sum_Plan и т.д. — попробуем Sum_Fact
sum_fact_col = next((c for c in cols if c.lower() == 'sum_fact'), None)
target_col = sum_fact_col or sum_col or 'Sum'
print(f"\nИспользую колонку: {target_col}")

cur.execute(f"""
SELECT
    COALESCE(d.Type_Statya, '(NULL)') AS Type_Statya,
    COUNT(*) AS Rows,
    SUM(f.[{target_col}]) AS Sum_Signed,
    SUM(CASE WHEN f.[{target_col}]<0 THEN -f.[{target_col}] ELSE f.[{target_col}] END) AS Sum_Abs
FROM Fact_PnL f
LEFT JOIN Dim_PL_Articles d ON d.PL_Article_ID = f.PL_Article_ID
GROUP BY d.Type_Statya
ORDER BY 4 DESC
""")
print(f"\n{'Type':<25} {'Rows':>8} {'Sum_Signed':>20} {'Sum_Abs':>20}")
print("-" * 80)
total_signed = 0
total_abs = 0
for r in cur.fetchall():
    t, rows, sumS, sumA = r[0], r[1], float(r[2] or 0), float(r[3] or 0)
    print(f"{t:<25} {rows:>8} {sumS:>20,.2f} {sumA:>20,.2f}")
    total_signed += sumS
    total_abs += sumA

print("-" * 80)
print(f"P&L (Σ signed) = {total_signed:,.2f}  ({'УБЫТОК' if total_signed < 0 else 'ПРИБЫЛЬ'})")
