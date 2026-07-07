import pyodbc

DSN = ('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;'
       'DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;TrustServerCertificate=yes')
cn = pyodbc.connect(DSN, timeout=15)
cur = cn.cursor()

print("=== Calendar columns ===")
cur.execute("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS "
            "WHERE TABLE_NAME='Calendar' ORDER BY ORDINAL_POSITION")
cols = cur.fetchall()
for c in cols:
    print(" ", c[0], c[1])

# find a date-typed column
datecols = [c[0] for c in cols if c[1] in ('date', 'datetime', 'datetime2', 'smalldatetime')]
print("\n=== Calendar range per date column ===")
for dc in datecols:
    cur.execute(f"SELECT MIN([{dc}]), MAX([{dc}]), COUNT(*) FROM Calendar")
    r = cur.fetchone()
    print(f"  {dc}: min={r[0]} max={r[1]} count={r[2]}")

# does Calendar contain Nov 2025 days / first-of-month?
for dc in datecols:
    cur.execute(f"SELECT COUNT(*) FROM Calendar WHERE [{dc}] >= '2025-11-01' AND [{dc}] < '2025-12-01'")
    n_nov = cur.fetchone()[0]
    cur.execute(f"SELECT COUNT(*) FROM Calendar WHERE [{dc}] = '2025-11-01'")
    n_nov1 = cur.fetchone()[0]
    cur.execute(f"SELECT COUNT(*) FROM Calendar WHERE [{dc}] = '2025-11-30'")
    n_nov30 = cur.fetchone()[0]
    print(f"  {dc}: days in Nov2025={n_nov}, has 2025-11-01={n_nov1}, has 2025-11-30={n_nov30}")

print("\n=== Fact_Balance Nov rows: distinct Period & Period_Month ===")
cur.execute("SELECT DISTINCT Period, Period_Month FROM Fact_Balance WHERE Period_Month='2025-11-01'")
for r in cur.fetchall():
    print("  Period=", r[0], "Period_Month=", r[1])
