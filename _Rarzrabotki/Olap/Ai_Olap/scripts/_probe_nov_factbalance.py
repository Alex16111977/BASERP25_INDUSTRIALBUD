import pyodbc

DSN = ('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;'
       'DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;TrustServerCertificate=yes')

cn = pyodbc.connect(DSN, timeout=15)
cur = cn.cursor()

print("=== Fact_Balance by Period_Month ===")
cur.execute("SELECT Period_Month, COUNT(*) AS cnt, SUM(Sum_Close) AS s "
            "FROM Fact_Balance GROUP BY Period_Month ORDER BY Period_Month")
for r in cur.fetchall():
    print(" ", r[0], "rows=", r[1], "sumClose=", r[2])

print("\n=== distinct Period (raw) around Nov 2025 ===")
cur.execute("SELECT DISTINCT Period FROM Fact_Balance "
            "WHERE Period_Month IN ('2025-11-01','2025-12-01') ORDER BY Period")
for r in cur.fetchall():
    print(" ", r[0])

print("\n=== ETL_Runs last 15 ===")
try:
    cur.execute("SELECT TOP 15 Run_ID, Script, Status, Period, Rows_Loaded, Started_At "
                "FROM ETL_Runs ORDER BY Run_ID DESC")
    for r in cur.fetchall():
        print(" ", r[0], r[1], r[2], "period=", r[3], "rows=", r[4], r[5])
except Exception as e:
    print("  ETL_Runs err:", e)
