# -*- coding: utf-8 -*-
"""READ-ONLY интроспекция OlapBASERP: какие Dim/Fact есть, состояние Fact_Balance
по денежным Source (после реализации Свод_ДенежныеСредства в 1С)."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

with get_olap_sql() as c:
    cur = c.cursor()
    cur.execute("SELECT name FROM sys.tables ORDER BY name")
    tables = [r[0] for r in cur.fetchall()]
    print("=== Таблицы OlapBASERP ===")
    print(", ".join(tables))

    for t in ("Dim_Items", "Dim_BankAccounts", "Dim_DenezhnyeSredstva",
              "Dim_Warehouses", "Dim_Cashboxes", "Dim_PAP_Articles", "Fact_Balance"):
        if t in tables:
            cur.execute(f"SELECT COUNT(*) FROM [{t}]")
            print(f"{t}: rows={cur.fetchone()[0]}")
        else:
            print(f"{t}: --- НЕТ ТАКОЙ ТАБЛИЦЫ ---")

    if "Dim_DenezhnyeSredstva" in tables:
        cur.execute("SELECT c.name, ty.name FROM sys.columns c JOIN sys.types ty "
                    "ON c.user_type_id=ty.user_type_id WHERE c.object_id=OBJECT_ID('Dim_DenezhnyeSredstva') ORDER BY c.column_id")
        print("Dim_DenezhnyeSredstva cols:", [(n, t) for n, t in cur.fetchall()])

    print("\n=== Fact_Balance: распределение по Source/Period ===")
    cur.execute("""SELECT Period_Month, Source, COUNT(*) c,
                          SUM(Sum_Close) close_,
                          SUM(CASE WHEN Cash_ID IS NOT NULL THEN 1 ELSE 0 END) cash_filled,
                          SUM(CASE WHEN Item_ID IS NOT NULL THEN 1 ELSE 0 END) item_filled,
                          SUM(CASE WHEN Warehouse_ID IS NOT NULL THEN 1 ELSE 0 END) wh_filled
                   FROM Fact_Balance GROUP BY Period_Month, Source ORDER BY Period_Month, Source""")
    for r in cur.fetchall():
        print(f"  {r.Period_Month} | {r.Source:32} | rows={r.c:5} | Close={float(r.close_ or 0):,.2f} "
              f"| cash={r.cash_filled} item={r.item_filled} wh={r.wh_filled}")

    print("\n=== Fact_Balance денежные Source по PAP_Article ===")
    cur.execute("""SELECT a.PAP_Article_Name, f.Source, SUM(f.Sum_Close) close_, COUNT(*) c
                   FROM Fact_Balance f LEFT JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
                   WHERE f.Source LIKE 'ДенежныеСредства%'
                   GROUP BY a.PAP_Article_Name, f.Source ORDER BY a.PAP_Article_Name""")
    rows = cur.fetchall()
    if rows:
        for r in rows:
            print(f"  {str(r.PAP_Article_Name)[:38]:38} | {r.Source:30} | Close={float(r.close_ or 0):,.2f} | n={r.c}")
    else:
        print("  (нет строк с денежными Source — ETL ещё не перезапускался после Свод_ДенежныеСредства)")
print("\nDONE probe")
