# -*- coding: utf-8 -*-
"""DIAG: почему контрагенты за декабрь 2025 не попадают в расшифровку.

Систематическая отладка. Сравниваем декабрь 2025 vs январь 2026 на двух
слоях, чтобы локализовать где рвётся цепочка:

  1С регистр _InfoRg56091 (А_ОтчётБаланс_Свод)  --ETL fact_balance-->  OlapBASERP.Fact_Balance  --PowerQuery-->  PL.pbix

A) OlapBASERP.Fact_Balance: есть ли Dec2025 строки, заполнен ли Counterparty_ID
   по расчётным Source (Свод_РасчетыСПартнерами).
B) BaseERP _InfoRg56091 + _Document56084: для Dec2025 — заполнен ли в самом 1С
   реквизит _Fld56098RRef (Контрагент) у расчётных строк; дата документа-
   регистратора (чтобы понять, перепроведён ли декабрьский свод после
   доработки 2026-05-18).

Только чтение. Запуск: mcp python-runner / прямой python.
"""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap")))

from ai_olap.core.connections import get_olap_sql, get_baserp_sql

OWN_ORG = bytes.fromhex("80D3000C29BBAC2311E653F06BEE36B2")
ZERO16 = b"\x00" * 16

print("=" * 78)
print("A) OlapBASERP.Fact_Balance — Dec2025 vs Jan2026 по Source")
print("=" * 78)
with get_olap_sql() as oc:
    cur = oc.cursor()
    cur.execute("""
        SELECT Period_Month, Source,
               COUNT(*) AS rows_,
               SUM(CASE WHEN Counterparty_ID IS NULL
                         OR Counterparty_ID IN ('00000000000000000000000000000000','')
                        THEN 0 ELSE 1 END) AS cp_filled,
               COUNT(DISTINCT Counterparty_ID) AS cp_distinct
        FROM dbo.Fact_Balance
        WHERE Period_Month IN ('2025-12-01','2026-01-01')
        GROUP BY Period_Month, Source
        ORDER BY Period_Month, Source
    """)
    rows = cur.fetchall()
    if not rows:
        print("  !!! Fact_Balance НЕ содержит строк за 2025-12-01 / 2026-01-01")
    cur_pm = None
    for r in rows:
        pm = str(r.Period_Month)
        if pm != cur_pm:
            print(f"\n  --- Period_Month={pm} ---")
            cur_pm = pm
        flag = "" if r.cp_filled else "  <-- контрагент НЕ заполнен"
        print(f"    {r.Source:<34} rows={r.rows_:>6} cp_filled={r.cp_filled:>6} "
              f"cp_distinct={r.cp_distinct:>5}{flag}")

print()
print("=" * 78)
print("B) BaseERP _InfoRg56091 (+_Document56084) — Dec2025 vs Jan2026")
print("    заполнен ли _Fld56098RRef (Контрагент) в самом 1С регистре")
print("=" * 78)
SQL_B = """
SELECT YEAR(d._Date_Time) AS y, MONTH(d._Date_Time) AS m,
       r._Fld56104RRef AS Source_ref,
       COUNT(*) AS rows_,
       SUM(CASE WHEN r._Fld56098RRef = 0x00000000000000000000000000000000
                THEN 0 ELSE 1 END) AS cp_filled,
       MIN(d._Date_Time) AS dmin, MAX(d._Date_Time) AS dmax,
       COUNT(DISTINCT r._RecorderRRef) AS recorders
FROM _InfoRg56091 r
INNER JOIN _Document56084 d ON d._IDRRef = r._RecorderRRef
WHERE r._Active = 0x01
  AND r._Fld56092RRef = ?
  AND d._Date_Time >= ? AND d._Date_Time < ?
GROUP BY YEAR(d._Date_Time), MONTH(d._Date_Time), r._Fld56104RRef
ORDER BY y, m, r._Fld56104RRef
"""
periods = [("2025-12", "2025-12-01", "2026-01-01"),
           ("2026-01", "2026-01-01", "2026-02-01")]
with get_baserp_sql() as bc:
    cur = bc.cursor()
    for label, lo, hi in periods:
        print(f"\n  --- период {label} (d._Date_Time {lo} .. {hi}) ---")
        cur.execute(SQL_B, OWN_ORG, lo, hi)
        got = cur.fetchall()
        if not got:
            print("    (нет строк регистра за период)")
            continue
        for r in got:
            src_hex = bytes(r.Source_ref).hex().upper() if r.Source_ref else "NULL"
            print(f"    Src={src_hex} rows={r.rows_:>6} cp_filled={r.cp_filled:>6} "
                  f"recorders={r.recorders} dDate=[{r.dmin}..{r.dmax}]")

print()
print("=" * 78)
print("C) Документ-регистратор _Document56084 (А_ФинРез_Баланс) Dec2025 own-org")
print("    дата/проведён — был ли декабрьский свод перепроведён после доработки")
print("=" * 78)
with get_baserp_sql() as bc:
    cur = bc.cursor()
    cur.execute("""
        SELECT d._IDRRef, d._Number, d._Date_Time, d._Posted, d._Marked
        FROM _Document56084 d
        WHERE d._Date_Time >= '2025-12-01' AND d._Date_Time < '2026-02-01'
        ORDER BY d._Date_Time
    """)
    for r in cur.fetchall():
        print(f"    id={bytes(r._IDRRef).hex().upper()} N={r._Number} "
              f"Date={r._Date_Time} Posted={bytes(r._Posted).hex()} "
              f"Marked={bytes(r._Marked).hex()}")

print("\nDONE")
