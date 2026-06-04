# -*- coding: utf-8 -*-
"""READ-ONLY верификация OlapBASERP.Fact_Balance после переноса Свод_ДенежныеСредства.
Проверяет: денежные Source == регистр/ПАП до копейки; Себест не регрессировала;
Cash_ID → Dim_DenezhnyeSredstva и Warehouse_ID → Dim_Warehouses покрытие join."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

TOL = 0.01
ЭТ = {
    "Денежные средства (безналичные)":      50_435_887.99,
    "Денежные средства (наличные)":          24_642_860.47,
    "Денежные средства в пути":                 133_447.50,
    "Денежные средства (у подотчетных лиц)":     53_148.99,
}
ЭТ_СУМ_КО = round(sum(ЭТ.values()), 2)   # 75 265 344,95
ЭТ_СЕБ_КО = 83_627_719.44
fail = []

with get_olap_sql() as c:
    cur = c.cursor()

    print("=== Fact_Balance Period=2026-01 по Source ===")
    cur.execute("""SELECT Source, COUNT(*) c, SUM(Sum_Close) close_,
                          SUM(CASE WHEN Cash_ID IS NOT NULL THEN 1 ELSE 0 END) cash_f,
                          SUM(CASE WHEN Warehouse_ID IS NOT NULL THEN 1 ELSE 0 END) wh_f
                   FROM Fact_Balance WHERE Period_Month='2026-01-01'
                   GROUP BY Source ORDER BY Source""")
    for r in cur.fetchall():
        print(f"  {r.Source:32} rows={r.c:5} Close={float(r.close_ or 0):,.2f} "
              f"cash={r.cash_f} wh={r.wh_f}")

    # 1) Себест регрессия
    seb = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                       "WHERE Period_Month='2026-01-01' AND Source='СебестоимостьТоваров'").fetchval()
    print(f"\nСебест Σ Close = {float(seb):,.2f} (эталон {ЭТ_СЕБ_КО:,.2f})")
    if abs(float(seb) - ЭТ_СЕБ_КО) > TOL:
        fail.append(f"РЕГРЕССИЯ Себест: {float(seb):,.2f} != {ЭТ_СЕБ_КО:,.2f}")

    # 2) денежные по статьям == эталон (== регистр == ПАП == УпрБаланс)
    print("\n=== Денежные Source по PAP_Article ===")
    cur.execute("""SELECT a.PAP_Article_Name nm, SUM(f.Sum_Close) close_, COUNT(*) c
                   FROM Fact_Balance f JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
                   WHERE f.Period_Month='2026-01-01' AND f.Source LIKE N'ДенежныеСредства%'
                   GROUP BY a.PAP_Article_Name ORDER BY a.PAP_Article_Name""")
    got = {}
    for r in cur.fetchall():
        got[str(r.nm)] = float(r.close_ or 0)
        print(f"  {str(r.nm)[:38]:38} Close={float(r.close_ or 0):,.2f} n={r.c}")
    for nm, et in ЭТ.items():
        v = got.get(nm)
        if v is None or abs(v - et) > TOL:
            fail.append(f"«{nm}» Close={v} != эталон {et:,.2f}")
    sum_ko = round(sum(got.values()), 2)
    print(f"\nΣ денежной группы = {sum_ko:,.2f} (эталон {ЭТ_СУМ_КО:,.2f})")
    if abs(sum_ko - ЭТ_СУМ_КО) > TOL:
        fail.append(f"Σ денежной группы {sum_ko:,.2f} != {ЭТ_СУМ_КО:,.2f}")

    # 3) Cash_ID → Dim_DenezhnyeSredstva покрытие (детали денежных)
    miss_cash = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
        WHERE f.Period_Month='2026-01-01' AND f.Source LIKE N'ДенежныеСредства%'
          AND f.Cash_ID IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_DenezhnyeSredstva d WHERE d.Cash_Account_ID=f.Cash_ID)""").fetchval()
    print(f"\nCash_ID без Dim_DenezhnyeSredstva = {miss_cash} (ожид 0)")
    if miss_cash:
        fail.append(f"{miss_cash} строк Cash_ID не join'ятся в Dim_DenezhnyeSredstva")

    # 4) Warehouse_ID → Dim_Warehouses покрытие (Себест)
    has_wh_dim = cur.execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
                             "WHERE TABLE_NAME='Dim_Warehouses'").fetchval()
    if has_wh_dim:
        miss_wh = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
            WHERE f.Period_Month='2026-01-01' AND f.Warehouse_ID IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM Dim_Warehouses d WHERE d.Warehouse_ID=f.Warehouse_ID)""").fetchval()
        wn = cur.execute("SELECT COUNT(*) FROM Dim_Warehouses").fetchval()
        print(f"Dim_Warehouses rows={wn}; Warehouse_ID без Dim = {miss_wh} (ожид 0)")
        if miss_wh:
            fail.append(f"{miss_wh} строк Warehouse_ID не join'ятся в Dim_Warehouses")
    else:
        print("Dim_Warehouses: НЕТ таблицы (DDL ещё не применён)")
        fail.append("Dim_Warehouses не создана")

    # 5) Individual_ID → Dim_Individuals покрытие (подотчётные лица)
    cur.execute("""SELECT COUNT(*) c,
                          SUM(CASE WHEN f.Individual_ID IS NOT NULL THEN 1 ELSE 0 END) ind_f
                   FROM Fact_Balance f
                   WHERE f.Period_Month='2026-01-01'
                     AND f.Source='ДенежныеСредстваУПодотчетныхЛиц'""")
    r = cur.fetchone()
    print(f"\nSource ДенежныеСредстваУПодотчетныхЛиц: rows={r.c} Individual_ID={r.ind_f}")
    miss_ind = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
        WHERE f.Period_Month='2026-01-01' AND f.Individual_ID IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_Individuals d WHERE d.Individual_ID=f.Individual_ID)""").fetchval()
    print(f"Individual_ID без Dim_Individuals = {miss_ind} (ожид 0)")
    if r.ind_f == 0:
        fail.append("подотчёт без Individual_ID (ETL не перезапускался / ФизическоеЛицо пусто)")
    if miss_ind:
        fail.append(f"{miss_ind} строк Individual_ID не join'ятся в Dim_Individuals")

print("\n" + "=" * 64)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: Fact_Balance денежные == регистр/ПАП/УпрБаланс до копейки; "
      "Себест не регрессировала; Cash_ID→Dim_DenezhnyeSredstva, "
      "Warehouse_ID→Dim_Warehouses, Individual_ID→Dim_Individuals "
      "(подотчёт) покрытие 100%")
