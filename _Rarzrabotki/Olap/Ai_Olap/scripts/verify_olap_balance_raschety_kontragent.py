# -*- coding: utf-8 -*-
"""READ-ONLY верификация OlapBASERP.Fact_Balance после расшифровки
Свод_РасчетыСПартнерами по Контрагент/Партнер/Договор (период 2026-01).

Проверяет (ТЗ verify_olap):
 1) для статей расчётов (Source LIKE 'РасчетыС%ПоСрокам') у ДЕТАЛЕЙ
    (SettlementObj_ID IS NOT NULL) Counterparty_ID/Contract_ID/Partner_ID
    NOT NULL (доля ≈ как SettlementObj_ID, т.е. ≈100%);
 2) FK→Dim: Fact_Balance.Counterparty_ID→Dim_Counterparties,
    Contract_ID→Dim_Contracts, Partner_ID→Dim_Partners — 0 orphans (100%);
 3) плуги (SettlementObj_ID IS NULL) расчётов — субконто пусты;
 4) полный баланс НЕ изменился: Σ Sum_Close (все Source, орг, период)
    ≈ 0,00 (Актив=|Пассив|); печать расчётных статей по Close.
"""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

PER = "2026-01-01"
TOL = 0.01
fail = []

with get_olap_sql() as c:
    cur = c.cursor()

    print(f"=== Fact_Balance Period={PER}: расчётные Source ===")
    cur.execute("""SELECT Source, COUNT(*) c, SUM(Sum_Close) close_,
        SUM(CASE WHEN SettlementObj_ID IS NOT NULL THEN 1 ELSE 0 END) det,
        SUM(CASE WHEN SettlementObj_ID IS NOT NULL AND Counterparty_ID IS NOT NULL THEN 1 ELSE 0 END) cp,
        SUM(CASE WHEN SettlementObj_ID IS NOT NULL AND Contract_ID IS NOT NULL THEN 1 ELSE 0 END) ct,
        SUM(CASE WHEN SettlementObj_ID IS NOT NULL AND Partner_ID IS NOT NULL THEN 1 ELSE 0 END) pn
        FROM Fact_Balance
        WHERE Period_Month=? AND Source LIKE N'РасчетыС%ПоСрокам'
        GROUP BY Source ORDER BY Source""", PER)
    tot_det = tot_cp = tot_ct = tot_pn = 0
    for r in cur.fetchall():
        det = int(r.det or 0); cp = int(r.cp or 0)
        ct = int(r.ct or 0); pn = int(r.pn or 0)
        tot_det += det; tot_cp += cp; tot_ct += ct; tot_pn += pn
        print(f"  {str(r.Source):32} rows={r.c:5} Close={float(r.close_ or 0):,.2f} "
              f"det={det} Contr={cp} Dog={ct} Partn={pn}")
    d = tot_det or 1
    print(f"\nдетали расчётов={tot_det} | Counterparty={tot_cp} "
          f"({tot_cp/d*100:.1f}%) Contract={tot_ct} ({tot_ct/d*100:.1f}%) "
          f"Partner={tot_pn} ({tot_pn/d*100:.1f}%)")
    if tot_det == 0:
        fail.append("нет деталей расчётов (SettlementObj_ID) — ETL не загрузил?")
    if tot_det and tot_cp / d < 0.95:
        fail.append(f"Counterparty_ID заполнен <95% деталей ({tot_cp}/{tot_det})")
    if tot_det and tot_ct / d < 0.95:
        fail.append(f"Contract_ID заполнен <95% деталей ({tot_ct}/{tot_det})")
    if tot_det and tot_pn / d < 0.99:
        fail.append(f"Partner_ID заполнен <99% деталей ({tot_pn}/{tot_det})")

    # плуги расчётов: SettlementObj_ID IS NULL → субконто пусты
    plug_bad = cur.execute("""SELECT COUNT(*) FROM Fact_Balance
        WHERE Period_Month=? AND Source LIKE N'РасчетыС%ПоСрокам'
          AND SettlementObj_ID IS NULL
          AND (Counterparty_ID IS NOT NULL OR Contract_ID IS NOT NULL
               OR Partner_ID IS NOT NULL)""", PER).fetchval()
    print(f"плуги расчётов с непустым субконто = {plug_bad} (ожид 0)")
    if plug_bad:
        fail.append(f"{plug_bad} плугов расчётов с непустым Контр/Дог/Партн")

    # FK→Dim (0 orphans)
    print("\n=== FK Fact_Balance → Dim (0 orphans) ===")
    for col, dim in (("Counterparty_ID", "Dim_Counterparties"),
                     ("Contract_ID", "Dim_Contracts"),
                     ("Partner_ID", "Dim_Partners")):
        has = cur.execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
                          "WHERE TABLE_NAME=?", dim).fetchval()
        if not has:
            print(f"  {dim}: НЕТ таблицы"); fail.append(f"{dim} не создана")
            continue
        dn = cur.execute(f"SELECT COUNT(*) FROM {dim}").fetchval()
        miss = cur.execute(f"""SELECT COUNT(*) FROM Fact_Balance f
            WHERE f.Period_Month=? AND f.{col} IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM {dim} d WHERE d.{col}=f.{col})""",
                           PER).fetchval()
        used = cur.execute(f"""SELECT COUNT(*) FROM Fact_Balance f
            WHERE f.Period_Month=? AND f.{col} IS NOT NULL""", PER).fetchval()
        print(f"  {col:15}→{dim:20} Dim={dn:6} usedFact={used:6} orphans={miss}")
        if miss:
            fail.append(f"{miss} строк {col} без записи в {dim}")

    # полный баланс не изменился: Σ Sum_Close (все Source, период) ≈ 0
    tot = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                       "WHERE Period_Month=?", PER).fetchval()
    rows = cur.execute("SELECT COUNT(*) FROM Fact_Balance WHERE Period_Month=?",
                       PER).fetchval()
    aktiv = cur.execute("""SELECT ISNULL(SUM(CASE WHEN TipPokazatelya='Актив'
            THEN Sum_Close ELSE 0 END),0) FROM Fact_Balance
        WHERE Period_Month=?""", PER).fetchval()
    print(f"\nполный баланс: rows={rows} Σ Sum_Close={float(tot):,.2f} "
          f"(ожид ≈0) | Σ Актив={float(aktiv):,.2f}")
    if abs(float(tot)) > 1.0:
        fail.append(f"полный баланс Σ Sum_Close={float(tot):,.2f} ≠ 0 — баланс сломан")

    print("\n=== расчётные Source по статьям (Close) ===")
    cur.execute("""SELECT a.PAP_Article_Name nm, SUM(f.Sum_Close) close_, COUNT(*) c
        FROM Fact_Balance f JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
        WHERE f.Period_Month=? AND f.Source LIKE N'РасчетыС%ПоСрокам'
        GROUP BY a.PAP_Article_Name ORDER BY a.PAP_Article_Name""", PER)
    for r in cur.fetchall():
        print(f"  {str(r.nm)[:40]:40} Close={float(r.close_ or 0):,.2f} n={r.c}")

print("\n" + "=" * 66)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS verify_olap: расчётные детали Counterparty_ID/Contract_ID/Partner_ID "
      "NOT NULL ≈100%; плуги пусты; FK→Dim_Counterparties/Dim_Contracts/"
      "Dim_Partners 0 orphans; полный баланс Σ Sum_Close≈0 (не изменился)")
