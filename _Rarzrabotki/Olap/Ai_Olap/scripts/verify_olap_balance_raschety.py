# -*- coding: utf-8 -*-
"""READ-ONLY верификация OlapBASERP.Fact_Balance после переноса Свод_РасчетыСПартнерами.
Проверяет: расчётные Source КО по статьям == эталон (== регистр == ПАП == УпрБаланс);
Себест/ДенСр не регрессировали; SettlementObj_ID → Dim (ОбъектыРасчетов) покрытие join."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

TOL = 0.01
ЭТ = {
    "Задолженность клиентов":            61_165_524.68,
    "Полученные авансы":                 -48_826_893.59,
    "Выданные авансы":                   68_949_869.33,
    "Задолженность перед поставщиками":  -131_756_106.33,
}
ЭТ_СУМ_КО = round(sum(ЭТ.values()), 2)   # -50 467 605,91
ЭТ_СЕБ_КО = 83_627_719.44
ЭТ_ДЕНСР_КО = 75_265_344.95
DIM = "Dim_ObjektyRaschetov"
fail = []

with get_olap_sql() as c:
    cur = c.cursor()

    print("=== Fact_Balance Period=2026-01 по Source ===")
    cur.execute("""SELECT Source, COUNT(*) c, SUM(Sum_Close) close_,
                          SUM(CASE WHEN SettlementObj_ID IS NOT NULL THEN 1 ELSE 0 END) so_f
                   FROM Fact_Balance WHERE Period_Month='2026-01-01'
                   GROUP BY Source ORDER BY Source""")
    for r in cur.fetchall():
        print(f"  {str(r.Source):34} rows={r.c:5} Close={float(r.close_ or 0):,.2f} "
              f"settlObj={r.so_f}")

    # 1) регрессия Себест / ДенСр
    seb = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                       "WHERE Period_Month='2026-01-01' AND Source='СебестоимостьТоваров'").fetchval()
    den = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                       "WHERE Period_Month='2026-01-01' AND Source LIKE N'ДенежныеСредства%'").fetchval()
    print(f"\nСебест Σ Close = {float(seb):,.2f} (эталон {ЭТ_СЕБ_КО:,.2f})")
    print(f"ДенСр  Σ Close = {float(den):,.2f} (эталон {ЭТ_ДЕНСР_КО:,.2f})")
    if abs(float(seb) - ЭТ_СЕБ_КО) > TOL:
        fail.append(f"РЕГРЕССИЯ Себест: {float(seb):,.2f} != {ЭТ_СЕБ_КО:,.2f}")
    if abs(float(den) - ЭТ_ДЕНСР_КО) > TOL:
        fail.append(f"РЕГРЕССИЯ ДенСр: {float(den):,.2f} != {ЭТ_ДЕНСР_КО:,.2f}")

    # 2) расчётные Source по статьям == эталон
    print("\n=== Расчётные Source по PAP_Article ===")
    cur.execute("""SELECT a.PAP_Article_Name nm, SUM(f.Sum_Close) close_, COUNT(*) c
                   FROM Fact_Balance f JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
                   WHERE f.Period_Month='2026-01-01' AND f.Source LIKE N'РасчетыС%ПоСрокам'
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
    print(f"\nΣ расчётной группы = {sum_ko:,.2f} (эталон {ЭТ_СУМ_КО:,.2f})")
    if abs(sum_ko - ЭТ_СУМ_КО) > TOL:
        fail.append(f"Σ расчётной группы {sum_ko:,.2f} != {ЭТ_СУМ_КО:,.2f}")

    # 3) SettlementObj_ID → Dim покрытие (детали расчётов)
    so_cnt = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
        WHERE f.Period_Month='2026-01-01' AND f.Source LIKE N'РасчетыС%ПоСрокам'
          AND f.SettlementObj_ID IS NOT NULL""").fetchval()
    print(f"\nстрок расчётов с SettlementObj_ID = {so_cnt}")
    has_dim = cur.execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
                           f"WHERE TABLE_NAME='{DIM}'").fetchval()
    if has_dim:
        dn = cur.execute(f"SELECT COUNT(*) FROM {DIM}").fetchval()
        miss = cur.execute(f"""SELECT COUNT(*) FROM Fact_Balance f
            WHERE f.Period_Month='2026-01-01' AND f.SettlementObj_ID IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM {DIM} d WHERE d.SettlementObj_ID=f.SettlementObj_ID)""").fetchval()
        print(f"{DIM} rows={dn}; SettlementObj_ID без Dim = {miss} (ожид 0)")
        if miss:
            fail.append(f"{miss} строк SettlementObj_ID не join'ятся в {DIM}")
    else:
        print(f"{DIM}: НЕТ таблицы (DDL ещё не применён)")
        fail.append(f"{DIM} не создана")

print("\n" + "=" * 64)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: Fact_Balance расчётные == регистр/ПАП/УпрБаланс до копейки; "
      "Себест/ДенСр не регрессировали; SettlementObj_ID→" + DIM + " покрытие 100%")
