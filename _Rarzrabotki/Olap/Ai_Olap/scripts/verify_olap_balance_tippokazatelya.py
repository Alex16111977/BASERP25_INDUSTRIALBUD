# -*- coding: utf-8 -*-
"""READ-ONLY верификация OlapBASERP.Fact_Balance.TipPokazatelya
(измерение А_ОтчетБаланс_Свод.ТипПоказателя, формула Отчет.Управленческий
Баланс АктивПассив→Пассив). Проверяет:
  - TipPokazatelya заполнен у ВСЕХ строк (нет NULL); значения ∈ Dim;
  - FK Fact_Balance.TipPokazatelya → Dim_TipPokazatelya 100% (0 orphans);
  - «Налоги» (Dim_PAP_Articles) → TipPokazatelya='Пассив';
  - ПОЛНЫЙ БАЛАНС: Σ Sum_Close по TipPokazatelya (OD-3) 2026-01:
    Актив + Пассив = 0; |Актив| == штатный Отчет.УправленческийБаланс
    288 787 750,11; декабрь 278 093 267,32;
  - 3 месяца присутствуют.
"""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

TOL = 1.0
ЭТ_АКТИВ = {"2025-12-01": 278_093_267.32, "2026-01-01": 288_787_750.11}
OD3 = ("Собственные средства", "Доходы текущего периода",
       "Расходы текущего периода")
fail = []

with get_olap_sql() as c:
    cur = c.cursor()

    dn = cur.execute("SELECT COUNT(*) FROM Dim_TipPokazatelya").fetchval()
    print(f"Dim_TipPokazatelya rows={dn} (ожид 4: Актив/Пассив/АктивПассив/ПустаяСсылка)")
    if dn != 4:
        fail.append(f"Dim_TipPokazatelya rows={dn} != 4")

    nullc = cur.execute("SELECT COUNT(*) FROM Fact_Balance "
                        "WHERE TipPokazatelya IS NULL").fetchval()
    orph = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
        WHERE f.TipPokazatelya IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_TipPokazatelya d
                          WHERE d.TipPokazatelya = f.TipPokazatelya)""").fetchval()
    print(f"Fact_Balance.TipPokazatelya: NULL={nullc} (ожид 0); "
          f"orphans (нет в Dim)={orph} (ожид 0)")
    if nullc:
        fail.append(f"{nullc} строк Fact_Balance.TipPokazatelya = NULL")
    if orph:
        fail.append(f"{orph} строк Fact_Balance.TipPokazatelya без Dim")

    print("\n=== распределение TipPokazatelya (2026-01, OD-3) ===")
    cur.execute("""SELECT f.TipPokazatelya tp, COUNT(*) n, SUM(f.Sum_Close) ko
        FROM Fact_Balance f
        JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
        WHERE f.Period_Month='2026-01-01'
          AND a.PAP_Article_Name NOT IN (N'Собственные средства',
              N'Доходы текущего периода',N'Расходы текущего периода')
        GROUP BY f.TipPokazatelya ORDER BY f.TipPokazatelya""")
    for r in cur.fetchall():
        print(f"  {str(r.tp):14} n={r.n:>6} Σ Close={float(r.ko or 0):,.2f}")

    # «Налоги» → Пассив
    cur.execute("""SELECT DISTINCT f.TipPokazatelya
        FROM Fact_Balance f JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
        WHERE a.PAP_Article_Name=N'Налоги'""")
    nal = [str(r.TipPokazatelya) for r in cur.fetchall()]
    print(f"\n«Налоги» TipPokazatelya = {nal} (ожид ['Пассив'])")
    if nal != ["Пассив"]:
        fail.append(f"«Налоги» TipPokazatelya={nal} != ['Пассив']")

    # ПОЛНЫЙ БАЛАНС по TipPokazatelya (Актив+Пассив=0; |Актив|==отчёт)
    print("\n=== ПОЛНЫЙ БАЛАНС по TipPokazatelya (OD-3) ===")
    for pm, эа in ЭТ_АКТИВ.items():
        cur.execute("""SELECT
            SUM(CASE WHEN f.TipPokazatelya=N'Актив' THEN f.Sum_Close ELSE 0 END) a,
            SUM(CASE WHEN f.TipPokazatelya=N'Пассив' THEN f.Sum_Close ELSE 0 END) p,
            SUM(f.Sum_Close) net
        FROM Fact_Balance f
        JOIN Dim_PAP_Articles d ON d.PAP_Article_ID=f.PAP_Article_ID
        WHERE f.Period_Month=? AND d.PAP_Article_Name NOT IN
            (N'Собственные средства',N'Доходы текущего периода',
             N'Расходы текущего периода')""", pm)
        r = cur.fetchone()
        a = float(r.a or 0); p = float(r.p or 0); net = float(r.net or 0)
        st = "OK" if (abs(net) <= TOL and abs(a - эа) <= TOL) else "!!! FAIL"
        print(f"  {pm}: Актив={a:,.2f} Пассив={p:,.2f} НЕТТО={net:,.2f} "
              f"| эталон |Актив|={эа:,.2f}  {st}")
        if abs(net) > TOL:
            fail.append(f"{pm} НЕТТО {net:,.2f} != 0")
        if abs(a - эа) > TOL:
            fail.append(f"{pm} |Актив| {a:,.2f} != отчёт {эа:,.2f}")

    cur.execute("""SELECT Period_Month, COUNT(*) n FROM Fact_Balance
        WHERE Period_Month IN ('2025-12-01','2026-01-01','2026-02-01')
        GROUP BY Period_Month ORDER BY Period_Month""")
    ms = {str(r.Period_Month)[:10]: r.n for r in cur.fetchall()}
    print(f"\nмесяцы Fact_Balance: {ms}")
    for m in ("2025-12-01", "2026-01-01", "2026-02-01"):
        if ms.get(m, 0) <= 0:
            fail.append(f"нет месяца {m}")

print("\n" + "=" * 64)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: Fact_Balance.TipPokazatelya заполнен (0 NULL), FK→"
      "Dim_TipPokazatelya 100%; «Налоги»→Пассив; ПОЛНЫЙ БАЛАНС по "
      "TipPokazatelya Актив=|Пассив| == штатный Отчет.УправленческийБаланс "
      "(дек 278 093 267,32 / янв 288 787 750,11); 3 месяца")
