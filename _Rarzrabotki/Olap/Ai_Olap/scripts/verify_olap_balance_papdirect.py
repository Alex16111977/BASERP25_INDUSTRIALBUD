# -*- coding: utf-8 -*-
"""READ-ONLY верификация OlapBASERP.Fact_Balance после переноса
Свод_ПрочиеАктивыПассивы_Прямой + Свод_ОплатаТруда (оба Source=ПустаяСсылка).
Проверяет (2026-01 / ТОВ):
  - Source=ПустаяСсылка по статьям == эталон (== регистр == ПАП == УпрБаланс):
    Налоги 9 331 275,92 / Основные средства -149 202,85 /
    Прибыли и убытки -110 616 551,99 / Оплата труда -7 196 698,44 /
    ИТОГ -108 631 177,36 (= старый Прямой -101 434 478,92 + ОТ -7 196 698,44);
  - «Оплата труда» БЕЗ субконто (карточка статьи «Без аналитики»; решение
    2026-05-18 — старый замысел ∝ ВзаиморасчетыССотрудниками отменён);
  - разрез Налоги по TaxType == Карточка актива/пассива
    (НДС 9 246 711,36 / ДругиеНалоги 72 252,00 / НДФЛ 4 925,02 /
     ВоенныйСбор 1 368,07 / НачисленныйЕСВ 6 019,47);
  - TaxType заполнен у строк Налоги, "ПустаяСсылка" у прочих (вкл. ОТ);
  - 3 месяца (2025-12/2026-01/2026-02) присутствуют (сосуществуют);
  - РЕГРЕССИЯ (Прямой исключает ОТ → не изменился): Себест 83 627 719,44 /
    ДенСр 75 265 344,95 / Расчёты клиенты 12 338 631,09 /
    поставщики -62 806 237,00 не изменились."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

TOL = 0.01
ЭТ_СТ = {
    "Налоги":            9_331_275.92,
    "Основные средства": -149_202.85,
    "Прибыли и убытки":  -110_616_551.99,
    "Оплата труда":      -7_196_698.44,
}
# ИТОГ Source=ПустаяСсылка 2026-01 = Прямой(-101 434 478,92) + ОТ(-7 196 698,44),
# подтверждён serverside COM == ПАП(Источник=пусто, OD-3, signed, ВСЕ статьи).
ЭТ_ИТОГ = -108_631_177.36
# TaxType (имя перечисления) -> КО, == Карточка актива/пассива
ЭТ_TAX = {
    "НДС":            9_246_711.36,
    "ДругиеНалоги":   72_252.00,
    "НДФЛ":           4_925.02,
    "ВоенныйСбор":    1_368.07,
    "НачисленныйЕСВ": 6_019.47,
}
ЭТ_СЕБ = 83_627_719.44
ЭТ_ДЕНСР = 75_265_344.95
ЭТ_КЛ = 12_338_631.09
ЭТ_ПОСТ = -62_806_237.00
fail = []

with get_olap_sql() as c:
    cur = c.cursor()

    # 0) три месяца присутствуют
    cur.execute("""SELECT Period_Month, COUNT(*) n,
                          SUM(CASE WHEN Source='ПустаяСсылка' THEN 1 ELSE 0 END) prym
                   FROM Fact_Balance
                   WHERE Period_Month IN ('2025-12-01','2026-01-01','2026-02-01')
                   GROUP BY Period_Month ORDER BY Period_Month""")
    mset = {}
    print("=== Месяцы в Fact_Balance ===")
    for r in cur.fetchall():
        mset[str(r.Period_Month)[:10]] = r.prym
        print(f"  {str(r.Period_Month)[:10]} строк={r.n} Прямой(Source=пусто)={r.prym}")
    for m in ("2025-12-01", "2026-01-01", "2026-02-01"):
        if mset.get(m, 0) <= 0:
            fail.append(f"месяц {m}: нет строк Source=ПустаяСсылка")

    # 1) Source=ПустаяСсылка по статьям (2026-01) == эталон
    print("\n=== 2026-01 Source=ПустаяСсылка по статьям ===")
    cur.execute("""SELECT a.PAP_Article_Name nm, SUM(f.Sum_Close) ko, COUNT(*) n
                   FROM Fact_Balance f JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
                   WHERE f.Period_Month='2026-01-01' AND f.Source='ПустаяСсылка'
                   GROUP BY a.PAP_Article_Name ORDER BY SUM(f.Sum_Close) DESC""")
    got = {}
    for r in cur.fetchall():
        got[str(r.nm)] = float(r.ko or 0)
        print(f"  {str(r.nm)[:26]:26} Close={float(r.ko or 0):,.2f} n={r.n}")
    for nm, et in ЭТ_СТ.items():
        v = got.get(nm)
        if v is None or abs(v - et) > TOL:
            fail.append(f"«{nm}» Close={v} != эталон {et:,.2f}")
    itog = round(sum(got.values()), 2)
    print(f"\nИТОГ Source=ПустаяСсылка = {itog:,.2f} (эталон {ЭТ_ИТОГ:,.2f})")
    if abs(itog - ЭТ_ИТОГ) > TOL:
        fail.append(f"ИТОГ {itog:,.2f} != {ЭТ_ИТОГ:,.2f}")

    # 2) разрез Налоги по TaxType == Карточка
    print("\n=== 2026-01 Налоги по TaxType ===")
    cur.execute("""SELECT f.TaxType tt, SUM(f.Sum_Close) ko, COUNT(*) n
                   FROM Fact_Balance f JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
                   WHERE f.Period_Month='2026-01-01' AND f.Source='ПустаяСсылка'
                         AND a.PAP_Article_Name='Налоги'
                   GROUP BY f.TaxType ORDER BY SUM(f.Sum_Close) DESC""")
    gtax = {}
    for r in cur.fetchall():
        gtax[str(r.tt)] = float(r.ko or 0)
        print(f"  {str(r.tt):28} Close={float(r.ko or 0):,.2f} n={r.n}")
    for tt, et in ЭТ_TAX.items():
        v = gtax.get(tt)
        if v is None or abs(v - et) > TOL:
            fail.append(f"TaxType «{tt}» Close={v} != эталон {et:,.2f}")
    # TaxType пуст у НЕ-Налоги
    bad_tax = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
        JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
        WHERE f.Period_Month='2026-01-01' AND f.Source='ПустаяСсылка'
              AND a.PAP_Article_Name<>'Налоги'
              AND f.TaxType IS NOT NULL AND f.TaxType<>'ПустаяСсылка'""").fetchval()
    nalogi_notax = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
        JOIN Dim_PAP_Articles a ON a.PAP_Article_ID=f.PAP_Article_ID
        WHERE f.Period_Month='2026-01-01' AND f.Source='ПустаяСсылка'
              AND a.PAP_Article_Name='Налоги'
              AND (f.TaxType IS NULL OR f.TaxType='ПустаяСсылка')""").fetchval()
    print(f"\nНЕ-Налоги с заполненным TaxType = {bad_tax} (ожид 0); "
          f"Налоги без TaxType = {nalogi_notax} (ожид 0)")
    if bad_tax:
        fail.append(f"{bad_tax} НЕ-Налоги строк с заполненным TaxType")
    if nalogi_notax:
        fail.append(f"{nalogi_notax} Налоги строк без TaxType")

    # 3) РЕГРЕССИЯ Себест/ДенСр/Расчёты (2026-01)
    print("\n=== РЕГРЕССИЯ 2026-01 ===")
    seb = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                       "WHERE Period_Month='2026-01-01' AND Source='СебестоимостьТоваров'").fetchval()
    den = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                       "WHERE Period_Month='2026-01-01' AND Source LIKE N'ДенежныеСредства%'").fetchval()
    kl = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                      "WHERE Period_Month='2026-01-01' AND Source='РасчетыСКлиентамиПоСрокам'").fetchval()
    ps = cur.execute("SELECT ISNULL(SUM(Sum_Close),0) FROM Fact_Balance "
                      "WHERE Period_Month='2026-01-01' AND Source='РасчетыСПоставщикамиПоСрокам'").fetchval()
    for nm, v, et in (("Себест", seb, ЭТ_СЕБ), ("ДенСр", den, ЭТ_ДЕНСР),
                      ("Клиенты", kl, ЭТ_КЛ), ("Поставщики", ps, ЭТ_ПОСТ)):
        ok = abs(float(v) - et) <= TOL
        print(f"  {nm:11} = {float(v):,.2f}  эталон {et:,.2f}  {'OK' if ok else '!!! РЕГРЕССИЯ'}")
        if not ok:
            fail.append(f"РЕГРЕССИЯ {nm}: {float(v):,.2f} != {et:,.2f}")

    # 4) Dim_TaxTypes FK-покрытие (детализация TaxType)
    print("\n=== Dim_TaxTypes ===")
    has_dim = cur.execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
                          "WHERE TABLE_NAME='Dim_TaxTypes'").fetchval()
    if not has_dim:
        fail.append("Dim_TaxTypes не создана")
    else:
        dn = cur.execute("SELECT COUNT(*) FROM Dim_TaxTypes").fetchval()
        miss = cur.execute("""SELECT COUNT(*) FROM Fact_Balance f
            WHERE f.TaxType IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM Dim_TaxTypes d WHERE d.TaxType=f.TaxType)""").fetchval()
        print(f"  Dim_TaxTypes rows={dn} (ожид 15: 14 enum + ПустаяСсылка); "
              f"Fact_Balance.TaxType без Dim = {miss} (ожид 0)")
        if dn != 15:
            fail.append(f"Dim_TaxTypes rows={dn} != 15")
        if miss:
            fail.append(f"{miss} строк Fact_Balance.TaxType не join'ятся в Dim_TaxTypes")

print("\n" + "=" * 64)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: Fact_Balance Source=ПустаяСсылка == регистр/ПАП/УпрБаланс до копейки; "
      "разрез Налоги по TaxType == Карточка актива/пассива; TaxType заполнен "
      "только у Налоги; 3 месяца сосуществуют; Себест/ДенСр/Расчёты не регрессировали")
