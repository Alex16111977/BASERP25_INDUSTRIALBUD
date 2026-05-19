# -*- coding: utf-8 -*-
"""READ-ONLY верификация расширения Dim_Contracts/Dim_ObjektyRaschetov +
Dim_TipyDogovorov/Dim_FinAgents. Проверяет: новые dim непусты; новые колонки
заполнены (% NOT NULL по непустым ссылкам); FK 0 orphans
TipDogovora→Dim_TipyDogovorov, FinAgent_ID→Dim_FinAgents; row-count не
регрессировал; enum-колонки = кириллица (не hex)."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

fail = []
with get_olap_sql() as c:
    cur = c.cursor()

    tc = cur.execute("SELECT COUNT(*) FROM Dim_Contracts").fetchval()
    to = cur.execute("SELECT COUNT(*) FROM Dim_ObjektyRaschetov").fetchval()
    td = cur.execute("SELECT COUNT(*) FROM Dim_TipyDogovorov").fetchval()
    tf = cur.execute("SELECT COUNT(*) FROM Dim_FinAgents").fetchval()
    print(f"rows: Dim_Contracts={tc} Dim_ObjektyRaschetov={to} "
          f"Dim_TipyDogovorov={td} Dim_FinAgents={tf}")
    if tc < 8000:
        fail.append(f"Dim_Contracts регрессировал rows={tc} (<8000)")
    if to < 14000:
        fail.append(f"Dim_ObjektyRaschetov регрессировал rows={to} (<14000)")
    if td < 2:
        fail.append(f"Dim_TipyDogovorov пуст rows={td}")
    if tf < 1:
        fail.append(f"Dim_FinAgents пуст rows={tf}")

    orf = cur.execute("""SELECT COUNT(*) FROM Dim_Contracts d
        WHERE d.FinAgent_ID IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_FinAgents f
                          WHERE f.FinAgent_ID = d.FinAgent_ID)""").fetchval()
    ort = cur.execute("""SELECT COUNT(*) FROM Dim_Contracts d
        WHERE d.TipDogovora IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM Dim_TipyDogovorov t
                          WHERE t.TipDogovora = d.TipDogovora)""").fetchval()
    print(f"FK orphans: FinAgent_ID={orf} (ожид 0); TipDogovora={ort} (ожид 0)")
    if orf:
        fail.append(f"{orf} Dim_Contracts.FinAgent_ID без Dim_FinAgents")
    if ort:
        fail.append(f"{ort} Dim_Contracts.TipDogovora без Dim_TipyDogovorov")

    badenum = cur.execute("""SELECT COUNT(*) FROM Dim_Contracts
        WHERE TipDogovora IS NOT NULL AND LEN(TipDogovora)=32
          AND TipDogovora NOT LIKE '%[^0-9a-f]%'""").fetchval()
    print(f"Dim_Contracts.TipDogovora hex-подобных (ожид 0)={badenum}")
    if badenum:
        fail.append(f"{badenum} TipDogovora остались hex (enum_resolver не сработал)")

    nm = cur.execute("""SELECT
        SUM(CASE WHEN Department_Name IS NOT NULL THEN 1 ELSE 0 END) dep,
        SUM(CASE WHEN Counterparty_Name IS NOT NULL THEN 1 ELSE 0 END) cp,
        SUM(CASE WHEN Org_Buh_Name IS NOT NULL THEN 1 ELSE 0 END) ob
        FROM Dim_Contracts""").fetchone()
    print(f"Dim_Contracts NOT NULL: Department_Name={nm.dep} "
          f"Counterparty_Name={nm.cp} Org_Buh_Name={nm.ob}")
    if (nm.dep or 0) == 0 and (nm.cp or 0) == 0 and (nm.ob or 0) == 0:
        fail.append("все денорм-имена Dim_Contracts NULL (JOIN не сработал)")

    oo = cur.execute("""SELECT
        SUM(CASE WHEN Object_Type_Name IS NOT NULL THEN 1 ELSE 0 END) ot,
        SUM(CASE WHEN TipRaschetov IS NOT NULL THEN 1 ELSE 0 END) tr
        FROM Dim_ObjektyRaschetov""").fetchone()
    print(f"Dim_ObjektyRaschetov NOT NULL: Object_Type_Name={oo.ot} "
          f"TipRaschetov={oo.tr}")
    if (oo.tr or 0) == 0:
        fail.append("TipRaschetov весь NULL (enum/JOIN не сработал)")

print("\n" + "=" * 64)
if fail:
    for f in fail:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: Dim_Contracts/Dim_ObjektyRaschetov расширены; "
      "Dim_TipyDogovorov/Dim_FinAgents непусты; FK 0 orphans; "
      "enum=кириллица; денорм-имена заполнены")
