# -*- coding: utf-8 -*-
"""Маппинг СтатьиАктивовПассивов -> регистры-расшифровки + заполненные аналитики.
Источник логики: Documents/А_ФинРез_Баланс/Ext/ObjectModule.bsl (свёртка).
Эмпирика: OlapBASERP.Fact_Balance (январь 2026 / ТОВ) JOIN Dim_PAP_Articles."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pyodbc

CX = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
                    "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;")

# Source(А_ИсточникБаланса) -> регистр-расшифровка + какие аналитики даёт
SRC = {
    "ПрочиеАктивыПассивы":          ("РегНак.ПрочиеАктивыПассивы (база, всегда)",
                                     "Аналитика ПАП: Контрагент/Организация, Партнёр, ОбъектыЭксплуатации, ФизЛицо, НематериальныйАктив"),
    "РасчетыСКлиентами":            ("РегНак.РасчетыСКлиентами",        "ОбъектРасчетов"),
    "РасчетыСПоставщиками":         ("РегНак.РасчетыСПоставщиками",     "ОбъектРасчетов"),
    "СебестоимостьТоваров":         ("РегНак.СебестоимостьТоваров",     "Номенклатура, Склад"),
    "ДенежныеСредстваБезналичные":  ("РегНак.ДенежныеСредстваБезналичные", "ДенежныеСредства=БанковскийСчет"),
    "ДенежныеСредстваНаличные":     ("РегНак.ДенежныеСредстваНаличные", "ДенежныеСредства=Касса"),
    "ПрочиеРасходы":                ("РегНак.ВзаиморасчетыССотрудниками (Этап4 «Оплата труда»)",
                                     "ФизическоеЛицо (пропорц. |сальдо|)"),
}

print("=" * 100)
print("1) ПОЛНЫЙ ПЕРЕЧЕНЬ СтатьиАктивовПассивов (Dim_PAP_Articles, full reload из 1С)")
print("=" * 100)
rows = CX.execute("""SELECT PAP_Article_Code, PAP_Article_Name, AktivPassiv, Is_Group,
       (SELECT p.PAP_Article_Name FROM Dim_PAP_Articles p WHERE p.PAP_Article_ID=d.Parent_ID) AS Parent
  FROM Dim_PAP_Articles d
 ORDER BY Is_Group DESC, PAP_Article_Name""").fetchall()
print(f"{'Код':<10} {'Статья':<42} {'АктивПассив':<12} {'Гр':<3} {'Родитель'}")
print("-" * 100)
for r in rows:
    print(f"{(r[0] or ''):<10} {(r[1] or '')[:42]:<42} {(r[2] or '—'):<12} "
          f"{('ГР' if r[3] else ''):<3} {(r[4] or '')}")
print(f"\nВсего строк Dim_PAP_Articles: {len(rows)}")

print("\n" + "=" * 100)
print("2) РАСКЛАДКА: Статья -> Source(регистр-расшифровка) -> заполненные аналитики (январь 2026/ТОВ)")
print("=" * 100)
q = """SELECT d.PAP_Article_Name AS A, d.AktivPassiv AS AP, f.Source AS Src,
       COUNT(*) AS Rws, SUM(f.Sum_Close) AS SC,
       MAX(IIF(f.Counterparty_ID  IS NOT NULL,1,0)) Cp,
       MAX(IIF(f.Partner_ID       IS NOT NULL,1,0)) Pr,
       MAX(IIF(f.Individual_ID    IS NOT NULL,1,0)) Ind,
       MAX(IIF(f.Item_ID          IS NOT NULL,1,0)) It,
       MAX(IIF(f.Warehouse_ID     IS NOT NULL,1,0)) Wh,
       MAX(IIF(f.Cash_ID          IS NOT NULL,1,0)) Csh,
       MAX(IIF(f.SettlementObj_ID IS NOT NULL,1,0)) SO,
       MAX(IIF(f.OperObject_ID    IS NOT NULL,1,0)) OO,
       MAX(IIF(f.Contract_ID      IS NOT NULL,1,0)) Ctr,
       MAX(IIF(f.Intangible_ID    IS NOT NULL,1,0)) Intg
  FROM Fact_Balance f JOIN Dim_PAP_Articles d ON f.PAP_Article_ID=d.PAP_Article_ID
 WHERE f.Period_Month='2026-01-01'
 GROUP BY d.PAP_Article_Name, d.AktivPassiv, f.Source
 ORDER BY d.PAP_Article_Name, f.Source"""
cur = CX.execute(q)
dims = ["Контрагент", "Партнёр", "ФизЛицо", "Номенкл", "Склад", "ДенСр",
        "ОбъектРасч", "ОбъектЭкспл", "Договор", "НМА"]
cur_article = None
for row in cur.fetchall():
    a, ap, src, rws, sc = row[0], row[1], row[2], row[3], float(row[4] or 0)
    flags = row[5:]
    if a != cur_article:
        cur_article = a
        print(f"\n■ {a}  [{ap or '—'}]")
    filled = [dims[i] for i, v in enumerate(flags) if v]
    reg, ana = SRC.get(src, (src, "?"))
    print(f"   {src:<26} → {reg}")
    print(f"      рядків={rws:<5} ΣКонОст={sc:,.2f}  заповнені аналітики: "
          f"{', '.join(filled) if filled else '— (тільки Статья)'}")

print("\n" + "=" * 100)
print("3) ЛЕГЕНДА (з ObjectModule.bsl свёртки, канон v1.4)")
print("=" * 100)
print("• Суми (НачОст/Приход/Расход/КонОст) ЗАВЖДИ з РегНак.ПрочиеАктивыПассивы.ОстаткиИОбороты.")
print("• Source = який типовий регістр має Регистратор цього руху ПАП (fill-when-unique,")
print("  LEFT JOIN по Регистратор; якщо унікальний — підставляється субконто).")
print("• Пріоритет Source: РасчетыСКлиентами > РасчетыСПоставщиками > СебестоимостьТоваров")
print("  > ДенСрБезнал > ДенСрНал > (інакше) ПрочиеАктивыПассивы.")
print("• «Оплата труда» — окремо: розклад по ФизическоеЛицо ∝ |сальдо| ВзаиморасчетыССотрудниками,")
print("  Source=ПрочиеРасходы (Σ ваг=1 → Σ по статті == ПАП).")
print("• Виключені (закритий період, не пишуться): Собственные средства,")
print("  Доходы текущего периода, Расходы текущего периода (та їх ієрархія).")
for s, (reg, ana) in SRC.items():
    print(f"   {s:<26}: {reg}  ⇒ аналітики: {ana}")
CX.close()
