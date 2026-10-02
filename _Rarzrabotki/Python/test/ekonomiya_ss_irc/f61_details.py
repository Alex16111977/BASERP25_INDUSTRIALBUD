# -*- coding: utf-8 -*-
"""Ф5. Списки позиций по корзинам для ANALIZ.md: частичные нули, не использовано, поза кошторисом, единица/состав,
несопоставимые, ТОП-80 %, НДС-артефакт по 8 домам, поставщики без НДС. Расчёт по out/*.json → out/details.txt."""
import os, sys, json, re
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from collections import defaultdict

L = lambda n: json.load(open(os.path.join(OUT, n), encoding="utf-8"))
A = L("analysis.json")["rows"]; C = L("core.json"); F = L("final.json"); PU = L("purch_erp.json"); MV = L("buh_moves.json"); RC = L("buh_receipts.json")
ON = {k: defaultdict(Decimal, {m: D(v) for m, v in d.items()}) for k, d in C["on"].items()}
H15 = [f"15-{i}" for i in range(1, 7)]; H30 = ["30-1", "30-2"]
out = []
w = out.append
dv = lambda a, b: a / b if b else Decimal(0)

cls_ = classify_row

for g, hs, plan_h in (("d15", H15, "15-1"), ("d30", H30, "30-1")):
    w(f"===== {g}")
    for cat in ("не использовано", "поза кошторисом", "единица/состав", "несопоставимо"):
        items = sorted([r for r in A.values() if cls_(r, g) == cat], key=lambda r: D(r[f"{g}_delta"]))
        w(f"-- {cat}: {len(items)} назв, Σ Δ {fmt(sum(D(r[g + '_delta']) for r in items))}")
        for r in items[:12]:
            o = ON[r["k"]]
            w(f"     {r['name'][:45]:<45} Δ {fmt(r[g + '_delta']):>12}  норма/дім {fmt(o['nr|'+plan_h+'|ПланКол'],3)} на {fmt(o['nr|'+plan_h+'|ПланГрн'])}  "
              f"списано {fmt(r[g + '_cq'],3)} {r['units']}")
    w(f"-- частичные нули ({g}): дома с нулевым списанием при норме и списании на других")
    for r in A.values():
        if cls_(r, g) != "сопоставимо": continue
        o = ON[r["k"]]
        zh = [h for h in hs if o[f"c|{h}|q"] == 0]
        if zh and len(zh) < len(hs):
            nq = o[f"nr|{plan_h}|ПланКол"]; ns = o[f"nr|{plan_h}|ПланГрн"]
            w(f"     {r['name'][:45]:<45} нулевые: {','.join(zh)}  норма/дім {fmt(nq,3)} × ціна {fmt(dv(ns,nq))} = {fmt(len(zh)*ns)}")

w("===== ТОП-80% экономии 15 м")
for n, d, p in F["top15"]: w(f"     {n[:45]:<45} {fmt(d):>12}  {fmt(p,1)} %")
w("===== ТОП-80% экономии в отчёте директору (расчёт по его формулам)")
for n, d, p in F["topdir"]: w(f"     {n[:45]:<45} {fmt(d):>12}  {fmt(p,1)} %")
w("===== ТОП «ETC − дефицит по остаткам»")
for n, d, p in F["topetc"]: w(f"     {n[:45]:<45} {fmt(d):>12}  {fmt(p,1)} %")
w("===== ТОП перерасходов по отчёту директору")
for r in sorted(A.values(), key=lambda r: -D(r["dev"]))[:12]:
    w(f"     {r['name'][:45]:<45} откл {fmt(r['dev']):>12} цен {fmt(r['dir_pf']):>12} кол {fmt(r['dir_qf']):>12} кат {r['dir_cat']}")
w("===== ТОП экономий по отчёту директору")
for r in sorted(A.values(), key=lambda r: D(r["dev"]))[:12]:
    w(f"     {r['name'][:45]:<45} откл {fmt(r['dev']):>12} цен {fmt(r['dir_pf']):>12} кол {fmt(r['dir_qf']):>12} ETC {fmt(r['etc']):>12}")

# поставщики без НДС и НДС-артефакт по 8 домам
nov = defaultdict(Decimal)
for r in PU:
    if D(r["НДС"]) == 0: nov[r["Контрагент"]] += D(r["Стоимость"])
w(f"===== Закупки без НДС: Σ {fmt(sum(nov.values()))}, поставщиков {len(nov)}")
for k, v in sorted(nov.items(), key=lambda x: -x[1]): w(f"     {k[:45]:<45} {fmt(v)}")
mis = defaultdict(lambda: [0, Decimal(0), Decimal(0)])
for r in RC:
    key = (r["СтавкаНом"], r["СтавкаСтр"]); mis[key][0] += 1; mis[key][1] += D(r["Нетто"]); mis[key][2] += D(r["НДС"])
w("===== Строки приходов BuhBud: (ставка номенклатуры, ставка строки) → строк, нетто, НДС")
for k, v in mis.items(): w(f"     {k}: {v[0]} строк, нетто {fmt(v[1])}, НДС {fmt(v[2])}")
tr = {h: sum(o[f"c|{h}|gt"] for o in ON.values()) for h in H15 + H30}
gr = {h: sum(o[f"c|{h}|gr"] for o in ON.values()) for h in H15 + H30}
rep = {h: D(F["house"][h]["ФактВНормеГрн"]) + D(F["house"][h]["ФактПонадГрн"]) for h in H15 + H30}
w("===== НДС-база 8 домов: отчёт (як план) / моя реконструкция як план / реальный брутто")
for h in H15 + H30: w(f"     {h}: {fmt(rep[h])} / {fmt(gr[h])} / {fmt(tr[h])}  Δреальн−отчёт {fmt(tr[h]-rep[h])}")
w(f"     ИТОГО: {fmt(sum(rep.values()))} / {fmt(sum(gr.values()))} / {fmt(sum(tr.values()))}  Δ {fmt(sum(tr.values())-sum(rep.values()))}"
  f"  из них ставка: {fmt(sum(tr.values())-sum(gr.values()))}, реконструкция: {fmt(sum(gr.values())-sum(rep.values()))}")
p = os.path.join(OUT, "details.txt")
open(p, "w", encoding="utf-8").write("\n".join(out))
print(p, len(out))
