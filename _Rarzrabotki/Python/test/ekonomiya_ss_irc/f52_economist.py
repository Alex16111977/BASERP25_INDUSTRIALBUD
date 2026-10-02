# -*- coding: utf-8 -*-
"""Ф4. «Локальний кошторис матеріали» экономиста (МАТЕРИАЛИ.xlsx, листы «IRS 15»/«IRS 30») против СС:
сопоставление строк по нормализованному наименованию с оригиналом СС (BuhBud КомплектующиеПоМодульно = файл СС)
и с рабочей ТЧ Комплектующие (ERP), разница цены/объёма/суммы на дом. Расчёт по файлам и out/ss_*.json."""
import os, sys, json, re
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from collections import defaultdict

SB = json.load(open(os.path.join(OUT, "ss_buh.json"), encoding="utf-8"))
SE = json.load(open(os.path.join(OUT, "ss_erp.json"), encoding="utf-8"))
wb = load_xlsx(os.path.join(ROOT, r"_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\МАТЕРИАЛИ.xlsx"))

def N(s):
    s = str(s or "").casefold().replace("\n", " ").replace("’", "'").replace("`", "'")
    s = re.sub(r"[\"«»]", "", s)
    return " ".join(s.split())

out = {}
for sheet, t in (("IRS 15", "15"), ("IRS 30", "30")):
    ws = wb[sheet]
    econ = []
    for r in range(5, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name or not isinstance(ws.cell(r, 4).value, (int, float)):
            continue
        econ.append({"row": r, "name": " ".join(str(name).split()), "unit": ws.cell(r, 3).value, "q": D(ws.cell(r, 4).value),
                     "p": D(ws.cell(r, 7).value or 0), "s": D(ws.cell(r, 8).value or 0)})
    orig = [x for x in SB["orig"] if (" 15 м" in x["СС"]) == (t == "15")]
    kom = [x for x in SE["kom"] if (" 15 м" in x["СС"]) == (t == "15")]
    oi = defaultdict(list)
    for x in orig: oi[N(x["НомСС"])].append(x)
    ki = defaultdict(list)
    for x in kom: ki[N(x["НомСС"])].append(x)
    used = set(); rows = []
    for e in econ:
        k = N(e["name"])
        o = oi.get(k, [])
        kk = ki.get(k, [])
        oq = sum(D(x["КолОриг"] or x["Кол"]) for x in o); os_ = sum(D(x["Сумма"]) for x in o)
        ks = sum(D(x["Сумма"]) for x in kk); kq = sum(D(x["Кол"]) for x in kk)
        on = (kk[0]["ОН"] if kk else (o[0]["ОН"] if o else None))
        if o: used.add(k)
        rows.append({**{f: (str(v) if isinstance(v, Decimal) else v) for f, v in e.items()},
                     "match": "оригинал СС" if o else ("Комплектующие" if kk else "нет в СС"),
                     "ОН": on, "orig_q": str(oq), "orig_unit": (o[0]["ЕдОриг"] or o[0]["ЕдСС"]) if o else None, "orig_s": str(os_),
                     "kom_q": str(kq), "kom_s": str(ks), "ds": str(e["s"] - (os_ if o else ks))})
    miss = [{"НомСС": x["НомСС"], "Сумма": x["Сумма"], "ОН": x["ОН"]} for x in orig if N(x["НомСС"]) not in used]
    se = sum(D(x["s"]) for x in econ); so = sum(D(x["Сумма"]) for x in orig); sk = sum(D(x["Сумма"]) for x in kom)
    matched = [x for x in rows if x["match"] == "оригинал СС"]
    dpos = sum(D(x["ds"]) for x in matched if D(x["ds"]) > 0); dneg = sum(D(x["ds"]) for x in matched if D(x["ds"]) < 0)
    print(f"{sheet}: строк экономиста {len(econ)}, Σ на дом {fmt(se)}; оригинал СС {len(orig)} строк Σ {fmt(so)}; Комплектующие Σ {fmt(sk)}")
    print(f"   сопоставлено с оригиналом: {len(matched)}; по Комплектующим: {sum(1 for x in rows if x['match']=='Комплектующие')};"
          f" нет в СС: {sum(1 for x in rows if x['match']=='нет в СС')}; строк оригинала без пары: {len(miss)}")
    print(f"   по сопоставленным: экономист дороже на {fmt(dpos)}, дешевле на {fmt(dneg)}")
    for x in sorted(matched, key=lambda x: D(x["ds"]))[:8]:
        print(f"      {x['name'][:45]:<45} эк {fmt(x['s'])} / СС {fmt(x['orig_s'])} Δ {fmt(x['ds'])} (кол {x['q']} vs {x['orig_q']}, цена {x['p']})")
    for x in [x for x in rows if x["match"] != "оригинал СС"][:10]:
        print("      НЕ СОПОСТАВЛЕНО:", x["name"][:60], x["s"], x["match"])
    for m in miss[:10]:
        print("      В СС НЕТ У ЭКОНОМИСТА:", m)
    out[sheet] = {"rows": rows, "miss": miss, "sum_econ": str(se), "sum_orig": str(so), "sum_kom": str(sk)}
dump_json(out, "economist_vs_ss.json")
