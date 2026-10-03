# -*- coding: utf-8 -*-
"""Неттинг по позициям (метод потоков M05 3.5) для незавершённых домов IRC 30 м №3–№7 по кэшу исследования
(core.json / analysis.json). Цена: (а) СС, (б) замещения (факт: склад проекта → расход 30 м → расход 15 м).
Помесячно по ship_real 00009: №3, №4 — отгрузка 10; №5, №6 — 11; №7 — 12; оплата за 1 мес. до отгрузки
(сентябрь → октябрь). Только расчёт по файлам, базы не трогает."""
import json, os, sys
from decimal import Decimal as D
from collections import defaultdict

sys.path.insert(0, r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\ekonomiya_ss_irc")
from common import fmt, classify_row

SRC = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\ekonomiya_ss_irc\out"
C = json.load(open(os.path.join(SRC, "core.json"), encoding="utf-8"))
A = json.load(open(os.path.join(SRC, "analysis.json"), encoding="utf-8"))["rows"]
ON = {k: defaultdict(lambda: D(0), {m: D(str(v)) for m, v in d.items()}) for k, d in C["on"].items()}
NAME = C["names"]
Z0 = D(0)
H15 = [f"15-{i}" for i in range(1, 7)]
MONTH = {"30-3": "2026-10", "30-4": "2026-10", "30-5": "2026-10", "30-6": "2026-10", "30-7": "2026-11"}

def dv(a, b):
    return a / b if b else Z0

tot = defaultdict(lambda: D(0))
items = []
for k, o in ON.items():
    a = A.get(k, {})
    nq, ns = o["nr|30-1|ПланКол"], o["nr|30-1|ПланГрн"]
    pss = dv(ns, nq)
    p_stock = dv(o["stk|P|gt"], o["stk|P|q"])
    q30 = o["c|30-1|q"] + o["c|30-2|q"]; g30 = o["c|30-1|gt"] + o["c|30-2|gt"]
    q15 = sum(o[f"c|{h}|q"] for h in H15); g15 = sum(o[f"c|{h}|gt"] for h in H15)
    p_c30 = dv(g30, q30); p_c15 = dv(g15, q15)
    p_rep, src = (p_stock, "склад") if p_stock > 0 else ((p_c30, "расход30") if p_c30 > 0 else ((p_c15, "расход15") if p_c15 > 0 else (pss, "СС")))
    flag = ""
    if pss > 0 and p_rep > 0:
        ratio = p_rep / pss
        if ratio < D("0.5") or ratio > D("2"):
            flag = f"цена×{ratio:.2f} — единица/состав, взята СС"
            p_rep = pss; src = "СС*"
    cat30 = a.get("d30_cat", "?"); cat15 = a.get("d15_cat", "?")
    try:
        ccat30 = classify_row({kk: (D(str(vv)) if kk.startswith("d30_") and kk[4:] in ("cq", "Nq", "cg", "Ns") else vv) for kk, vv in a.items()}, "d30") if a else "?"
    except Exception:
        ccat30 = cat30
    oos = ns <= 0 and q30 > 0                      # поза кошторисом на 30 м: «норма» = средний расход 30 м №1–№2
    if oos:
        nq_eff = q30 / 2; price_ss = dv(g30, q30)
    else:
        nq_eff = nq; price_ss = pss
    rec = {"k": k, "name": NAME.get(k, k), "nq": nq_eff, "pss": price_ss, "p_rep": (dv(g30, q30) if oos else p_rep),
           "src": src, "flag": flag, "cat30": ccat30, "cat15": cat15, "oos": oos, "def": {}, "def_rep": {}}
    P_q, P_g = o["stk|P|q"], o["stk|P|gt"]
    own_q = o["stkh|30-3|q"]
    if nq_eff <= 0 and ns > 0:                    # план суммой без количества
        d3 = max(Z0, ns - o["c|30-3|gt"] - o["stkh|30-3|gt"])
        rec["def"] = {"30-3": d3, **{f"30-{j}": ns for j in range(4, 8)}}
        rec["def_rep"] = dict(rec["def"])
    elif nq_eff > 0:
        need = max(Z0, nq_eff - o["c|30-3|q"])
        f_own = min(need, max(Z0, own_q)); need -= f_own
        f_P = min(need, max(Z0, P_q)); need -= f_P
        P_rem = max(Z0, P_q) - f_P
        qd = {"30-3": need}
        for j in range(4, 8):
            take = min(nq_eff, P_rem); P_rem -= take
            qd[f"30-{j}"] = nq_eff - take
        rec["def"] = {h: v * price_ss for h, v in qd.items()}
        rec["def_rep"] = {h: v * rec["p_rep"] for h, v in qd.items()}
        rec["surplus_q"] = P_rem
        rec["surplus_g"] = P_rem * p_stock if p_stock else Z0
    for h, v in rec["def"].items():
        tot[("ss", "oos" if oos else "norm", MONTH[h])] += v
    for h, v in rec["def_rep"].items():
        tot[("rep", "oos" if oos else "norm", MONTH[h])] += v
    rec["sum_ss"] = sum(rec["def"].values(), Z0); rec["sum_rep"] = sum(rec["def_rep"].values(), Z0)
    items.append(rec)

print("== Дефицит незавершённых 30 м №3–№7 по позициям (с НДС) ==")
for base in ("ss", "rep"):
    for kind in ("norm", "oos"):
        o10, o11 = tot[(base, kind, "2026-10")], tot[(base, kind, "2026-11")]
        print(f"  {base:<4} {kind:<5} окт {fmt(o10):>14}  ноя {fmt(o11):>14}  итого {fmt(o10 + o11):>14}")
by = defaultdict(lambda: [D(0), D(0), 0])
for r in items:
    if r["oos"]:
        continue
    c = r["cat30"]
    by[c][0] += r["sum_ss"]; by[c][1] += r["sum_rep"]; by[c][2] += 1 if r["sum_ss"] else 0
print("== Дефицит (норм. позиции) по категории расхода 30 м №1–№2: СС / замещения / позиций ==")
for c, v in sorted(by.items(), key=lambda x: -x[1][0]):
    print(f"  {c:<22} {fmt(v[0]):>14} {fmt(v[1]):>14} {v[2]}")
by15 = defaultdict(lambda: [D(0), D(0)])
for r in items:
    if not r["oos"] and r["cat30"] == "не использовано":
        by15[r["cat15"]][0] += r["sum_ss"]; by15[r["cat15"]][1] += r["sum_rep"]
print("   из «не использовано на 30 м»: по категории на 15 м:", {k: (fmt(v[0]), fmt(v[1])) for k, v in by15.items()})
print("== ТОП-25 дефицита по СС ==")
for r in sorted(items, key=lambda r: -r["sum_ss"])[:25]:
    print(f"  {r['name'][:34]:<34} {('ВНЕ' if r['oos'] else ''):<3} {r['cat30'][:14]:<14} СС {fmt(r['sum_ss']):>12} зам {fmt(r['sum_rep']):>12} [{r['src']}] {r['flag']}")
print("== Излишек склада проекта после 30 м №3–№7 (по цене склада) ==", fmt(sum(r.get("surplus_g", Z0) for r in items)))
print("   из них позиции вне сметы:", fmt(sum(r.get("surplus_g", Z0) for r in items if r["oos"] or (r["nq"] <= 0))))
json.dump([{**{kk: (str(vv) if isinstance(vv, D) else vv) for kk, vv in r.items() if kk not in ("def", "def_rep")},
            "def": {h: str(v) for h, v in r["def"].items()}, "def_rep": {h: str(v) for h, v in r["def_rep"].items()}}
           for r in items], open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "v03.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=0)
