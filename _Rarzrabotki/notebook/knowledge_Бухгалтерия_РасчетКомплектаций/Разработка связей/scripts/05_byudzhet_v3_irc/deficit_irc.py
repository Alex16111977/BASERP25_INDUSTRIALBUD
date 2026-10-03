# -*- coding: utf-8 -*-
"""Расчёт по кэшу исследования (out/core.json, analysis.json) — ничего не пишет в базы.
A) нестача к норме СС 30 м №3–№7 по остаткам (как исследование) + разбивка по домам №4..№7 + оценка по ценам запаса/закупки;
B) ГИПОТЕЗА: 30 м №3–№7 расходуют как 30 м №1–№2 (списано + лежит на их складах), докупка по ценам запаса/закупки."""
import sys, json
sys.path.insert(0, r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\ekonomiya_ss_irc")
from common import D, fmt, OUT, classify_row
import os
from decimal import Decimal
Z = Decimal(0)
C = json.load(open(os.path.join(OUT, "core.json"), encoding="utf-8"))
A = json.load(open(os.path.join(OUT, "analysis.json"), encoding="utf-8"))["rows"]
def isnum(s):
    try: Decimal(s); return True
    except Exception: return False
R = {k: {f: (D(v) if isinstance(v, str) and isnum(v) and f not in ("k", "name") else v) for f, v in r.items()} for k, r in A.items()}
ON = {k: {m: D(v) for m, v in d.items()} for k, d in C["on"].items()}
g = lambda o, m: o.get(m, Z)
dv = lambda a, b: a / b if b else Z

tot = {k: Z for k in ("A3", "A4", "A5", "A6", "A7", "A_ss", "A_px", "A_unused30", "A_unused30_px", "B_def_px", "B_def_ss", "B_cover_P",
                      "A_cover_P", "B_P_left", "A_noqty", "pz_def", "inc_def")}
rowsA = []
for k, o in ON.items():
    r = R[k]; name = C["names"][k]; units = C["units"][k]
    nq, ns = g(o, "nr|30-1|ПланКол"), g(o, "nr|30-1|ПланГрн")
    pss = dv(ns, nq)
    Pq, Pg = g(o, "stk|P|q"), g(o, "stk|P|gt")
    own_q = g(o, "stkh|30-3|q"); c3 = g(o, "c|30-3|q")
    p_stock = dv(Pg, Pq) if Pq > 0 else Z
    pu = dv(g(o, "pu_g"), g(o, "pu_q"))
    ub = [u.split(":", 1)[1] for u in units if u.startswith("BUH:")]; ue = [u.split(":", 1)[1] for u in units if u.startswith("ERP:")]
    same_unit = ub and ue and ub[0] == ue[0]
    def px(fallback):
        if p_stock > 0 and (fallback == 0 or D("0.2") <= p_stock / fallback <= 5): return p_stock
        if pu > 0 and same_unit and fallback > 0 and D("0.5") <= pu / fallback <= 2: return pu
        if pu > 0 and same_unit and fallback == 0: return pu
        return fallback
    # ---------- A: норма (как исследование f41 Е/Ж)
    if nq > 0:
        need3 = max(Z, nq - c3)
        f_own = min(need3, max(Z, own_q)); need3b = need3 - f_own
        f_P = min(need3b, max(Z, Pq)); d3q = need3b - f_P
        Prem = max(Z, Pq - f_P)
        hq = []
        for i in range(4):
            take = min(nq, max(Z, Prem - nq * i)); hq.append(nq - take)
        tot["A3"] += d3q * pss
        for i, h in enumerate(("A4", "A5", "A6", "A7")): tot[h] += hq[i] * pss
        dq = d3q + sum(hq)
        tot["A_ss"] += dq * pss; tot["A_px"] += dq * px(pss)
        c30 = classify_row(r, "d30")
        if c30 == "не использовано":
            tot["A_unused30"] += dq * pss; tot["A_unused30_px"] += dq * px(pss)
        if c30 in ("несопоставимо", "единица/состав"):
            tot["inc_def"] += dq * pss
        zero = sum(1 for h in ("30-1", "30-2") if g(o, f"c|{h}|q") == 0)
        if c30 == "сопоставимо" and zero == 1:
            tot["pz_def"] += dq * pss
        rowsA.append((name, dq * pss, dq * px(pss), c30, [x * pss for x in [d3q] + hq]))
    elif ns > 0:   # план суммой без количества
        d = max(Z, ns - g(o, "c|30-3|gt")) + ns * 4
        tot["A_noqty"] += d; tot["A_ss"] += d; tot["A_px"] += d
        rowsA.append((name, d, d, "план без к-сти", []))
    # ---------- B: ГИПОТЕЗА расход как 30 м №1–№2 (списано + остаток их складов), на дом — среднее двух
    rate = (g(o, "c|30-1|q") + g(o, "stkh|30-1|q") + g(o, "c|30-2|q") + g(o, "stkh|30-2|q")) / 2
    rate_g = (g(o, "c|30-1|gt") + g(o, "stkh|30-1|gt") + g(o, "c|30-2|gt") + g(o, "stkh|30-2|gt")) / 2
    p_rate = dv(rate_g, rate)
    if rate > 0:
        needB3 = max(Z, rate - c3 - max(Z, own_q))
        fP3 = min(needB3, max(Z, Pq)); dB3 = needB3 - fP3
        PremB = max(Z, Pq - fP3)
        fP47 = min(rate * 4, PremB); dB47 = rate * 4 - fP47
        dq = dB3 + dB47
        pr = px(p_rate if p_rate > 0 else pss)
        tot["B_def_px"] += dq * pr
        tot["B_def_ss"] += dq * (pss if pss > 0 else pr)
        tot["B_cover_P"] += (fP3 + fP47) * p_stock
        tot["B_P_left"] += max(Z, Pq - fP3 - fP47) * p_stock
    else:
        tot["B_P_left"] += max(Z, Pq) * p_stock

for k_, v in tot.items(): print(f"{k_:<16} {fmt(v)}")
print("A контроль: 30м№3 + №4..№7 =", fmt(tot["A3"] + tot["A4"] + tot["A5"] + tot["A6"] + tot["A7"] + tot["A_noqty"]))
print("ТОП-15 нестачи A (по цене СС):")
for nm, s, s2, c, hs in sorted(rowsA, key=lambda x: -x[1])[:15]:
    print(f"  {nm[:38]:<38} {fmt(s):>14} по ценам запаса/закупки {fmt(s2):>14}  кат.30м={c}  по домам 3..7: " + " / ".join(fmt(x, 0) for x in hs))
print("Позиции нестачи A с категорией 30 м «не использовано»:")
for nm, s, s2, c, hs in sorted([x for x in rowsA if x[3] == "не использовано" and x[1] > 0], key=lambda x: -x[1]):
    print(f"  {nm[:38]:<38} {fmt(s):>12}")
