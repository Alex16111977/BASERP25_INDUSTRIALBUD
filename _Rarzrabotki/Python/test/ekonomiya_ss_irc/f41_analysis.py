# -*- coding: utf-8 -*-
"""Ф1–Ф3. Расчёты по core.json: (А) факторы отчёта директору (закупки vs СС 13 домов: ценовой + сверх-закупка,
ETC по правилам 04_plan_na_fakt_eac), (Б) факторы по РАСХОДУ на завершённых домах 15 м №1–№6 (и 30 м №1–№2 отдельно),
(В) мост по названию для 9 домов в границах, (Г) потребность незавершённых домов против остатков.
Все числа — «расчёт» по выгрузкам; итоги отчётов сверяются, а не подменяются. Выход: out/analysis.json."""
import os, sys, json
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from collections import defaultdict

C = json.load(open(os.path.join(OUT, "core.json"), encoding="utf-8"))
ON = {k: defaultdict(Decimal, {m: D(v) for m, v in d.items()}) for k, d in C["on"].items()}
NAME = C["names"]; UNIT = C["units"]
H15 = [f"15-{i}" for i in range(1, 7)]
H30D = ["30-1", "30-2"]
SCOPE = H15 + ["30-1", "30-2"]   # границы заказчика 02.10.2026
Z = Decimal(0)

def dv(a, b):
    return a / b if b else Z

res = {}
tot = defaultdict(Decimal)
for k, o in ON.items():
    r = {"k": k, "name": NAME.get(k, k), "units": UNIT.get(k, [])}
    # --- норма
    for f in ("n15_q", "n15_s", "n30_q", "n30_s", "n9_q", "n9_s", "n13_q", "n13_s", "pu_q", "pu_g", "pu_n", "pu_pmin", "pu_pmax"):
        r[f] = o[f]
    p_ss = dv(o["n13_s"], o["n13_q"]); r["p_ss"] = p_ss
    r["p15"] = dv(o["n15_s"], o["n15_q"]); r["p30"] = dv(o["n30_s"], o["n30_q"])
    p_pu = dv(o["pu_g"], o["pu_q"]); r["p_pu"] = p_pu
    r["incomparable"] = bool(p_ss and p_pu and (p_pu / p_ss >= 10 or p_pu / p_ss <= D("0.1")))
    # --- (А) отчёт директору: ETC и факторы
    plan_s, plan_q, f_s, f_q = o["n13_s"], o["n13_q"], o["pu_g"], o["pu_q"]
    if plan_s <= 0:
        etc = Z
    elif plan_q <= 0:
        etc = max(Z, plan_s - f_s)
    elif f_q <= 0:
        etc = plan_s
    else:
        etc = max(Z, plan_q - f_q) * plan_s / plan_q
    dev = f_s + etc - plan_s
    if plan_s > 0 and plan_q > 0 and f_q > 0:
        pf = f_q * (p_pu - p_ss); qf = max(Z, f_q - plan_q) * p_ss
        resid = dev - pf - qf
    else:
        pf = Z; qf = dev; resid = Z      # поза кошторисом / без закупок / план без к-сті
    r.update(etc=etc, dev=dev, dir_pf=pf, dir_qf=qf, dir_resid=resid,
             dir_cat=("поза кошторисом" if plan_s <= 0 else ("план без к-сті" if plan_q <= 0 else ("не закуплено" if f_q <= 0 else "закуплено"))))
    # --- (Б) расход на домах
    # план для расхода — в группировке отчёта по нормам (карточная назва при заполненной номенклатуре), на один дом
    for grp, houses, nq, ns in (("d15", H15, o["nr|15-1|ПланКол"], o["nr|15-1|ПланГрн"]),
                                ("d30", H30D, o["nr|30-1|ПланКол"], o["nr|30-1|ПланГрн"])):
        cq = [o[f"c|{h}|q"] for h in houses]
        cg = sum(o[f"c|{h}|gt"] for h in houses); cr = sum(o[f"c|{h}|gr"] for h in houses); cn = sum(o[f"c|{h}|n"] for h in houses)
        Cq = sum(cq); N = len(houses); Nq = nq * N; Ns = ns * N
        r[f"{grp}_cq"] = Cq; r[f"{grp}_cg"] = cg; r[f"{grp}_cr"] = cr; r[f"{grp}_cn"] = cn
        r[f"{grp}_avg"] = Cq / N; r[f"{grp}_min"] = min(cq); r[f"{grp}_max"] = max(cq)
        r[f"{grp}_Nq"] = Nq; r[f"{grp}_Ns"] = Ns
        delta = cg - Ns
        pss = dv(ns, nq); pc = dv(cg, Cq)
        inc = bool(pss and pc and (pc / pss >= 10 or pc / pss <= D("0.1")))
        if ns > 0 and nq > 0 and Cq > 0 and not inc:
            pe = Cq * (pc - pss); qe = (Cq - Nq) * pss; cat = "сопоставимо"
        elif ns > 0 and nq > 0 and Cq == 0:
            pe = Z; qe = -Ns; cat = "не использовано"
        elif ns <= 0 and cg > 0:
            pe = Z; qe = cg; cat = "поза кошторисом"
        elif inc:
            pe = Z; qe = Z; cat = "несопоставимо"
        else:
            pe = Z; qe = delta; cat = "план без к-сті" if ns > 0 else "—"
        r[f"{grp}_delta"] = delta; r[f"{grp}_pe"] = pe; r[f"{grp}_qe"] = qe; r[f"{grp}_cat"] = cat
        r[f"{grp}_ne"] = delta - pe - qe          # несопоставимое (должно быть = delta только для cat «несопоставимо»)
        r[f"{grp}_dq_pct"] = dv(Cq - Nq, Nq) * 100 if Nq else None
        r[f"{grp}_dp_pct"] = dv(pc - pss, pss) * 100 if (pss and Cq) else None
    # --- (В) мост: 9 домов
    pin = sum(o[m] for m in o if m.startswith("in|") and m.endswith(">P|n"))
    pin_q = sum(o[m] for m in o if m.startswith("in|") and m.endswith(">P|q"))
    r["b_in_P_n"] = pin; r["b_in_P_q"] = pin_q
    r["b_in_MSP_n"] = o["in|631>MSP|n"]; r["b_in_X_n"] = o["in|631>X|n"]
    def mv(a_pred, b_pred, unit="n"):
        s = Z
        for m, v in o.items():
            if m.startswith("mv|") and m.endswith("|" + unit):
                a, b = m.split("|")[1].split(">")
                if a_pred(a) and b_pred(b): s += v
        return s
    isH = lambda x: x.startswith("H:"); isP = lambda x: x == "P"; isX = lambda x: x == "X"; isM = lambda x: x == "MSP"
    for u in ("n", "q"):
        r[f"b_P_H_{u}"] = mv(isP, isH, u); r[f"b_H_P_{u}"] = mv(isH, isP, u)
        r[f"b_X_H_{u}"] = mv(isX, isH, u); r[f"b_M_H_{u}"] = mv(isM, isH, u)
        r[f"b_X_P_{u}"] = mv(isX, isP, u); r[f"b_P_X_{u}"] = mv(isP, isX, u); r[f"b_P_M_{u}"] = mv(isP, isM, u)
        r[f"b_M_X_{u}"] = mv(isM, isX, u)
        r[f"b_cons_{u}"] = sum(o[f"c|{h}|{u}"] for h in SCOPE)
        r[f"b_outP_{u}"] = sum(o[m] for m in o if m.startswith("out|P|") and m.endswith("|" + u))
        r[f"b_stockP_{u}"] = r[f"b_in_P_{u}"] + r[f"b_X_P_{u}"] + r[f"b_H_P_{u}"] - r[f"b_P_H_{u}"] - r[f"b_P_X_{u}"] - r[f"b_P_M_{u}"] - r[f"b_outP_{u}"]
        r[f"b_stockH_{u}"] = r[f"b_P_H_{u}"] + r[f"b_X_H_{u}"] + r[f"b_M_H_{u}"] - r[f"b_H_P_{u}"] - r[f"b_cons_{u}"]
    r["b_cons_gt"] = sum(o[f"c|{h}|gt"] for h in SCOPE); r["b_cons_gr"] = sum(o[f"c|{h}|gr"] for h in SCOPE)
    # отчёт по нормам: в нормі/понад по домам (как в отчёте, з ПДВ «як план»)
    r["nr_in_g"] = sum(o[f"nr|{h}|ФактВНормеГрн"] for h in SCOPE); r["nr_over_g"] = sum(o[f"nr|{h}|ФактПонадГрн"] for h in SCOPE)
    r["nr_in_q"] = sum(o[f"nr|{h}|ФактВНормеКол"] for h in SCOPE); r["nr_over_q"] = sum(o[f"nr|{h}|ФактПонадКол"] for h in SCOPE)
    r["nr_econ_g"] = sum(o[f"nr|{h}|ЭкономияГрн"] for h in SCOPE); r["nr_econ15_g"] = sum(o[f"nr|{h}|ЭкономияГрн"] for h in H15)
    r["nr_plan_g"] = sum(o[f"nr|{h}|ПланГрн"] for h in SCOPE)
    r["nr_hstock_g"] = sum(o[f"nr|{h}|ОстатокОбъектаГрн"] for h in SCOPE); r["nr_hstock_q"] = sum(o[f"nr|{h}|ОстатокОбъектаКол"] for h in SCOPE)
    r["nr_projstock_g"] = o["nr_projstock_g"]; r["nr_projstock_q"] = o["nr_projstock_q"]
    # --- (Г) потребность незавершённого 30 м №3 против остатков (в единицах BuhBud; только где единица и цена сопоставимы)
    rem3 = max(Z, o["n30_q"] - o["c|30-3|q"])
    r["need30_3_q"] = rem3
    r["stock_all_q"] = r["b_stockP_q"] + r["b_stockH_q"]
    r["avail_after30_3_q"] = r["stock_all_q"] - rem3
    r["need_4_7_q"] = o["n30_q"] * 4
    r["deficit_4_7_q"] = max(Z, r["need_4_7_q"] - max(Z, r["avail_after30_3_q"]))
    r["deficit_4_7_s"] = r["deficit_4_7_q"] * r["p30"]
    r["deficit_30_3_q"] = max(Z, rem3 - r["stock_all_q"]); r["deficit_30_3_s"] = r["deficit_30_3_q"] * r["p30"]
    # --- (Д) нижняя граница 15 м: доли неполноты списания
    nq15, ns15 = o["nr|15-1|ПланКол"], o["nr|15-1|ПланГрн"]
    pss15 = dv(ns15, nq15)
    zeros = sum(1 for h in H15 if o[f"c|{h}|q"] == 0)
    r["d15_zero_houses"] = zeros
    r["d15_partial_zero"] = (Decimal(zeros) * nq15 * pss15) if (r["d15_cat"] == "сопоставимо" and 0 < zeros < 6) else Z
    r["left15_g"] = sum(o[f"nr|{h}|ОстатокОбъектаГрн"] for h in H15)
    r["left15_gt"] = sum(o[f"stkh|{h}|gt"] for h in H15)
    r["left30_12_gt"] = sum(o[f"stkh|{h}|gt"] for h in ("30-1", "30-2"))
    # --- (Е) оценка завершения 30 м №3 (цена — из остатков, объём — норма; т.е. только ценовая экономия)
    nq30, ns30 = o["nr|30-1|ПланКол"], o["nr|30-1|ПланГрн"]
    pss30 = dv(ns30, nq30)
    own_q, own_g = o["stkh|30-3|q"], o["stkh|30-3|gt"]
    P_q, P_g = o["stk|P|q"], o["stk|P|gt"]
    need = max(Z, nq30 - o["c|30-3|q"]) if nq30 > 0 else Z
    f_own = min(need, max(Z, own_q)); v_own = f_own * dv(own_g, own_q)
    need2 = need - f_own
    f_P = min(need2, max(Z, P_q)); v_P = f_P * dv(P_g, P_q)
    defi = (need2 - f_P) * pss30
    if nq30 <= 0 and ns30 > 0:            # план суммой без количества
        defi = max(Z, ns30 - o["c|30-3|gt"]); v_own = v_P = Z
    r["e303_cost"] = o["c|30-3|gt"] + v_own + v_P + defi
    r["e303_from_own"] = v_own; r["e303_from_P"] = v_P; r["e303_def"] = defi
    r["e303_own_surplus"] = max(Z, own_g - v_own) if own_q > 0 else max(Z, own_g)
    # --- (Ж) вне границ: 30 м №4–№7 из оставшегося склада проекта (только для сверки с бюджетом экономиста)
    need47 = nq30 * 4
    P_rem = max(Z, P_q - f_P)
    f47 = min(need47, P_rem); v47 = f47 * dv(P_g, P_q)
    d47 = (need47 - f47) * pss30
    if nq30 <= 0 and ns30 > 0: d47 = ns30 * 4; v47 = Z
    r["e47_from_P"] = v47; r["e47_def"] = d47
    r["P_left_after"] = max(Z, P_g - v_P - v47)
    res[k] = r
    for f, v in r.items():
        if isinstance(v, Decimal): tot[f] += v

json.dump({"rows": {k: {f: (str(v) if isinstance(v, Decimal) else v) for f, v in r.items()} for k, r in res.items()},
           "tot": {f: str(v) for f, v in tot.items()}}, open(os.path.join(OUT, "analysis.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)

P = lambda f: fmt(tot[f])
print("(А) Отчёт директору, расчёт: План13", P("n13_s"), "Факт", P("pu_g"), "ETC", P("etc"), "Откл", P("dev"),
      "| ценовой", P("dir_pf"), "кол-во", P("dir_qf"), "остаток разложения", P("dir_resid"))
cats = defaultdict(lambda: [Z, Z, Z, 0])
for r in res.values():
    c = cats[r["dir_cat"]]; c[0] += r["dev"]; c[1] += r["pu_g"]; c[2] += r["etc"]; c[3] += 1
print("   по категориям:", {k: (fmt(v[0]), fmt(v[1]), fmt(v[2]), v[3]) for k, v in cats.items()})
for g in ("d15", "d30"):
    cats = defaultdict(lambda: [Z, Z, Z, 0])
    for r in res.values():
        c = cats[r[f"{g}_cat"]]; c[0] += r[f"{g}_delta"]; c[1] += r[f"{g}_pe"]; c[2] += r[f"{g}_qe"]; c[3] += 1
    print(f"(Б) {g}: норма {P(g + '_Ns')} расход реальн.брутто {P(g + '_cg')} «як план» {P(g + '_cr')} нетто {P(g + '_cn')}"
          f" Δ {P(g + '_delta')} = цена {P(g + '_pe')} + кол-во {P(g + '_qe')} + несопост. {P(g + '_ne')}")
    print("   по категориям:", {k: (fmt(v[0]), fmt(v[1]), fmt(v[2]), v[3]) for k, v in cats.items()})
print("(В) мост нетто: в P", P("b_in_P_n"), "P→H", P("b_P_H_n"), "X→H", P("b_X_H_n"), "MSP→H", P("b_M_H_n"), "H→P", P("b_H_P_n"),
      "P→X", P("b_P_X_n"), "P→MSP", P("b_P_M_n"), "X→P", P("b_X_P_n"), "списано 9 домов", P("b_cons_n"), "прочее из P", P("b_outP_n"),
      "| остаток P", P("b_stockP_n"), "остаток H", P("b_stockH_n"))
print("    отчёт по нормам (9 домов): в нормі", P("nr_in_g"), "понад", P("nr_over_g"), "економія", P("nr_econ_g"),
      "15м", P("nr_econ15_g"), "план", P("nr_plan_g"), "ост.домов", P("nr_hstock_g"), "ост.проекта", P("nr_projstock_g"))
print("(Г) дефицит к норме: 30м №3", P("deficit_30_3_s"), "30м №4–7", P("deficit_4_7_s"))

print("(Д) 15м: Δ", P("d15_delta"), "не использовано/несопост./частичные нули:",
      fmt(sum(r["d15_delta"] for r in res.values() if r["d15_cat"] == "не использовано")),
      fmt(sum(r["d15_delta"] for r in res.values() if r["d15_cat"] == "несопоставимо")), P("d15_partial_zero"),
      "остатки домов 15м: отчёт", P("left15_g"), "реальн.брутто", P("left15_gt"), "| остатки 30м №1-2", P("left30_12_gt"))
print("(Е) 30м №3 оценка стоимости:", P("e303_cost"), "= расход", fmt(sum(r["d30_cg"]*0 for r in res.values())), "из своего склада", P("e303_from_own"),
      "со склада проекта", P("e303_from_P"), "дефицит по цене СС", P("e303_def"), "излишек своего склада", P("e303_own_surplus"))
print("(Ж) 30м №4–7: со склада проекта", P("e47_from_P"), "дефицит по цене СС", P("e47_def"), "склад проекта после всех", P("P_left_after"))
