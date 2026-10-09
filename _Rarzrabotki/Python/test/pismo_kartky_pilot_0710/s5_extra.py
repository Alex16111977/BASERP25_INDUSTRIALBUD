# -*- coding: utf-8 -*-
"""Доп. проверки: картка в ≥2 группах плана; V2 (факт картки по проекту + перенос ОН); к-сть BuhBud vs ERP."""
import sys, json, re
from collections import defaultdict, Counter
sys.stdout.reconfigure(encoding="utf-8")
E = json.load(open("erp_mat_0710.json", encoding="utf-8"))
B = json.load(open("buh_leaves.json", encoding="utf-8"))
R = json.load(open("layer_0710.json", encoding="utf-8"))
NULL = "00000000-0000-0000-0000-000000000000"
def norm(t): return re.sub(r" {2,}", " ", (t or "").strip()).lower()
def unorm(u): return {"пог.м": "м"}.get(norm(u).replace(" ", ""), norm(u).replace(" ", ""))
def gkey(r):
    if r["on"].strip(): return r["on"].strip()
    if r["nom_uuid"] and r["nom_uuid"] != NULL: return "[НОМ] " + r["nom_name"].strip()
    return "[ТЕКСТ] " + r["an"].strip().upper()
cc2spec = {B[k]["cc_uuid"]: k for k in B}
PROJ = {"МД IRC 2026": [("000000005", 6), ("000000101", 7)], "МД ООН Админ": [("000000099", 4)]}
for P, specs in PROJ.items():
    PL = [r for r in E if r["podr"] == P and r["type"] == "СС"]
    FA = [r for r in E if r["podr"] == P and r["type"] == "ПТУ"]
    idx = {(cc2spec.get(r["doc_uuid"]), r["etap_uuid"], r["an"]): r for r in PL}
    card_groups = defaultdict(lambda: defaultdict(float))
    qty_cmp = defaultdict(lambda: [0.0, set(), None])  # erp row id -> [Σ leaf kol*n?] per house
    for sp, n in specs:
        ccs = defaultdict(float)
        for p in B[sp]["cc_pos"]: ccs[(p["etap_uuid"], norm(p["pos"]))] += p["sum"]
        for l in B[sp]["leaves"]:
            if sp == "000000005" and abs(ccs.get((l["etap_uuid"], norm(l["pos"])), 0)) < 0.005: continue
            r = idx.get((sp, l["etap_uuid"], l["pos"].strip()))
            if r is None: continue
            g = gkey(r)
            if not l["np"] and l["vid"] == "Деталь" and l["nom_uuid"] not in ("", NULL):
                card_groups[l["nom_uuid"] + "|" + l["nom_name"]][g] += l["sum"] * n
            q = qty_cmp[id(r)]; q[0] += l["kol"] * n; q[1].add(unorm(l["nom_ed"] if l["vid"] == "Деталь" else l["ed_str"])); q[2] = r
    multi = {c: dict(v) for c, v in card_groups.items() if len(v) > 1}
    print("=====", P)
    print(" картки, запланированные в ≥2 группах ERP:", len(multi), "| план на них %.2f" % sum(sum(v.values()) for v in multi.values()))
    for c, v in multi.items(): print("    ", c.split("|")[1][:50], {k: round(x, 2) for k, x in v.items()})
    # к-сть плана: BuhBud vs ERP по позициям с одной единицей
    same_u = diff_round = diff_big = 0; big = []
    for q in qty_cmp.values():
        r = q[2]
        if len(q[1]) != 1 or unorm(r["ed"]) not in q[1] or r["pk"] == 0: continue
        same_u += 1
        d = q[0] - r["pk"]
        if abs(d) < 1e-6: continue
        rel = abs(d) / r["pk"]
        if rel <= 0.02 or abs(d) <= 0.01 * (7 if P.startswith("МД IRC") else 4): diff_round += 1
        else: diff_big += 1; big.append((r["doc"][:12], r["an"][:40], r["ed"], r["pk"], round(q[0], 3)))
    print(" позиции с той же единицей карты и ОН: %d; к-сть BuhBud ≠ ERP: округление %d, существенно %d" % (same_u, diff_round, diff_big))
    for b in big[:40]: print("    ", b)
    # V2: факт картки по всему проекту, перенос ОН явной строкой
    res = R[P]["res"]
    fact_card = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))
    for r in FA:
        f = fact_card[r["nom_uuid"]][gkey(r)]; f[0] += r["fk"]; f[1] += r["fg"]
    planned_in = {}
    for c, v in card_groups.items():
        planned_in[c.split("|")[0]] = max(v, key=v.get)
    T = defaultdict(float)
    for cu, gg in fact_card.items():
        pg = planned_in.get(cu)
        if not pg: continue
        for h, f in gg.items():
            if h != pg:
                T[h] += f[1]   # пришло в h (деньги чужой позиции)
                T[pg] -= f[1]  # ушло из плановой группы
    print(" V2 перенос ОН: Σ T_g = %.2f (должно 0); групп с переносом %d" % (sum(T.values()), sum(1 for v in T.values() if abs(v) > 0.005)))
    for g in ["Кутики внутрішні", "Плінтус на стільницю", "Гладкий лист метал", "Гладкий лист", "Мембрана", "Паробар'єр", "Вітробарєр",
              "Кутики декоративні", "Держак ринви", "Комплект кріплення труб", "Хомут", "Вентилятор", "Витяжка", "Дріт", "Кабель"]:
        if g in res:
            v = res[g]
            print("    %-26s L %11.2f | V1: ΣLc %11.2f інші %10.2f R %11.2f | V2-перенос %+11.2f" % (g, v["L"], v["Lc"], v["others"], v["R"], T.get(g, 0.0)))
