# -*- coding: utf-8 -*-
"""Шаг 2: мост план ERP (Этап, Аналитика) ↔ листья BuhBud (ЭтапСС, ПозицияСС)."""
import sys, json, re
from collections import defaultdict, Counter
sys.stdout.reconfigure(encoding="utf-8")

E = json.load(open("erp_mat_0710.json", encoding="utf-8"))
B = json.load(open("buh_leaves.json", encoding="utf-8"))

def norm(t):
    return re.sub(r" {2,}", " ", (t or "").strip()).lower()

# СС ERP (по UUID) -> код спецификации
cc2spec = {B[k]["cc_uuid"]: k for k in B}
домов = {"000000101": 7, "000000005": 6, "000000099": 4}
итог = {}
for P in ("МД IRC 2026", "МД ООН Админ"):
    plan = [r for r in E if r["podr"] == P and r["type"] == "СС"]
    specs = sorted(set(cc2spec.get(r["doc_uuid"], "?") for r in plan))
    print("=====", P, "плановых строк", len(plan), "спецификации по UUID СС:", specs,
          Counter(cc2spec.get(r["doc_uuid"], "?") for r in plan))
    for sp in specs:
        PL = [r for r in plan if cc2spec.get(r["doc_uuid"]) == sp]
        LV = B[sp]["leaves"]
        n = домов[sp]
        # этапы: совпадают ли UUID этапов
        eu_erp = Counter((r["etap_uuid"], r["etap"]) for r in PL)
        eu_buh = Counter((l["etap_uuid"], l["etap"]) for l in LV)
        same_uuid = set(u for u, _ in eu_erp) & set(u for u, _ in eu_buh)
        print(" --", sp, B[sp]["name"], "| ERP строк", len(PL), "листьев", len(LV),
              "| этапы ERP", len(eu_erp), "BuhBud", len(eu_buh), "общих UUID", len(same_uuid))
        for (u, nm) in eu_erp:
            if u not in same_uuid:
                print("    этап ERP без пары UUID:", nm, u)
        idx = {}
        for name, f in (("K1 uuid+текст", lambda e, t: (e, t)),
                        ("K2 uuid+норм", lambda e, t: (e, norm(t))),
                        ("K3 имя этапа+норм", None),
                        ("K4 только норм", None)):
            pass
        L_k1 = defaultdict(list); L_k2 = defaultdict(list); L_k3 = defaultdict(list); L_k4 = defaultdict(list)
        for l in LV:
            L_k1[(l["etap_uuid"], l["pos"])].append(l)
            L_k2[(l["etap_uuid"], norm(l["pos"]))].append(l)
            L_k3[(norm(l["etap"]), norm(l["pos"]))].append(l)
            L_k4[norm(l["pos"])].append(l)
        cnt = Counter(); sums = Counter(); unmatched = []
        matched_leaf_keys = set()
        for r in PL:
            how = None
            if (r["etap_uuid"], r["an"]) in L_k1: how = "K1"; ls = L_k1[(r["etap_uuid"], r["an"])]
            elif (r["etap_uuid"], norm(r["an"])) in L_k2: how = "K2"; ls = L_k2[(r["etap_uuid"], norm(r["an"]))]
            elif (norm(r["etap"]), norm(r["an"])) in L_k3: how = "K3"; ls = L_k3[(norm(r["etap"]), norm(r["an"]))]
            elif norm(r["an"]) in L_k4: how = "K4"; ls = L_k4[norm(r["an"])]
            else: ls = []
            cnt[how or "нет"] += 1; sums[how or "нет"] += r["pg"]
            for l in ls: matched_leaf_keys.add(l["key"])
            if not how: unmatched.append(r)
            r["_spec"] = sp; r["_how"] = how; r["_leaves"] = [l["key"] for l in ls]
        tot = sum(r["pg"] for r in PL)
        print("    сопоставление строк:", dict(cnt), "| суммы:", {k: round(v, 2) for k, v in sums.items()},
              "| план ERP %.2f" % tot)
        for r in unmatched:
            print("    НЕ СОПОСТ: этап=%r поз=%r ОН=%r план=%.2f" % (r["etap"], r["an"], r["on"], r["pg"]))
        lv_un = [l for l in LV if l["key"] not in matched_leaf_keys]
        print("    листьев без плановой строки ERP:", len(lv_un), "Σнорм×дом %.2f" % (sum(l["sum"] for l in lv_un) * n))
        for l in lv_un[:25]:
            print("      лист: этап=%r поз=%r карт=%r Σ=%.2f НП=%s вид=%s" % (l["etap"], l["pos"], l["nom_name"], l["sum"], l["np"], l["vid"]))
        итог[sp] = {"rows": len(PL), "cnt": dict(cnt), "sums": dict(sums), "tot": tot}
json.dump({"plan": [r for r in E if r["type"] == "СС"]}, open("bridge_plan.json", "w", encoding="utf-8"), ensure_ascii=False, indent=0)
