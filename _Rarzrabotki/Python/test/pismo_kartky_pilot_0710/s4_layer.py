# -*- coding: utf-8 -*-
"""Шаги 3–5: гейты и карточный слой по промту §3.3–3.5 на данных 07.10.2026 (только расчёт, без баз)."""
import sys, json, re, time
from collections import defaultdict, Counter
sys.stdout.reconfigure(encoding="utf-8")
T0 = time.time()
E = json.load(open("erp_mat_0710.json", encoding="utf-8"))
B = json.load(open("buh_leaves.json", encoding="utf-8"))
NULL = "00000000-0000-0000-0000-000000000000"
VERBOSE = "-v" in sys.argv
FIX_MULTI = "on" if "--fix" in sys.argv else "off"

def norm(t):
    return re.sub(r" {2,}", " ", (t or "").strip()).lower()

def unorm(u):
    u = norm(u).replace(" ", "")
    return {"пог.м": "м", "м.п.": "м", "мп": "м", "м2": "м2", "м²": "м2"}.get(u, u)

def gkey(r):
    if r["on"].strip():
        return r["on"].strip()
    if r["nom_uuid"] and r["nom_uuid"] != NULL:
        return "[НОМ] " + r["nom_name"].strip()
    return "[ТЕКСТ] " + r["an"].strip().upper()

cc2spec = {B[k]["cc_uuid"]: k for k in B}
PROJ = {"МД IRC 2026": [("000000005", 6), ("000000101", 7)], "МД ООН Админ": [("000000099", 4)]}
XLSX_DEV = {"МД IRC 2026": -383128.91, "МД ООН Админ": 152466.63}
SPEC = ["Сигналізація виявлення диму", "Двері МДФ внутрішні", "Кутики внутрішні", "Комплект септика", "Гайка",
        "Профільна труба", "Лист ГК (металопрокат)", "лист пвл"]
REPORT = {}

def fmt(v):
    return ("%+.2f" % v) if v else "0"

for P, specs in PROJ.items():
    t1 = time.time()
    R = [r for r in E if r["podr"] == P]
    PL = [r for r in R if r["type"] == "СС"]
    FA = [r for r in R if r["type"] == "ПТУ"]
    # ---------- группы ERP (как xlsx) ----------
    G = defaultdict(lambda: dict(pk=0.0, pg=0.0, fk=0.0, fg=0.0, pnf=0.0))
    for r in R:
        g = G[gkey(r)]
        g["pk"] += r["pk"]; g["pg"] += r["pg"]; g["fk"] += r["fk"]; g["fg"] += r["fg"]; g["pnf"] += r["pnf"]
    for k, g in G.items():
        g["L"] = g["fg"] + g["pnf"] - g["pg"]
        if g["pk"]:
            cp = g["pg"] / g["pk"]
            g["c_price"] = g["fg"] - g["fk"] * cp
            g["c_qty"] = max(0.0, g["fk"] - g["pk"]) * cp
        else:
            g["c_price"], g["c_qty"] = g["L"], 0.0
    # ---------- мост план ERP <- листья ----------
    plan_idx = defaultdict(list)
    for r in PL:
        plan_idx[(cc2spec.get(r["doc_uuid"]), r["etap_uuid"], r["an"])].append(r)
    leaves = []
    excl15 = []; multi = []
    for sp, n in specs:
        cc_sum = defaultdict(float)
        for p in B[sp]["cc_pos"]:
            cc_sum[(p["etap_uuid"], norm(p["pos"]))] += p["sum"]
        for l in B[sp]["leaves"]:
            rows = plan_idx.get((sp, l["etap_uuid"], l["pos"].strip()), [])
            if not rows:
                rows = [r for (s, e, a), rr in plan_idx.items() for r in rr if s == sp and e == l["etap_uuid"] and norm(a) == norm(l["pos"])]
            l = dict(l); l["spec"] = sp; l["n"] = n
            if sp == "000000005" and abs(cc_sum.get((l["etap_uuid"], norm(l["pos"])), 0.0)) < 0.005:
                l["excl"] = "Р22 (сумма позиции СС = 0)"; excl15.append(l)
            if not rows:
                l["g"] = None
                leaves.append(l)
            elif len(set(gkey(r) for r in rows)) == 1:
                l["g"] = gkey(rows[0]); leaves.append(l)
            else:
                # ключ (Этап, позиция) -> несколько строк ERP с разной ОН (дубли строк СС, BuhBud их слил)
                multi.append((l["pos"], [gkey(r) for r in rows]))
                same = [r for r in rows if norm(r["on"]) == norm(l["on"])]
                leaf_ons = set(norm(x["on"]) for x in B[sp]["leaves"] if x["etap_uuid"] == l["etap_uuid"] and x["pos"].strip() == l["pos"].strip())
                covers = all(norm(r["on"]) in leaf_ons for r in rows)
                if FIX_MULTI == "on" and len(same) == 1 and covers:
                    l["g"] = gkey(same[0]); l["bridge"] = "ОН листа"; leaves.append(l)
                elif FIX_MULTI == "on":
                    tot = sum(r["pg"] for r in rows) or 1.0
                    for r in rows:
                        x = dict(l); x["g"] = gkey(r); x["bridge"] = "пропорция"
                        x["kol"] = l["kol"] * r["pg"] / tot; x["sum"] = l["sum"] * r["pg"] / tot
                        leaves.append(x)
                else:
                    l["g"] = gkey(max(rows, key=lambda r: r["pg"])); leaves.append(l)
    LV = [l for l in leaves if not l.get("excl")]
    print("=" * 100)
    print(P, "| листьев", len(leaves), "| исключено Р22:", len(excl15),
          "Σ×1 %.2f" % sum(l["sum"] for l in excl15), "| без плановой строки:", sum(1 for l in LV if l["g"] is None),
          "| ключ → несколько групп ERP:", len(multi))
    for sp, n in specs:
        s1 = sum(l["sum"] for l in LV if l["spec"] == sp)
        s0 = sum(l["sum"] for l in leaves if l["spec"] == sp)
        print("  спец %s: Σнорм все %.2f, без Р22 %.2f (на дом), ×%d = %.2f" % (sp, s0, s1, n, s1 * n))
    if excl15:
        print("  Р22-исключённые листья 15 м (сумма СС=0):", Counter(l["g"] for l in excl15).most_common())
    # ---------- карточки ----------
    explicit_all = set(l["nom_uuid"] for l in LV if l["vid"] == "Деталь" and not l["np"] and l["nom_uuid"] not in ("", NULL))
    cards = defaultdict(lambda: defaultdict(lambda: dict(pk=0.0, pg=0.0, fk=0.0, fg=0.0, kind="", name="", ed="", nleaves=0,
                                                          types=set(), split=False, path="", buh_on="", erp_ed="")))
    pos_cards = defaultdict(set)
    for l in LV:
        if l["g"] is None:
            continue
        if l["np"]:
            cid, kind, name = "НП|" + l["pos"], "НП", "не постачається: " + l["pos"]
        elif l["vid"] != "Деталь":
            cid, kind, name = "КОТЁЛ|" + l["on"], "котёл", "котёл ОН «%s»" % l["on"]
        else:
            cid, kind, name = l["nom_uuid"], "картка", l["nom_name"]
        c = cards[l["g"]][cid]
        c["pk"] += l["kol"] * l["n"]; c["pg"] += l["sum"] * l["n"]; c["kind"] = kind; c["name"] = name
        c["ed"] = l["nom_ed"] if kind == "картка" else l["ed_str"]; c["nleaves"] += 1; c["types"].add(l["spec"])
        c["buh_on"] = l["on"]
        if l["lvl"] >= 1 and l["spec"] == "000000101":
            c["kind"] = "комплект-деталь"; c["path"] = l["path"]
        pos_cards[(l["spec"], l["etap_uuid"], l["pos"])].add(cid)
    split_pos = {k: v for k, v in pos_cards.items() if len(v) > 1}
    # ---------- гейт 1 ----------
    g1_ok, g1_bad = [], []
    for g, gd in G.items():
        if gd["pg"] == 0 and g not in cards:
            continue
        s = sum(c["pg"] for c in cards.get(g, {}).values())
        nl = sum(c["nleaves"] * 1 for c in cards.get(g, {}).values())
        n_max = max(n for _, n in specs)
        tol = max(0.01, 0.01 * nl * n_max)
        d = s - gd["pg"]
        (g1_ok if abs(d) <= tol else g1_bad).append((g, gd["pg"], s, d, tol))
    print("  ГЕЙТ 1 (Σ План грн карток = План групи): OK %d, FAIL %d; Σ карток %.2f vs План СС %.2f" % (
        len(g1_ok), len(g1_bad), sum(c["pg"] for g in cards for c in cards[g].values()), sum(r["pg"] for r in PL)))
    for g, pg, s, d, tol in sorted(g1_bad, key=lambda x: -abs(x[3])):
        print("     FAIL %-40s план ERP %12.2f  картки %12.2f  Δ %+10.2f (допуск %.2f)" % (g[:40], pg, s, d, tol))
    worst_ok = sorted(g1_ok, key=lambda x: -abs(x[3]))[:3]
    print("     макс. Δ среди OK:", [(g, round(d, 2)) for g, _, _, d, _ in worst_ok])
    # ---------- факт ----------
    fact_card = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, "", "", ""]))  # card -> group -> [fk, fg, name, ed, nom_on]
    for r in FA:
        f = fact_card[r["nom_uuid"]][gkey(r)]
        f[0] += r["fk"]; f[1] += r["fg"]; f[2] = r["nom_name"]; f[3] = r["nom_ed"]; f[4] = r["nom_on"]
    # гейт 2: Σ факта по группе = факт группы
    g2 = defaultdict(float)
    for cu, gg in fact_card.items():
        for g, f in gg.items():
            g2[g] += f[1]
    g2_bad = [(g, G[g]["fg"], g2.get(g, 0.0)) for g in G if abs(G[g]["fg"] - g2.get(g, 0.0)) > 0.005]
    print("  ГЕЙТ 2 (Σ факт карток ERP по ОН = факт групи): FAIL %d из %d (тавтология — тот же источник)" % (len(g2_bad), len(G)))
    # единицы BuhBud vs ERP для явных карток
    ed_mis = []
    for g, cc in cards.items():
        for cid, c in cc.items():
            if c["kind"] in ("картка", "комплект-деталь") and cid in fact_card:
                ed_erp = next(iter(fact_card[cid].values()))[3]
                c["erp_ed"] = ed_erp
                if unorm(ed_erp) != unorm(c["ed"]):
                    ed_mis.append((g, c["name"], c["ed"], ed_erp))
    # ---------- V1: факт картки только в своей группе ----------
    res = {}
    cross = []   # план в g, закупка в h
    for g, gd in G.items():
        cc = cards.get(g, {})
        purchases = {cu: gg[g] for cu, gg in fact_card.items() if g in gg}
        expl = {cid for cid, c in cc.items() if c["kind"] in ("картка", "комплект-деталь")}
        kotly = {cid: c for cid, c in cc.items() if c["kind"] == "котёл"}
        assigned = defaultdict(lambda: [0.0, 0.0]); others = []; others_sum = 0.0
        for cu, f in purchases.items():
            if cu in expl:
                assigned[cu][0] += f[0]; assigned[cu][1] += f[1]
                continue
            hit = None
            if cu not in explicit_all:
                for kid, k in kotly.items():
                    if unorm(f[3]) == unorm(k["ed"]):
                        hit = kid; break
            if hit:
                assigned[hit][0] += f[0]; assigned[hit][1] += f[1]
            else:
                others.append((f[2], f[0], f[1], f[3], "план в другой группе" if cu in explicit_all else ""))
                others_sum += f[1]
        rows = []
        Lc = pr = qt = 0.0
        for cid, c in cc.items():
            fk, fg = assigned[cid] if cid in assigned else (0.0, 0.0)
            if c["kind"] == "НП":
                lc = p_ = q_ = 0.0
            elif c["pk"] > 0:
                cp = c["pg"] / c["pk"]
                if fk == 0 and fg == 0:
                    p_ = q_ = 0.0
                else:
                    p_ = fg - fk * cp
                    q_ = max(0.0, fk - c["pk"]) * cp
                lc = p_ + q_
            else:
                lc = max(0.0, fg - c["pg"]); p_, q_ = lc, 0.0
            rows.append(dict(cid=cid, name=c["name"], kind=c["kind"], ed=c["ed"], erp_ed=c.get("erp_ed", ""), pk=c["pk"], pg=c["pg"],
                             cp=(c["pg"] / c["pk"] if c["pk"] else 0), fk=fk, fg=fg, cf=(fg / fk if fk else 0),
                             price=p_, qty=q_, L=lc, types=sorted(c["types"])))
            Lc += lc; pr += p_; qt += q_
            # план карты в g, а закупки вне g
            if c["kind"] in ("картка", "комплект-деталь") and cid in fact_card:
                for h, f in fact_card[cid].items():
                    if h != g:
                        cross.append((g, h, c["name"], c["pg"], f[1], f[0]))
        Rg = gd["L"] - Lc - others_sum
        ncards = sum(1 for r in rows if r["kind"] != "НП")
        hetero_reasons = []
        if ncards >= 2: hetero_reasons.append("%d карток" % ncards)
        if any(r["kind"] == "котёл" for r in rows): hetero_reasons.append("котёл")
        if any(r["kind"] == "комплект-деталь" for r in rows): hetero_reasons.append("комплект")
        if gd["L"] and abs(Rg) >= 0.05 * abs(gd["L"]): hetero_reasons.append("|R|≥5%")
        cls_g = "ціна" if abs(gd["c_price"]) >= abs(gd["c_qty"]) else "обсяг"
        cls_c = ("ціна" if abs(pr) >= abs(qt) else "обсяг") if rows else "—"
        res[g] = dict(L=gd["L"], pg=gd["pg"], pk=gd["pk"], fk=gd["fk"], fg=gd["fg"], gp=gd["c_price"], gq=gd["c_qty"],
                      Lc=Lc, pr=pr, qt=qt, others=others_sum, others_list=others, R=Rg, ncards=ncards, rows=rows,
                      hetero=hetero_reasons, cls_g=cls_g, cls_c=cls_c)
    # ---------- V2: перенос ОН (план картки в g, закупка в h) ----------
    plan_by_card = defaultdict(lambda: defaultdict(float))
    for g, cc in cards.items():
        for cid, c in cc.items():
            if c["kind"] in ("картка", "комплект-деталь"):
                plan_by_card[cid][g] += c["pg"]
    T = defaultdict(float)
    for cid, gg in plan_by_card.items():
        pgk = max(gg, key=gg.get)
        for h, f in fact_card.get(cid, {}).items():
            if h != pgk:
                T[h] += f[1]; T[pgk] -= f[1]
    home = {cid: max(gg, key=gg.get) for cid, gg in plan_by_card.items()}
    for g in res:
        res[g]["T"] = T.get(g, 0.0)
        # V2: картка в «домашней» группе плана — весь факт по проекту, в остальных группах плана — без факта;
        # інші2 — только картки, которых нет в плане проекта; перенос T — отдельной строкой
        Lc2 = pr2 = qt2 = 0.0
        for r in res[g]["rows"]:
            fk, fg = r["fk"], r["fg"]
            if r["kind"] in ("картка", "комплект-деталь"):
                if home.get(r["cid"]) == g:
                    fk = sum(f[0] for f in fact_card.get(r["cid"], {}).values())
                    fg = sum(f[1] for f in fact_card.get(r["cid"], {}).values())
                else:
                    fk = fg = 0.0
            if r["kind"] == "НП" or (fk == 0 and fg == 0):
                p_ = q_ = 0.0
            elif r["pk"] > 0:
                p_ = fg - fk * r["cp"]; q_ = max(0.0, fk - r["pk"]) * r["cp"]
            else:
                p_, q_ = max(0.0, fg - r["pg"]), 0.0
            r["fk2"], r["fg2"], r["price2"], r["qty2"], r["L2"] = fk, fg, p_, q_, p_ + q_
            Lc2 += p_ + q_; pr2 += p_; qt2 += q_
        oth2 = sum(o[2] for o in res[g]["others_list"] if not o[4])
        res[g].update(Lc2=Lc2, pr2=pr2, qt2=qt2, others2=oth2, R2=res[g]["L"] - Lc2 - oth2 - res[g]["T"],
                      cls_c2=("ціна" if abs(pr2) >= abs(qt2) else "обсяг") if res[g]["rows"] else "—")
    print("  V2: ΣL_c %.2f (ціна %.2f, обсяг %.2f), інші(непланові) %.2f, Σперенос %.2f, ΣR2 %.2f, Σ|R2| %.2f" % (
        sum(v["Lc2"] for v in res.values()), sum(v["pr2"] for v in res.values()), sum(v["qt2"] for v in res.values()),
        sum(v["others2"] for v in res.values()), sum(v["T"] for v in res.values()),
        sum(v["R2"] for v in res.values()), sum(abs(v["R2"]) for v in res.values())))
    # ---------- гейт 3 ----------
    tot_L = sum(v["L"] for v in res.values())
    tot_dec = sum(v["Lc"] + v["others"] + v["R"] for v in res.values())
    print("  ГЕЙТ 3: Σ L_g = %.2f; Σ(ΣL_c+інші+R_g) = %.2f; xlsx %.2f → Δ %.2f (тавтология по построению R_g)" % (
        tot_L, tot_dec, XLSX_DEV[P], tot_dec - XLSX_DEV[P]))
    print("  Σ по проекту: ΣL_c %.2f (ціновий %.2f, кількісний %.2f), інші %.2f, ΣR_g %.2f" % (
        sum(v["Lc"] for v in res.values()), sum(v["pr"] for v in res.values()), sum(v["qt"] for v in res.values()),
        sum(v["others"] for v in res.values()), sum(v["R"] for v in res.values())))
    print("  Σ |R_g| по проекту %.2f; групп с |R_g|≥5%%|L_g| и |R_g|≥1000: %d" % (
        sum(abs(v["R"]) for v in res.values()), sum(1 for v in res.values() if v["L"] and abs(v["R"]) >= max(1000, 0.05 * abs(v["L"])))))
    print("  несовпадение единиц картки BuhBud vs ERP (явные картки с фактом): %d" % len(ed_mis))
    for x in ed_mis[:15]:
        print("     ", x)
    print("  позиции СС, разложенные на ≥2 картки: %d" % len(split_pos))
    # кросс-группы (план в g, закупка в h)
    cr = defaultdict(lambda: [0.0, 0.0, set()])
    for g, h, nm, pg, fg, fk in cross:
        cr[(g, h)][0] += pg; cr[(g, h)][1] += fg; cr[(g, h)][2].add(nm)
    print("  пары «план в ОН X — закупка в ОН Y» (картки с фактом): %d пар, Σ план %.2f, Σ закупка %.2f" % (
        len(cr), sum(v[0] for v in cr.values()), sum(v[1] for v in cr.values())))
    for (g, h), v in sorted(cr.items(), key=lambda x: -x[1][1])[:20]:
        print("     план «%s» → закупка «%s»: план %.2f, закупка %.2f [%s]" % (g, h, v[0], v[1], "; ".join(sorted(v[2]))[:90]))
    # ---------- ТОП ----------
    kosht = [(g, v) for g, v in res.items() if v["pg"] > 0]
    econ = sorted([x for x in kosht if x[1]["L"] < 0], key=lambda x: x[1]["L"])[:15]
    over = sorted([x for x in kosht if x[1]["L"] > 0], key=lambda x: -x[1]["L"])[:15]
    for title, lst in (("ТОП ЕКОНОМІЙ", econ), ("ТОП ПЕРЕВИТРАТ", over)):
        print("  ---- %s ----" % title)
        print("  %-34s %12s | %11s %11s | %11s %11s %10s %11s %10s | %3s %-5s %-5s %s" % (
            "група", "L_g", "груп.ціна", "груп.обсяг", "ΣLc ціна", "ΣLc обсяг", "інші", "R_g", "перенос V2", "к-т", "кл.гр", "кл.к", "неоднор."))
        for g, v in lst:
            chg = " <<ЗМІНА" if v["cls_g"] != v["cls_c"] else ""
            print("  %-34s %12.2f | %11.2f %11.2f | %11.2f %11.2f %10.2f %11.2f %10.2f | %3d %-5s %-5s %s%s" % (
                g[:34], v["L"], v["gp"], v["gq"], v["pr"], v["qt"], v["others"], v["R"], v["T"], v["ncards"], v["cls_g"], v["cls_c"],
                ",".join(v["hetero"]), chg))
    REPORT[P] = dict(res=res, cross={"%s → %s" % k: [v[0], v[1], sorted(v[2])] for k, v in cr.items()},
                     g1_bad=g1_bad, excl15=[(l["pos"], l["g"], l["sum"]) for l in excl15], ed_mis=ed_mis,
                     econ=[g for g, _ in econ], over=[g for g, _ in over])
    print("  время расчёта проекта %.2fs" % (time.time() - t1))

# ---------- детально по особым группам ----------
for P in REPORT:
    res = REPORT[P]["res"]
    for g in SPEC:
        k = next((x for x in res if x.lower() == g.lower()), None)
        if not k:
            continue
        v = res[k]
        print("#" * 100)
        print("%s | «%s»: L_g %.2f | план %.2f (к-сть %.3f) факт %.2f (к-сть %.3f) | груп: ціна %.2f обсяг %.2f [%s] | картки: ціна %.2f обсяг %.2f [%s] | інші %.2f | R_g %.2f | %s" % (
            P, k, v["L"], v["pg"], v["pk"], v["fg"], v["fk"], v["gp"], v["gq"], v["cls_g"], v["pr"], v["qt"], v["cls_c"], v["others"], v["R"], ",".join(v["hetero"])))
        for r in sorted(v["rows"], key=lambda r: -abs(r["L"])):
            print("    %-14s %-48s ед %-5s/%-5s план %9.3f × %10.2f = %11.2f | факт %9.3f × %10.2f = %11.2f | ціна %+11.2f обсяг %+10.2f L_c %+11.2f" % (
                r["kind"], r["name"][:48], r["ed"], r["erp_ed"], r["pk"], r["cp"], r["pg"], r["fk"], r["cf"], r["fg"], r["price"], r["qty"], r["L"]))
        for o in v["others_list"]:
            print("    інша картка   %-48s к-сть %9.3f %-5s грн %11.2f %s" % (o[0][:48], o[1], o[3], o[2], o[4]))

def ser(o):
    if isinstance(o, set):
        return sorted(o)
    raise TypeError
json.dump(REPORT, open("layer_0710.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=ser)
print("ВСЕГО %.2fs" % (time.time() - T0))
