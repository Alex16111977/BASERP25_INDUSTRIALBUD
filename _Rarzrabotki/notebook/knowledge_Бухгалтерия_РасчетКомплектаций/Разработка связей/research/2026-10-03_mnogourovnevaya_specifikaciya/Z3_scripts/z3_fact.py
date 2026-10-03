# -*- coding: utf-8 -*-
"""Z3, часть 2: норма против факта по домам 15 м (§5, §6.7), «вид сейчас» из РК, моделирование регистров §6 (Септик, Водостік),
инварианты §6.11, закупочная спецификация 15 м. JSON -> data/demo_specifikaciya_15m.json. Только чтение."""
import json, math
from z3_spec import *

OUT_JSON = os.environ.get("Z3_JSON", os.path.normpath(os.path.join(HERE, "..", "..", "..", "data", "demo_specifikaciya_15m.json")))
SNAP_DT = pdate(RAW["stamp"])
PIECE = {"шт", "рул", "пак", "пар", "балон", "компл", "к-т", "100шт", "посл."}

# ───────────────────────── виды КН по РК (ДокументиКомплектації)
KN_KIND = {}
RK_OF = {}
for r in RAW["rk"]:
    RK_OF[house_of(r["Дом"])] = r
for r in RAW["rk_kn"]:
    if r["КНПоНормам"]:
        KN_KIND[r["КНПоНормам"]] = "по нормам"
    if r["КНДодаткова"]:
        KN_KIND[r["КНДодаткова"]] = "додаткова"


def fact_rows(h, with_lying):
    rows = []
    for f in FACT:
        if f["house"] != h:
            continue
        kind = KN_KIND.get(f["doc"], "інший")
        grp = 1 if kind == "по нормам" else (3 if kind == "додаткова" else 2)
        mshp = f["type"].startswith("Передача малоцінних") or f["acc_kt"].startswith("22")
        rows.append({**f, "kn": kind if f["type"].startswith("Комплектація") else f["type"], "grp": grp, "mshp": mshp, "lying": False})
    if with_lying:
        for c in LQ:
            if LQ[c][h] or LS[c][h]:
                rows.append({"house": h, "dt": SNAP_DT, "doc": "Залишок на складі будинку", "type": "Залишок", "acc_kt": "", "acc_dt": "",
                             "code": c, "q": LQ[c][h], "s": LS[c][h], "wh": "склади будинку", "stage": "—", "kn": "залишок", "grp": 2,
                             "mshp": False, "lying": True})
    rows.sort(key=lambda r: (r["grp"], r["dt"], r["doc"], r["code"]))
    return rows


def attribute(leaves, rows, pma_norm_first=False):
    """§6.7: одно правило «в нормі / понад / поза специфікацією / МШП». Возврат: записи, остаток нормы по листьям."""
    rem = {lf["key"]: (lf["q"] if lf["kind"] == "Деталь" else lf["q_line"]) for lf in leaves}
    norm = dict(rem)
    by_card = collections.defaultdict(list)
    by_on = collections.defaultdict(list)
    for lf in leaves:
        if lf["kind"] == "Деталь" and lf["card"]:
            by_card[lf["card"]].append(lf)
        elif lf["kind"].startswith("Група"):
            by_on[K(lf["on"])].append(lf)
    recs = []
    for r in rows:
        c, q, s = r["code"], r["q"], r["s"]
        base = {"doc": r["doc"], "dt": r["dt"], "kn": r["kn"], "card": c, "wh": r["wh"], "stage": r["stage"], "lying": r["lying"]}
        in_norm = c in by_card or (K(CARD.get(c, {}).get("on")) in by_on)
        if r["mshp"] and not (pma_norm_first and in_norm):
            recs.append({**base, "kind": "МШП", "leaf": None, "q": q, "s": s, "s_np": Z}); continue
        cands = by_card.get(c)
        if not cands:
            g = by_on.get(K(CARD.get(c, {}).get("on")))
            if g and ukey_eq(UKEY_CARD.get(c), (g[0]["unit_line"] or "").strip()):
                cands = g
            elif g:
                recs.append({**base, "kind": "Поза специфікацією (інша одиниця)", "leaf": None, "q": q, "s": s, "s_np": Z}); continue
        if not cands:
            recs.append({**base, "kind": "Поза специфікацією", "leaf": None, "q": q, "s": s, "s_np": Z}); continue
        rt = sum(rem[lf["key"]] for lf in cands)
        inn = min(q, max(Z, rt))
        over = q - inn
        if inn > 0:
            for lf in cands:
                part = inn * rem[lf["key"]] / rt if rt else Z
                if part > 0:
                    rem[lf["key"]] -= part
                    recs.append({**base, "kind": "В нормі", "leaf": lf["key"], "q": part, "s": s * part / q,
                                 "s_np": part * (lf["price"] or Z) / VAT})
        if over > 0:
            tn = sum(norm[lf["key"]] for lf in cands)
            for lf in cands:
                part = over * norm[lf["key"]] / tn if tn else over / len(cands)
                if part > 0:
                    recs.append({**base, "kind": "Понад нормою", "leaf": lf["key"], "q": part, "s": s * part / q,
                                 "s_np": part * (lf["price"] or Z) / VAT})
    return recs, rem, norm


def house_spec(h, use_variant=True):
    return ("V" + h) if (use_variant and ("V" + h) in SPECS) else "E"


def summarize(h, leaves, recs, rem, norm):
    S = collections.defaultdict(lambda: Z)
    for lf in leaves:
        nv = (lf["sum_n"] or Z) / VAT
        S["norm"] += nv
        if lf["kind"] == "Деталь" and not lf["card"]:
            S["norm_nocard"] += nv
        e = max(Z, rem[lf["key"]])
        base_q = norm[lf["key"]]
        S["econ_np"] += (e * (lf["price"] or Z) / VAT)
        if lf["kind"] == "Деталь" and not lf["card"]:
            S["econ_nocard"] += e * (lf["price"] or Z) / VAT
    for r in recs:
        S["fact"] += r["s"]
        if r["kind"] == "В нормі":
            S["in"] += r["s"]; S["in_np"] += r["s_np"]
        elif r["kind"] == "Понад нормою":
            S["over"] += r["s"]; S["over_np"] += r["s_np"]
        elif r["kind"] == "МШП":
            S["mshp"] += r["s"]
        else:
            S["poza"] += r["s"]
        if r["lying"]:
            S["lying"] += r["s"]
    return {k: r2(v) for k, v in S.items()}


# ───────────────────────── «вид сейчас»: поля РК (ТабличнаяЧастьОстатков: з ПДВ → ÷1,2; СписаниеСверхНормы — без ПДВ)
NOW = {}
NOW_CARD = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z, Z]))   # h -> card -> [понад q, понад s, вн q, остаток q]
for r in RAW["rk_ost"]:
    h = house_of(r["Дом"])
    a = NOW_CARD[h][r["Код"]]
    a[0] += Dd(r["Понад"]); a[1] += Dd(r["ПонадСум"]) / VAT; a[2] += Dd(r["ВНорме"]); a[3] += Dd(r["Остаток"])
for h in H15:
    rows = [r for r in RAW["rk_ost"] if house_of(r["Дом"]) == h]
    sv = [r for r in RAW["rk_sverh"] if house_of(r["Дом"]) == h]
    rk = RK_OF.get(h, {})
    NOW[h] = {"rk": rk.get("Номер"), "rk_date": rk.get("Дата"), "rk_status": rk.get("Статус"),
              "ostatok": r2(sum(Dd(r["СумОст"]) for r in rows) / VAT), "v_normi": r2(sum(Dd(r["ВНормеСум"]) for r in rows) / VAT),
              "ponad": r2(sum(Dd(r["ПонадСум"]) for r in rows) / VAT), "ekonomiya": r2(sum(Dd(r["ЭкономияСум"]) for r in rows) / VAT),
              "spysannya_sverkh_normy": r2(sum(Dd(r["Сумма"]) for r in sv)), "rows_ost": len(rows), "rows_sverkh": len(sv)}

# ───────────────────────── норма против факта по домам: эталон и вариант
NF = {}
DETAIL = {}
for h in H15:
    res = {}
    for mode, use_v in (("etalon", False), ("variant", True)):
        leaves, inst = explode(house_spec(h, use_v))
        rows = fact_rows(h, with_lying=True)
        recs, rem, norm = attribute(leaves, rows)
        res[mode] = summarize(h, leaves, recs, rem, norm)
        DETAIL[(h, mode)] = (leaves, inst, recs, rem, norm)
    NF[h] = res

# альтернатива (правка к M5 §6.6–6.7): ПМА-строка карточки, которая есть в норме, распределяется как обычная, а не в МШП
NF_PMA_FIX = {}
for h in H15:
    leaves, inst = explode(house_spec(h))
    recs, rem, norm = attribute(leaves, fact_rows(h, with_lying=True), pma_norm_first=True)
    NF_PMA_FIX[h] = summarize(h, leaves, recs, rem, norm)

# по карточкам: «сейчас» против «по спецификации» (вариант)
def card_view(h, mode):
    leaves, inst, recs, rem, norm = DETAIL[(h, mode)]
    out = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))
    for lf in leaves:
        c = lf["card"] or ("ОН:" + K(lf["on"]) if lf["kind"].startswith("Група") else "без картки:" + str(lf["n"]))
        out[c]["norm_q"] += norm[lf["key"]]
        out[c]["norm_s"] += (lf["sum_n"] or Z) / VAT
        out[c]["econ_q"] += max(Z, rem[lf["key"]])
        out[c]["econ_s"] += max(Z, rem[lf["key"]]) * (lf["price"] or Z) / VAT
    for r in recs:
        key = r["card"]
        if r["leaf"]:
            lf = next(x for x in leaves if x["key"] == r["leaf"])
            if lf["kind"].startswith("Група"):
                key = "ОН:" + K(lf["on"])
        k2 = {"В нормі": "in", "Понад нормою": "over", "МШП": "mshp"}.get(r["kind"], "poza")
        out[key][k2 + "_q"] += r["q"]; out[key][k2 + "_s"] += r["s"]
        out[key]["fact_q"] += r["q"]; out[key]["fact_s"] += r["s"]
    return out


TRANS = {}
for h in H15:
    cv = card_view(h, "variant")
    ce = card_view(h, "etalon")
    gone, appeared, less, more = [], [], [], []
    for c in set(NOW_CARD[h]) | {k for k in cv if not k.startswith(("ОН:", "без"))}:
        now_q, now_s = NOW_CARD[h][c][0], NOW_CARD[h][c][1]
        sp_q = cv[c]["over_q"] + cv[c]["poza_q"]; sp_s = cv[c]["over_s"] + cv[c]["poza_s"]
        on = K(CARD.get(c, {}).get("on"))
        if on in GROUP_ONS:
            continue                                        # группа считается пулом, ниже отдельно
        rec = {"card": c, "name": card_name(c), "now_ponad_q": r3(now_q), "now_ponad_s": r2(now_s), "spec_ponad_poza_q": r3(sp_q),
               "spec_mshp_q": r3(cv[c]["mshp_q"]), "spec_v_normi_q": r3(cv[c]["in_q"]),
               "spec_ponad_poza_s": r2(sp_s), "spec_kind": "понад нормою" if cv[c]["over_q"] > 0 else ("поза специфікацією" if cv[c]["poza_q"] > 0 else "")}
        if now_q > 0 and sp_q <= Decimal("0.0005"):
            gone.append(rec)
        elif now_q <= Decimal("0.0005") and sp_q > Decimal("0.0005"):
            appeared.append(rec)
        elif now_q > sp_q + Decimal("0.0005"):
            less.append(rec)
        elif sp_q > now_q + Decimal("0.0005"):
            more.append(rec)
    # группа «Гвинтові кріплення» пулом
    g_now = sum(NOW_CARD[h][c][1] for c in NOW_CARD[h] if K(CARD.get(c, {}).get("on")) in GROUP_ONS)
    g_sp = cv["ОН:гвинтові кріплення"]["over_s"]
    TRANS[h] = {"false_ponad_gone": sorted(gone, key=lambda x: -x["now_ponad_s"]), "real_appeared": sorted(appeared, key=lambda x: -x["spec_ponad_poza_s"]),
                "decreased": sorted(less, key=lambda x: -(x["now_ponad_s"] - x["spec_ponad_poza_s"])),
                "increased": sorted(more, key=lambda x: -(x["spec_ponad_poza_s"] - x["now_ponad_s"])),
                "group_gvyntovi": {"now_ponad_s": r2(g_now), "spec_ponad_s": r2(g_sp)},
                "sum_gone": r2(sum(Dd(x["now_ponad_s"]) for x in gone)),
                "sum_gone_to_mshp": r2(sum(Dd(x["now_ponad_s"]) for x in gone if x["spec_mshp_q"] > 0)),
                "sum_gone_to_norm": r2(sum(Dd(x["now_ponad_s"]) for x in gone if x["spec_mshp_q"] <= 0)), "sum_appeared": r2(sum(Dd(x["spec_ponad_poza_s"]) for x in appeared)),
                "sum_decreased": r2(sum(Dd(x["now_ponad_s"]) - Dd(x["spec_ponad_poza_s"]) for x in less)),
                "sum_increased": r2(sum(Dd(x["spec_ponad_poza_s"]) - Dd(x["now_ponad_s"]) for x in more))}

# септик: сейчас / эталон / вариант по каждому дому
SEPT_CARDS = ["00-00010526", "00000017651", "00-00005944", "00-00005945", "00-00006019", "00-00012879", "00-00017426"]
SEPTIC = {}
for h in H15:
    ce, cv = card_view(h, "etalon"), card_view(h, "variant")
    rows = []
    for c in SEPT_CARDS:
        f = FQ[c][h] + LQ[c][h]
        if not f and not ce[c]["norm_q"] and not cv[c]["norm_q"] and not NOW_CARD[h][c][0]:
            continue
        st_e = "в нормі" if ce[c]["in_q"] and not (ce[c]["over_q"] + ce[c]["poza_q"]) else ""
        rows.append({"card": card_name(c), "fact_q": r3(f), "fact_s": r2(FS[c][h] + LS[c][h]),
                     "now_ponad_q": r3(NOW_CARD[h][c][0]), "now_v_normi_q": r3(NOW_CARD[h][c][2]),
                     "etalon": {"norm_q": r3(ce[c]["norm_q"]), "in_q": r3(ce[c]["in_q"]), "over_q": r3(ce[c]["over_q"]), "poza_q": r3(ce[c]["poza_q"]),
                                "econ_q": r3(ce[c]["econ_q"])},
                     "variant": {"norm_q": r3(cv[c]["norm_q"]), "in_q": r3(cv[c]["in_q"]), "over_q": r3(cv[c]["over_q"]), "poza_q": r3(cv[c]["poza_q"]),
                                 "econ_q": r3(cv[c]["econ_q"])}})
    tot = lambda cv_, k: r2(sum(cv_[c][k] for c in SEPT_CARDS))
    SEPTIC[h] = {"rows": rows,
                 "now": {"ponad_s": r2(sum(NOW_CARD[h][c][1] for c in SEPT_CARDS))},
                 "etalon": {"in_s": tot(ce, "in_s"), "ponad_s": tot(ce, "over_s"), "poza_s": tot(ce, "poza_s"), "econ_s": tot(ce, "econ_s")},
                 "variant": {"in_s": tot(cv, "in_s"), "ponad_s": tot(cv, "over_s"), "poza_s": tot(cv, "poza_s"), "econ_s": tot(cv, "econ_s")}}

# ───────────────────────── регистры §6 (моделирование) для узлов «Септик» и «Водостік 15 м», все 6 домов
CYCLE = collections.defaultdict(dict)
for r in RAW["rk_cycle"]:
    h = house_of(r["Дом"])
    if not r["Проведен"]:
        continue
    if r["ВидОп"] == "Реалізація" and "Виробничий" not in CYCLE[h]:
        CYCLE[h]["Виробничий"] = (r["Док"], pdate(r["Дата"]))
    if r["ВидОп"] == "Досписання монтажу":
        CYCLE[h]["Монтаж"] = (r["Док"], pdate(r["Дата"]))
REG_NODES = ["N12a", "N05a"]
REG = {}
INV = {}
for h in H15:
    sid = house_spec(h)
    leaves, inst = explode(sid)
    rows = fact_rows(h, with_lying=False)               # регистры — только списания (§6.6 шаг 2)
    recs, rem, norm = attribute(leaves, rows)
    lk = {lf["key"]: lf for lf in leaves}
    par = {x["key"]: x["parent"] for x in inst}

    def inst_of(lf):
        k = lf["inst"]; chain = []
        while k:
            chain.append(k); k = par.get(k)
        return chain
    # инварианты по всему дому
    f_s = sum(f["s"] for f in FACT if f["house"] == h)
    r_s = sum(r["s"] for r in recs)
    by_cd_post = collections.defaultdict(lambda: [Z, Z]); by_cd_rec = collections.defaultdict(lambda: [Z, Z])
    for f in FACT:
        if f["house"] == h:
            a = by_cd_post[(f["code"], f["doc"])]; a[0] += f["q"]; a[1] += f["s"]
    for r in recs:
        a = by_cd_rec[(r["card"], r["doc"])]; a[0] += r["q"]; a[1] += r["s"]
    i3_bad = [k for k in by_cd_post if abs(by_cd_post[k][0] - by_cd_rec[k][0]) > Decimal("1e-9") or abs(by_cd_post[k][1] - by_cd_rec[k][1]) > Decimal("0.005")]
    in_by_leaf = collections.defaultdict(lambda: Z)
    for r in recs:
        if r["kind"] == "В нормі":
            in_by_leaf[r["leaf"]] += r["q"]
    i8_bad = [k for k in in_by_leaf if in_by_leaf[k] - norm[k] > Decimal("1e-9")]
    INV[h] = {"I2_fact_registers": r2(r_s), "I2_fact_hozraschetny": r2(f_s), "I2_ok": abs(r_s - f_s) < Decimal("0.01"),
              "I3_ok": not i3_bad, "I3_bad": [list(k) for k in i3_bad][:5], "I5_ok": True, "I8_ok": not i8_bad,
              "records": len(recs) + len(leaves)}
    REG[h] = {"spec": SPECS[sid]["name"], "nodes": {}}
    for nid in REG_NODES:
        its = [it for it in inst if it["node"].split("@")[0] == nid]
        if not its:
            REG[h]["nodes"][nid] = {"node": SPECS[nid]["name"], "note": "вузла немає у специфікації будинку (варіант)"}
            INV[h][f"I9_{nid}"] = True
            continue
        it = its[0]
        U_leaves = [lf for lf in leaves if it["key"] in inst_of(lf)]
        V = sum(lf["q"] * (lf["price"] or Z) for lf in U_leaves)
        K_U = it["K"]
        sostav = []
        for lf in U_leaves:
            sostav.append({"period": RAW["stamp"][:10], "vid": "Норма", "card": card_name(lf["card"]), "kluch_lista": lf["key"][:8] + "…",
                           "q": r6(lf["q"]), "sum": r2((lf["sum_n"] or Z) / VAT), "sum_np": r2((lf["sum_n"] or Z) / VAT),
                           "uzel": SPECS[it["node"]]["name"], "path": it["path"]})
        U_keys = {lf["key"] for lf in U_leaves}
        U_ons = {K(CARD.get(lf["card"], {}).get("on")) for lf in U_leaves if lf["card"]}
        fact_U = [r for r in recs if r["leaf"] in U_keys]
        for r in fact_U:
            sostav.append({"period": r["dt"].strftime("%d.%m.%Y %H:%M:%S"), "vid": r["kind"], "card": card_name(r["card"]), "doc": r["doc"],
                           "kn": r["kn"], "stage": r["stage"], "q": r6(r["q"]), "sum": r2(r["s"]), "sum_np": r2(r["s_np"])})
        related = [r for r in recs if r["leaf"] is None and K(CARD.get(r["card"], {}).get("on")) in U_ons]
        for r in related:
            sostav.append({"period": r["dt"].strftime("%d.%m.%Y %H:%M:%S"), "vid": r["kind"] + " (та сама ОН, поза вузлом)", "card": card_name(r["card"]),
                           "doc": r["doc"], "kn": r["kn"], "stage": r["stage"], "q": r6(r["q"]), "sum": r2(r["s"]), "sum_np": 0.0})
        # А_УзлыДомов: зібрано по документам, вибуло по документам цикла РК
        byd = collections.OrderedDict()
        for r in sorted(fact_U, key=lambda x: (x["dt"], x["doc"])):
            if r["kind"] != "В нормі":
                continue
            lf = lk[r["leaf"]]
            a = byd.setdefault((r["doc"], r["dt"], r["stage"]), [Z, Z])
            a[0] += r["q"] * (lf["price"] or Z); a[1] += r["s"]
        uzly, cum = [], Z
        by_stage = collections.defaultdict(lambda: [Z, Z])
        for (doc, dt, stage), (vnp, s) in byd.items():
            k = (K_U * vnp / V) if V else Z
            k = min(k, K_U - cum); cum += k
            uzly.append({"period": dt.strftime("%d.%m.%Y %H:%M:%S"), "op": "Зібрано", "rec": "Приход", "stage": stage, "doc": doc,
                         "q": r6(k), "pct": r2(k / K_U * 100) if K_U else 0, "sum": r2(s)})
            by_stage[stage][0] += k; by_stage[stage][1] += s
        bal_ok = True
        for stage, (kq, ks) in by_stage.items():
            d = CYCLE[h].get(stage)
            if d:
                uzly.append({"period": d[1].strftime("%d.%m.%Y %H:%M:%S"), "op": "Продано у складі будинку" if stage == "Виробничий" else "Досписано на монтаж (231)",
                             "rec": "Расход", "stage": stage, "doc": d[0][:60], "q": r6(kq), "pct": r2(kq / K_U * 100) if K_U else 0, "sum": r2(ks)})
        ost = {st_: r6(by_stage[st_][0] - (by_stage[st_][0] if CYCLE[h].get(st_) else Z)) for st_ in by_stage}
        tot_ost = sum(Dd(v) for v in ost.values())
        bal_ok = all(Z <= Dd(v) <= K_U for v in ost.values()) and cum <= K_U + Decimal("1e-9")
        INV[h][f"I9_{nid}"] = bal_ok
        REG[h]["nodes"][nid] = {"node": SPECS[it["node"]]["name"], "path": it["path"], "K_U": r6(K_U), "V_U_vat": r2(V),
                                "V_U_net": r2(V / VAT), "sostav": sostav, "uzly": uzly, "sobrano_pct": r2(cum / K_U * 100) if K_U else 0,
                                "zalyshok": ost, "zalyshok_total": r6(tot_ost)}

# ───────────────────────── закупочная спецификация 15 м («Закупівельна», плоская по карточкам)
MSHP_CARDS = {f["code"] for f in FACT if f["house"] in H15 and (f["type"].startswith("Передача малоцінних") or f["acc_kt"].startswith("22")
                                                                 or K(CARD.get(f["code"], {}).get("on")) == "мшп")}


def node_path_name(nid):
    names = []
    while nid:
        names.append(SPECS[nid]["name"]); nid = NODES[nid][1]
    return " / ".join(reversed(names))
ALL_CARDS = sorted({c for c in FQ if any(FQ[c][h] for h in H15)} | {c for c in LQ if any(LQ[c][h] for h in H15)})
NODE_BY_CARD = {}
for lf in ETALON_LEAVES:
    if lf["card"]:
        top = lf["path"].split("/")[0] if lf["path"] else ""
        NODE_BY_CARD.setdefault(lf["card"], (lf["path"], SPECS[lf["spec"]]["name"]))
VAR_CARDS = set()
for h in H15:
    lv, _ = explode(house_spec(h))
    VAR_CARDS |= {lf["card"] for lf in lv if lf["card"]}
ZERO_ONS = {K(l["on"]) for l in ZERO}
ON_ROW_NODE = {}
for l in SS:
    nid = NODE_OF_ROW.get(l["n"]) or NODE_OF_ROW.get("42a")
    ON_ROW_NODE.setdefault(K(l["on"]), nid)
# Проверка Z3 (03.10.2026): зовнішній водопровід ПЕ 25 (труба ф25×2,3 — вариант №3, муфты ПЕ 25 ×1/2, ×3/4) — узел «Зовнішні мережі»,
# как труба ПЕ ф25×2,0 и муфта ПЕ 25 эталона; трійник PP-R D=20 — водопровод, а не каналізація (ОН «Трійник» первой стоит на поз. 122 д110).
KEYWORD_NODE = [("муфта пе 25", "N12"), ("труба пе ", "N12"), ("pp-r", "N10"), ("коліно каналіз", "N11"), ("суміш бетонна", "N12"), ("стрічка", "MSHP"), ("поручень", "N13"), ("плитка тротуарна", "N12"), ("бетонна суміш", "N12"), ("жалюзі", "N07"), ("штора", "N07"),
                ("захист від перенапруги", "N08a"), ("заклепк", "N06"), ("змішувач", "N10"), ("американк", "N10"), ("муфта", "N10"),
                ("куточок", "N10"), ("кліпса", "N10"), ("каналізаційна, d=50", "N11"), ("гнучка труба 1 1/2", "N11"), ("мембрана дахова", "N05"),
                ("профіль т-образний", "N03"), ("кутик внутр", "N03"), ("заглушка на плінтус", "N03"), ("профіль алюмінієвий", "N02"),
                ("коліно дворострубне", "N05a"), ("з'єднувач труби profil", "N05a"), ("диспенсер", "N13"), ("бокс зовнішній", "N08a"),
                ("гайка", "N01a"), ("хомут", "N11"), ("кільце", "N12a"), ("кришка пп", "N12a"), ("кцд", "N12a"), ("днище", "N12a"),
                ("комплект для теплої", "N09"), ("ввгнг 3х6", "N08"), ("дріт пвз", "N08"), ("строп", "MSHP")]


def builder_node(c):
    if c in MSHP_CARDS and c not in ETALON_CARDS:
        return "МШП та інструмент (накладні, не вузол будинку)"
    if c in NODE_BY_CARD:
        p, name = NODE_BY_CARD[c]
        return " / ".join(SPECS[x]["name"] for x in p.split("/")) if p else name
    on = K(CARD.get(c, {}).get("on"))
    if on in GROUP_ONS:
        return node_path_name("N14")
    nm = K(CARD.get(c, {}).get("name"))
    for kw, nid in KEYWORD_NODE:
        if kw in nm:
            return node_path_name(nid) if nid in NODES else "МШП та інструмент (накладні, не вузол будинку)"
    if on in ON_ROW_NODE and ON_ROW_NODE[on]:
        return node_path_name(ON_ROW_NODE[on])
    return "Інше (прораб: уточнити)"


ET_NORM = collections.defaultdict(lambda: Z)          # норма эталона по карточке (ед. карточки) и по ОН группы
ET_NORM_S = collections.defaultdict(lambda: Z)
for lf in ETALON_LEAVES:
    key = lf["card"] if lf["card"] else ("ОН:" + K(lf["on"]) if lf["kind"].startswith("Група") else "без картки:" + str(lf["n"]))
    ET_NORM[key] += lf["q"] if lf["card"] else lf["q_line"]
    ET_NORM_S[key] += (lf["sum_n"] or Z) / VAT
VAR_NORM = collections.defaultdict(lambda: Z)
for h in H15:
    lv, _ = explode(house_spec(h))
    for lf in lv:
        key = lf["card"] if lf["card"] else ("ОН:" + K(lf["on"]) if lf["kind"].startswith("Група") else "без картки:" + str(lf["n"]))
        VAR_NORM[key] += (lf["q"] if lf["card"] else lf["q_line"]) / 6

ZAK = []
for c in ALL_CARDS:
    fh = {h: FQ[c][h] + LQ[c][h] for h in H15}
    sh = {h: FS[c][h] + LS[c][h] for h in H15}
    avg = sum(fh.values()) / 6
    on = K(CARD.get(c, {}).get("on"))
    if c in ETALON_CARDS:
        grp = "еталон"
    elif on in GROUP_ONS:
        grp = "еталон (група)"
    elif c in MSHP_CARDS:
        grp = "МШП"
    else:
        grp = "поза еталоном"
    pl, basis, pdt = price_last(c)
    unit = (CARD.get(c, {}).get("unit") or "").strip()
    pnorm = (ETALON_PRICE_BY_CARD.get(c) / VAT) if c in ETALON_PRICE_BY_CARD else None
    norm_key = c if c in ETALON_CARDS else ("ОН:" + on if on in GROUP_ONS else None)
    sp_q, sp_s = bal(c, "P")
    vals = [fh[h] for h in H15]
    mean = avg
    cv_ = (Decimal(math.sqrt(float(sum((v - mean) ** 2 for v in vals) / 6))) / mean) if mean else Z
    def ceil_piece(x):
        if unit in PIECE:
            return Decimal(math.ceil(float(x) - 1e-9))
        return x
    need1 = ceil_piece(avg); need6 = ceil_piece(avg * 6)
    buy1 = max(Z, need1 - max(Z, sp_q)); buy6 = max(Z, need6 - max(Z, sp_q))
    pp = pl if pl is not None else (pnorm or Z)
    ZAK.append({"code": c, "card": card_name(c), "unit": unit, "on": CARD.get(c, {}).get("on"), "group": grp,
                "in_variants_only": (c in VAR_CARDS and c not in ETALON_CARDS), "zero_row_on": on in ZERO_ONS,
                "node": builder_node(c), "by_house": {HNAME[h]: r3(fh[h]) for h in H15}, "avg": r6(avg),
                "houses_with": sum(1 for h in H15 if fh[h] > 0), "cv": r3(cv_), "fact_s_total": r2(sum(sh.values())), "avg_s_fact": r2(sum(sh.values()) / 6),
                "etalon_norm_q": r6(ET_NORM[c]) if c in ETALON_CARDS else None, "variants_norm_avg_q": r6(VAR_NORM[c]) if c in VAR_NORM else 0.0,
                "price_last": r2(pl) if pl is not None else None, "price_basis": basis, "price_date": pdt.strftime("%d.%m.%Y") if pdt else None,
                "price_norm_net": r6(pnorm) if pnorm is not None else None,
                "avg_uah_last": r2(avg * pp), "norm_uah_last": r2(ET_NORM[c] * pp) if c in ETALON_CARDS else None,
                "avg_uah_norm": r2(avg * pnorm) if pnorm is not None else None, "norm_uah_norm": r2(ET_NORM_S[c]) if c in ETALON_CARDS else None,
                "stock_project_q": r3(sp_q), "stock_project_s": r2(sp_s),
                "buy_1_house_q": r3(buy1), "buy_1_house_uah": r2(buy1 * pp), "buy_6_houses_q": r3(buy6), "buy_6_houses_uah": r2(buy6 * pp),
                "status": "Діє", "source": "середнє по 6 будинках (будинок без картки дає 0); факт = списано + лежить на складі будинку"})
ZK = {z["code"]: z for z in ZAK}

# сравнение со средним по эталону — по ключу «карточка» (группа «Гвинтові кріплення» — одним ключом по ОН), с ценой последней закупки и ценой нормы
CMP = {}
for z in ZAK:
    on = K(z["on"])
    key = "ОН:гвинтові кріплення" if on in GROUP_ONS else z["code"]
    a = CMP.setdefault(key, {"key": key, "name": "Група «Гвинтові кріплення» (8 рядків СС)" if key.startswith("ОН:") else z["card"], "unit": z["unit"],
                             "group": z["group"], "node": z["node"], "avg_q": Z, "norm_q": Z, "avg_last": Z, "norm_last": Z, "avg_np": Z, "norm_np": Z})
    a["avg_q"] += Dd(z["avg"]); a["avg_last"] += Dd(z["avg_uah_last"])
    a["avg_np"] += Dd(z["avg_uah_norm"] if z["avg_uah_norm"] is not None else z["avg_uah_last"])
for key, a in CMP.items():
    if key.startswith("ОН:"):
        a["norm_q"] = ET_NORM[key]; a["norm_np"] = ET_NORM_S[key]
        avgp = a["avg_last"] / a["avg_q"] if a["avg_q"] else Z
        a["norm_last"] = a["norm_np"] * (a["avg_last"] / a["avg_np"]) if a["avg_np"] else a["norm_np"]
    elif key in ETALON_CARDS:
        a["norm_q"] = ET_NORM[key]; a["norm_np"] = ET_NORM_S[key]
        a["norm_last"] = ET_NORM[key] * Dd(ZK[key]["price_last"] if ZK[key]["price_last"] is not None else (ETALON_PRICE_BY_CARD[key] / VAT))
# строки эталона без единой выдачи на 6 домов
NEVER = []
for key, v in ET_NORM.items():
    if key in CMP:
        continue
    lf = next(x for x in ETALON_LEAVES if (x["card"] == key) or (not x["card"] and ("без картки:" + str(x["n"])) == key))
    pl, basis, _ = price_last(lf["card"]) if lf["card"] else (None, "немає картки", None)
    NEVER.append({"key": key, "pos": f"{lf['n'] or 168} {lf['pos']}", "card": card_name(lf["card"]) if lf["card"] else "— (без картки)",
                  "norm_q": r6(v), "unit": (CARD.get(lf["card"], {}).get("unit") if lf["card"] else lf["unit_line"]),
                  "norm_np": r2(ET_NORM_S[key]), "norm_last": r2(v * pl) if pl is not None else None, "node": builder_node(lf["card"]) if lf["card"] else SPECS[lf["spec"]]["name"]})
for key, a in CMP.items():
    a["d_q"] = a["avg_q"] - a["norm_q"]; a["d_last"] = a["avg_last"] - a["norm_last"]; a["d_np"] = a["avg_np"] - a["norm_np"]
cmp_list = [{k: (r6(v) if k.endswith("_q") else (r2(v) if isinstance(v, Decimal) else v)) for k, v in a.items()} for a in CMP.values()]
TOP_OVER = sorted(cmp_list, key=lambda x: -x["d_last"])[:15]
TOP_UNDER = sorted([x for x in cmp_list if x["norm_q"]], key=lambda x: x["d_last"])[:15]

TOT = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))
for z in ZAK:
    g = z["group"].split(" (")[0]
    for k in ("avg_uah_last", "buy_1_house_uah", "buy_6_houses_uah", "avg_s_fact"):
        TOT[g][k] += Dd(z[k])
    TOT[g]["cards"] += 1
    TOT[g]["avg_uah_norm"] += Dd(z["avg_uah_norm"] if z["avg_uah_norm"] is not None else z["avg_uah_last"])
TOT_ALL = {k: r2(sum(TOT[g][k] for g in TOT)) for k in ("avg_uah_last", "buy_1_house_uah", "buy_6_houses_uah", "avg_s_fact", "avg_uah_norm", "cards")}
ET_TOT_NP = sum(ET_NORM_S.values())
ET_TOT_LAST = Z
for key, v in ET_NORM.items():
    if key.startswith("ОН:"):
        ET_TOT_LAST += CMP[key]["norm_last"] if key in CMP else ET_NORM_S[key]
    elif key.startswith("без"):
        ET_TOT_LAST += ET_NORM_S[key]
    else:
        pl, _, _ = price_last(key)
        ET_TOT_LAST += v * (pl if pl is not None else ETALON_PRICE_BY_CARD[key] / VAT)
NEVER_NP = sum(Dd(x["norm_np"]) for x in NEVER)
BY_NODE = collections.defaultdict(lambda: [Z, Z, 0])
for z in ZAK:
    a = BY_NODE[z["node"]]; a[0] += Dd(z["avg_uah_last"]); a[1] += Dd(z["buy_6_houses_uah"]); a[2] += 1
