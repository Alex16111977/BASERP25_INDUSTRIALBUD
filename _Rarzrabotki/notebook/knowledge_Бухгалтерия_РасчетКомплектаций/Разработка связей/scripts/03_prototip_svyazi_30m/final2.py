# -*- coding: utf-8 -*-
"""Часть 2: показательные позиции."""
from final import *

EXAMPLES = [
    {"position": "Комплект септика", "lines": [170], "on": ["комплект септика"], "now_extra": [], "map_extra": []},
    {"position": "Гайка М48 / М24 / М12", "lines": [8, 10, 18], "on": ["гайка"], "now_extra": [], "map_extra": []},
    {"position": "Шайба / гровер М12", "lines": [19, 20], "on": ["шайба"], "now_extra": [], "map_extra": []},
    {"position": "Держак ринви / кріплення ринви і труби", "lines": [72, 73, 129, 130],
     "on": ["кріплення ринви", "держак ринви", "комплект кріплення труб"], "now_extra": ["00000000353"], "map_extra": []},
    {"position": "З'єднувач ринви / труби", "lines": [74], "on": ["з'єднувач ринви"], "now_extra": [], "map_extra": []},
    {"position": "Дозатор / диспенсер (мило, рушники)", "lines": [157, 158], "on": ["диспенсер рідкого мила", "диспенсер рушників"],
     "now_extra": [], "map_extra": []},
    {"position": "Кутики: сталевий 32×32×3 і декоративні / меблевий", "lines": [13, 53, 60, 61],
     "on": ["кутики сталеві", "кутики (металопрокат)", "кутики декоративні", "плінтус на стільницю"], "now_extra": [], "map_extra": []},
    {"position": "Кабель ВВГ 3х4 / 3х2,5 / 3х1,5", "lines": [81, 82, 83], "on": ["кабель"], "now_extra": [], "map_extra": ["00-00019514"]},
    {"position": "Клей монтажний і піна-клей", "lines": [58, 62, 79], "on": ["клей", "піна клей"], "now_extra": ["00-00000594"],
     "map_extra": ["00-00021278"]},
    {"position": "Муфта ПЕ 25 / з'єднувач компресійний", "lines": [141], "on": ["з'єднувач компресійний", "муфта"], "now_extra": [],
     "map_extra": ["00-00005140", "00-00005141", "00-00022599"]},
]


def st_done(over_q, over_g, econ_q, econ_g, unit):
    p = []
    if over_q > Decimal("0.0005") or over_g > Decimal("0.005"):
        p.append((f"понад {fq(over_q)} {unit} = " if over_q > Decimal("0.0005") else "понад ") + f"{fmt(over_g)} грн")
    if econ_q > Decimal("0.0005"):
        p.append(f"економія {fq(econ_q)} {unit} = {fmt(econ_g)} грн")
    return "; ".join(p) if p else "в нормі"


def ex_now(e, h):
    unit = "/".join(sorted({unit_of_on(k) for k in e["on"] if unit_of_on(k)})) or "шт"
    N = sum(ON_NORM.get(k, [Z, Z])[0] for k in e["on"])
    get = lambda k, f: (NOW[h].get(k) or {}).get(f, Z)
    F = sum(get(k, "F") for k in e["on"]); Sh = sum(get(k, "Sh") for k in e["on"])
    oq = sum(get(k, "over_q") for k in e["on"]); og = sum(get(k, "over_g") for k in e["on"])
    eq = sum(get(k, "econ_q") for k in e["on"]); eg = sum(get(k, "econ_g") for k in e["on"])
    if h in DONE:   # ОН с нормой, но без фактов (нет записи в NOW) -> вся норма в экономии
        for k in e["on"]:
            if k not in NOW[h] and ON_NORM.get(k):
                eq += ON_NORM[k][0]; eg += ON_NORM[k][1]
    parts = []
    for k in e["on"]:
        r = NOW[h].get(k)
        if r is None and not ON_NORM.get(k):
            continue
        if r is None:
            r = {"N": ON_NORM[k][0], "F": Z, "Sh": Z, "over_q": Z, "over_g": Z, "econ_q": ON_NORM[k][0] if h in DONE else Z,
                 "econ_g": ON_NORM[k][1] if h in DONE else Z, "D": Z if h in DONE else ON_NORM[k][0]}
        parts.append({"on": ON_NAME.get(k, k), "unit": unit_of_on(k) or "(норми немає)", "norm": f3(r["N"]), "fact": f3(r["F"]),
                      "lying": f3(r["Sh"]), "over_q": f3(r["over_q"]), "over_uah": f2(r["over_g"]), "econ_q": f3(r["econ_q"]),
                      "econ_uah": f2(r["econ_g"]), "treba": f3(r["D"]) if h not in DONE else 0})
    for c in e["now_extra"]:
        cq = cons[c][h][0] + st[c]["H:" + h][0]
        cg = BY_CARD_OVER.get(("now", h, c), Z)
        if cq:
            parts.append({"card_in_other_on": CARD[c]["name"], "on_card": CARD[c]["on"], "fact": f3(cons[c][h][0]),
                          "lying": f3(st[c]["H:" + h][0]), "over_uah": f2(cg)})
            og += cg
    rec = {"norm": f3(N), "fact": f3(F), "lying": f3(Sh), "unit": unit}
    if h in DONE:
        rec["status"] = st_done(oq, og, eq, eg, unit)
    else:
        D = sum((NOW[h].get(k) or {"D": ON_NORM.get(k, [Z])[0]})["D"] for k in e["on"])
        mv = sum(NOW_BUY[k]["byh"][h]["move"] for k in e["on"] if k in NOW_BUY)
        by = sum(NOW_BUY[k]["byh"][h]["buy"] for k in e["on"] if k in NOW_BUY)
        bu = sum(NOW_BUY[k]["byh"][h]["buy"] * NOW_BUY[k]["price_ss"] for k in e["on"] if k in NOW_BUY)
        s = f"треба {fq(D)} {unit}: зі складу проєкту {fq(mv)}, докупити {fq(by)} = {fmt(bu)} грн"
        if h == "30-3":
            s = f"списано {fq(F)}, лежить {fq(Sh)}; " + s + (f"; понад {fq(oq)} = {fmt(og)} грн" if oq > 0 else "")
        rec["status"] = s
        rec.update({"treba": f3(D), "move": f3(mv), "buy": f3(by), "buy_uah": f2(bu)})
    rec["over_uah"] = f2(og); rec["econ_uah"] = f2(eg)
    rec["by_on"] = parts
    return rec


def ex_map(e, h):
    parts, og_t, eg_t, buy_t = [], Z, Z, Z
    for n in e["lines"]:
        r = MAPPED[h][n]; l = LINE[n]
        unit = "компл." if n == SEPTIC_LINE else l["unit"]
        p = {"n": n, "line": l["text"], "unit": unit, "norm": f3(r["N"]), "fact": f3(r["F"]), "lying": f3(r["Sh"]),
             "over_q": f3(r["over_q"]), "over_uah": f2(r["over_g"]), "econ_q": f3(r["econ_q"]), "econ_uah": f2(r["econ_g"])}
        if n == SEPTIC_LINE:
            p["components_fact"] = r["comp"]; p["extra_parts"] = r["extra_parts"]
        if h in DONE:
            if n == SEPTIC_LINE and r["extra_parts"] and r["F"] >= 1:
                p["status"] = ("1 комплект у нормі; зайве: " + ", ".join(f"{k} {fq(v)} шт" for k, v in r["extra_parts"].items())
                               + f" = {fmt(r['over_g'])} грн")
            else:
                p["status"] = st_done(r["over_q"] if n != SEPTIC_LINE else Z, r["over_g"], r["econ_q"], r["econ_g"], unit)
        else:
            b = MAP_BUY[n]["byh"][h]
            bu = b["buy_g_ss"] if n == SEPTIC_LINE else b["buy"] * MAP_BUY[n]["price_ss"]
            s = f"треба {fq(r['D'])}: зі складу {fq(b['move'])}, докупити {fq(b['buy'])} = {fmt(bu)} грн"
            if n == SEPTIC_LINE and any(b["buy_parts"].values()):
                s += " (деталі: " + ", ".join(f"{k} {fq(v)}" for k, v in b["buy_parts"].items() if v) + " шт)"
            if h == "30-3":
                s = f"списано {fq(r['F'])}, лежить {fq(r['Sh'])}; " + s + (f"; понад {fq(r['over_q'])} = {fmt(r['over_g'])} грн" if r["over_q"] > 0 else "")
            p["status"] = s; p["treba"] = f3(r["D"]); p["move"] = f3(b["move"]); p["buy"] = f3(b["buy"]); p["buy_uah"] = f2(bu)
            buy_t += bu
        og_t += r["over_g"]; eg_t += r["econ_g"]
        parts.append(p)
    for c in e["map_extra"]:
        cq = cons[c][h][0] + st[c]["H:" + h][0]
        if cq:
            g = cons[c][h][2] + st[c]["H:" + h][2]
            parts.append({"card_outside_ss": CARD[c]["name"], "fact": f3(cons[c][h][0]), "lying": f3(st[c]["H:" + h][0]), "over_uah": f2(g),
                          "status": f"поза СС: {fq(cq)} = {fmt(g)} грн"})
            og_t += g
    rec = {"norm": "; ".join(f"{p['line']} {fq(p['norm'])} {p['unit']}" for p in parts if "line" in p),
           "fact": "; ".join(f"{fq(p['fact'])}" for p in parts if "line" in p),
           "status": " | ".join((p.get("line") or p.get("card_outside_ss"))[:30] + ": " + p["status"] for p in parts),
           "over_uah": f2(og_t), "econ_uah": f2(eg_t), "by_line": parts}
    if h not in DONE:
        rec["buy_uah"] = f2(buy_t)
    return rec


examples = []
for e in EXAMPLES:
    cards, seen = [], set()
    for n in e["lines"]:
        if n == SEPTIC_LINE:
            cards.append({"code": KIT, "name": CARD[KIT]["name"], "on_card": CARD[KIT]["on"], "role": "Виключена", "coef": 0,
                          "note": MAP[KIT]["note"]}); seen.add(KIT)
        for c, mth, lk in sorted(BY_LINE.get(n, []), key=lambda x: x[2]["prio"]):
            if c in seen:
                continue
            seen.add(c)
            cards.append({"code": c, "name": CARD[c]["name"], "on_card": CARD[c]["on"], "role": lk["role"], "coef": float(lk["coef"]),
                          "prio": lk["prio"], "line": n, "method": mth, "conf": lk["conf"], "note": lk["note"]
                          + (f"; {lk['per_kit']} шт на комплект" if lk.get("per_kit") and "на комплект" not in lk["note"] else "")})
    for k in e["on"]:
        for c in ON_CARDS.get(k, []):
            if c not in seen and (any(cons[c][h][0] for h in HOUSES) or st[c]["P"][0]):
                seen.add(c); m = MAP[c]
                cards.append({"code": c, "name": CARD[c]["name"], "on_card": CARD[c]["on"], "role": m["kind"], "coef": 0,
                              "note": m.get("note", "") or "зв'язана з іншою строкою"})
    for c in e["now_extra"] + e["map_extra"]:
        if c not in seen:
            seen.add(c); m = MAP[c]
            cards.append({"code": c, "name": CARD[c]["name"], "on_card": CARD[c]["on"], "role": m["kind"],
                          "coef": float(m["links"][0]["coef"]) if m["links"] else 0,
                          "line": m["links"][0]["line"] if m["links"] else None,
                          "note": m.get("note", "") or (m["links"][0]["note"] if m["links"] else "")})
    by_h = [{"n": HN[h], "now": ex_now(e, h), "mapped": ex_map(e, h)} for h in HOUSES]
    fams = {fam(c["code"]) for c in cards}
    gone = appear = Z
    for h in ("30-1", "30-2", "30-3"):
        a = collections.defaultdict(Decimal); b = collections.defaultdict(Decimal)
        for (v, hh, c), g in BY_CARD_OVER.items():
            if hh == h and fam(c) in fams:
                (a if v == "now" else b)[fam(c)] += g
        for k in set(a) | set(b):
            d = a.get(k, Z) - b.get(k, Z)
            if d > 0:
                gone += d
            else:
                appear += -d
    now_buy = sum(i["uah_ss"] for i in NI if i["key"] in e["on"]); now_buy_l = sum(i["uah_last"] for i in NI if i["key"] in e["on"])
    map_buy = sum(i["uah_ss"] for i in MI if i["key"] in e["lines"]); map_buy_l = sum(i["uah_last"] for i in MI if i["key"] in e["lines"])
    now_ec = sum(Dd(x["now"]["econ_uah"]) for x in by_h); map_ec = sum(Dd(x["mapped"]["econ_uah"]) for x in by_h)
    examples.append({
        "position": e["position"],
        "ss_lines": [{"n": n, "text": LINE[n]["text"], "on_in_ss": LINE[n]["on"], "norm_per_house": f3(LINE[n]["q"]), "unit": LINE[n]["unit"],
                      "norm_uah_per_house": f2(LINE[n]["sum"]), "price_ss": f2(LINE[n]["price"])} for n in e["lines"]],
        "now_on": [ON_NAME.get(k, k) for k in e["on"]],
        "cards": cards, "by_house": by_h,
        "money": {"now_dokupyty_3_7_uah": f2(now_buy), "mapped_dokupyty_3_7_uah": f2(map_buy),
                  "now_dokupyty_3_7_uah_last_price": f2(now_buy_l), "mapped_dokupyty_3_7_uah_last_price": f2(map_buy_l),
                  "delta_dokupyty_uah": f2(map_buy - now_buy),
                  "false_ponad_gone_1_3_uah": f2(gone), "hidden_ponad_revealed_1_3_uah": f2(appear),
                  "now_ekonomiya_1_2_uah": f2(now_ec), "mapped_ekonomiya_1_2_uah": f2(map_ec)},
    })

if __name__ == "__main__":
    for x in examples:
        print("=" * 100)
        print(x["position"], json.dumps(x["money"], ensure_ascii=False))
        for b in x["by_house"]:
            print("  ", b["n"], "| NOW:", b["now"]["status"], "| MAP:", b["mapped"]["status"])
    print(json.dumps(totals, ensure_ascii=False, indent=1))
    print(json.dumps(coverage["by_method"], ensure_ascii=False, indent=1))
    print("manual", coverage["manual_lines"], coverage["manual_norm_uah_per_house"])
    print(json.dumps(coverage["fact_1_2_by_method"], ensure_ascii=False))
    for d in dokupyty_by_house:
        print(d["n"], d["now_uah"], d["mapped_uah"], d["mapped_uah_confirmed"], d["v3_G_h"])
    for m in moves[:25]:
        print(m["on"], m["now_uah"], m["v3_uah"], m["mapped_uah"], m["delta_mapped_vs_now"], m["no_fact_on_1_2"])
    for b in by_house[:3]:
        print(json.dumps(b, ensure_ascii=False))
