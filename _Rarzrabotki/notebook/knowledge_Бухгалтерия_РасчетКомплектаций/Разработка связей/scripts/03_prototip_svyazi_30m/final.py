# -*- coding: utf-8 -*-
"""Сборка demo_svyaz_ss_30m.json: прототип соответствия «строка СС ↔ карточки» и два вида по домам 30 м №1–№7."""
from views import *

OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\Мой расчет\data\demo_svyaz_ss_30m.json"
IRC = json.load(open(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\Мой расчет\data\irc_fakt.json", encoding="utf-8"))
HN = {h: "№" + h[3:] for h in HOUSES}
f2 = lambda x: r2(x)
f3 = lambda x: r3(x)


def fmt(x, n=2):
    s = f"{Dd(x).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP):,.{n}f}".replace(",", " ").replace(".", ",")
    return s


def fq(x):
    x = Dd(x)
    if x == x.to_integral_value():
        return str(int(x))
    return fmt(x, 3).rstrip("0").rstrip(",")


def is_kit_family(c):
    return c == KIT or c in KIT_CARDS or str(c).startswith("kit:")


# ─────────────────────────────── 0. дома
RK = {house(r["Подр"] + " IR") or house("IR " + r["Подр"]): r for r in RAW["rk"]}
SHIP_PLAN = {"30-3": "2026-10", "30-4": "2026-10", "30-5": "2026-10", "30-6": "2026-10", "30-7": "2026-11"}


def d10(s):
    if not s:
        return ""
    a = s.split(" ")[0].split(".")
    return f"{a[2]}-{a[1]}-{a[0]}"


houses = []
for h in HOUSES:
    Fg = sum(cons[c][h][2] for c in cons); Fn = sum(cons[c][h][1] for c in cons)
    Sg = sum(st[c]["H:" + h][2] for c in st)
    rk = RK.get(h)
    if h in DONE:
        status = "відвантажено і списано"
        ship = d10(rk["ДатаОтгр"]) if rk else ""
    elif h == "30-3":
        status = "частково: списано каркас/метал, частина матеріалів на складі дому"
        ship = f"план {SHIP_PLAN[h]} (бюджет v3); у РК №{rk['Номер'][-2:]} дата відвантаження {d10(rk['ДатаОтгр'])}" if rk else SHIP_PLAN[h]
    else:
        status = "не розпочато (фактів списання немає)"
        ship = f"план {SHIP_PLAN[h]} (бюджет v3)"
    houses.append({"n": HN[h], "key": h, "status": status, "ship_date": ship,
                   "ship_kazna": d10(rk["ДатаКазна"]) if rk and h in DONE else "",
                   "rk_doc": (f"Списання по СС №{rk['Номер']} ({rk['Статус']})" if rk else ""),
                   "written_off_uah": f2(Fg), "written_off_net": f2(Fn), "house_stock_uah": f2(Sg)})

# ─────────────────────────────── 1. соответствие по 170 строкам и покрытие
METHOD_ORDER = ["1:1 за ОН", "N:1 за ОН", "за розміром", "комплект", "аліас / пересорт", "групова норма", "без карточки"]
METHOD_TXT = {
    "1:1 за ОН": "авто: одна строка СС у загальної назви — одна картка",
    "N:1 за ОН": "авто: одна строка СС у загальної назви — кілька карток (дублі/варіанти)",
    "за розміром": "авто: кілька строк в одній загальній назві, картку знайдено за розміром у назві (М12/М24/М48, 3х2,5, WGn-2, 10А …)",
    "комплект": "склад зі службової комплектації ІБ00-000174: днище ×1, кільце ×2, люк ×1, плита ×1; картка «Комлпект септика» — виключена (напівфабрикат)",
    "аліас / пересорт": "аліас з дослідження та причин понад норму: картка під іншою загальною назвою закриває цю строку (з коефіцієнтом одиниць)",
    "групова норма": "«Гвинтові кріплення»: строки взаємозамінні, контроль — сумою по групі (як зараз)",
    "без карточки": "жодна картка не списувалась і не лежить на складі — потрібна ручна зв'язка або рішення «строка не потрібна»",
}
mapping_rows = []
cov = {m: {"lines": 0, "norm_uah_per_house": Z, "manual_check_lines": 0} for m in METHOD_ORDER}
manual_lines = []
for l in SS:
    n = l["n"]
    me = line_method(n)
    cf = line_conf(n)
    cov[me]["lines"] += 1
    cov[me]["norm_uah_per_house"] += l["sum"]
    links = []
    if n == SEPTIC_LINE:
        links.append({"code": KIT, "card": CARD[KIT]["name"], "role": "Виключена", "coef": 0, "prio": 0, "conf": "high",
                      "note": MAP[KIT]["note"]})
    if l["k"] in GROUP_ON:
        for c, m in MAP.items():
            if m["kind"] == "Групова":
                links.append({"code": c, "card": CARD[c]["name"], "role": "Групова", "coef": 1, "prio": 1, "conf": "high", "note": "пул групи"})
    for c, mth, lk in sorted(BY_LINE.get(n, []), key=lambda x: x[2]["prio"]):
        links.append({"code": c, "card": CARD[c]["name"], "role": lk["role"], "coef": float(lk["coef"]), "prio": lk["prio"],
                      "conf": lk["conf"], "note": lk["note"] + (f"; на комплект {lk['per_kit']}" if lk.get("per_kit") and "на комплект" not in lk["note"] else "")})
    need_manual = me == "без карточки" or cf in ("medium", "low")
    if need_manual:
        cov[me]["manual_check_lines"] += 1
        manual_lines.append({"n": n, "text": l["text"], "method": me, "conf": cf, "norm_uah": f2(l["sum"]),
                             "why": "немає картки" if me == "без карточки" else "довіра medium — підтвердити"})
    mapping_rows.append({"n": n, "etap": l["etap"], "on": l["on"], "text": l["text"], "unit": l["unit"], "norm_q": f3(l["q"]),
                         "norm_uah": f2(l["sum"]), "method": me, "conf": cf, "links": links})
norm_house = sum(l["sum"] for l in SS)
# покрытие факта №1–№2 (брутто списания) по способам
fact_cov = collections.defaultdict(Decimal)
for c in U:
    g = sum(cons[c][h][2] for h in DONE)
    if not g:
        continue
    m = MAP[c]
    if m["method"] == "комплект":
        key = "комплект"
    elif m["kind"] == "Групова":
        key = "групова норма"
    elif m["kind"] == "поза СС":
        key = "картки поза СС"
    elif m["kind"] == "МШП":
        key = "МШП / інструмент"
    elif m["method"] == "аліас":
        key = "аліас / пересорт"
    elif m["method"] == "за розміром":
        key = "за розміром"
    elif m["method"] in ("за ОН", "за ОН N:1"):
        key = "за ОН (1:1 / N:1)"
    else:
        key = "не визначено"
    fact_cov[key] += g
fact_tot = sum(fact_cov.values())
auto_methods = ("1:1 за ОН", "N:1 за ОН", "за розміром")
coverage = {
    "norm_uah_per_house": f2(norm_house), "lines_total": len(SS),
    "by_method": [{"method": m, "what": METHOD_TXT[m], "lines": cov[m]["lines"], "norm_uah_per_house": f2(cov[m]["norm_uah_per_house"]),
                   "share_norm_pct": round(float(cov[m]["norm_uah_per_house"] / norm_house * 100), 1),
                   "of_them_need_manual_check": cov[m]["manual_check_lines"]} for m in METHOD_ORDER],
    "auto_lines": sum(cov[m]["lines"] for m in auto_methods),
    "auto_norm_uah_per_house": f2(sum(cov[m]["norm_uah_per_house"] for m in auto_methods)),
    "special_lines": sum(cov[m]["lines"] for m in ("комплект", "аліас / пересорт", "групова норма")),
    "special_norm_uah_per_house": f2(sum(cov[m]["norm_uah_per_house"] for m in ("комплект", "аліас / пересорт", "групова норма"))),
    "manual_lines": len(manual_lines),
    "manual_norm_uah_per_house": f2(sum(Dd(x["norm_uah"]) for x in manual_lines)),
    "manual_list": manual_lines,
    "fact_1_2_by_method": [{"method": k, "uah": f2(v), "share_pct": round(float(v / fact_tot * 100), 1)}
                           for k, v in sorted(fact_cov.items(), key=lambda x: -x[1])],
    "fact_1_2_total_uah": f2(fact_tot),
    "cards_universe": len(U),
    "cards_by_method": dict(collections.Counter(m["method"] for m in MAP.values())),
}
cards_outside = []
for c, m in MAP.items():
    if m["kind"] in ("поза СС",):
        cards_outside.append({"code": c, "card": CARD[c]["name"], "on_card": CARD[c]["on"], "note": m.get("note", ""),
                              "fact_1": f3(cons[c]["30-1"][0]), "fact_2": f3(cons[c]["30-2"][0]), "fact_3": f3(cons[c]["30-3"][0]),
                              "fact_1_2_uah": f2(cons[c]["30-1"][2] + cons[c]["30-2"][2]), "stock_P": f3(st[c]["P"][0]),
                              "stock_P_uah": f2(st[c]["P"][2])})
cards_outside.sort(key=lambda x: -x["fact_1_2_uah"])

# ─────────────────────────────── 2. по домам: понад / економія / треба, ложное
BY_CARD_OVER_MAP_EXTRA = {}
for h in HOUSES:
    for kk in ("poza", "mshp", "unmapped"):
        for c, v in MAPPED_EX[h][kk].items():
            BY_CARD_OVER[("map", h, c)] = BY_CARD_OVER.get(("map", h, c), Z) + v[1]


def fam(c):
    return "septic" if is_kit_family(c) else c


def fam_name(fk):
    return "Комплект септика (складові + напівфабрикат)" if fk == "septic" else CARD[fk]["name"]


def over_diff(h):
    a, b = collections.defaultdict(Decimal), collections.defaultdict(Decimal)
    for (v, hh, c), g in BY_CARD_OVER.items():
        if hh != h:
            continue
        (a if v == "now" else b)[fam(c)] += g
    gone = {k: a[k] - b.get(k, Z) for k in a if a[k] - b.get(k, Z) > Decimal("0.005")}
    appear = {k: b[k] - a.get(k, Z) for k in b if b[k] - a.get(k, Z) > Decimal("0.005")}
    return gone, appear


def econ_diff(h):
    """Економія по ОН строки: «сейчас» по ОН, «с соответствием» — сумма экономий строк этого ОН."""
    a = {k: v["econ_g"] for k, v in NOW[h].items() if v["econ_g"]}
    b = collections.defaultdict(Decimal)
    for n, v in MAPPED[h].items():
        if not v["econ_g"]:
            continue
        k = "гвинтові кріплення" if n == GROUP_KEY else LINE[n]["k"]
        b[k] += v["econ_g"]
    gone = {k: a[k] - b.get(k, Z) for k in a if a[k] - b.get(k, Z) > Decimal("0.005")}
    appear = {k: b[k] - a.get(k, Z) for k in b if b[k] - a.get(k, Z) > Decimal("0.005")}
    return gone, appear


def mshp_sum(h, view):
    if view == "now":
        return NOW[h].get("мшп", {}).get("over_g", Z)
    return sum(v[1] for v in MAPPED_EX[h]["mshp"].values())


by_house = []
for h in HOUSES:
    now_over = sum(v["over_g"] for v in NOW[h].values())
    map_over_lines = sum(v["over_g"] for v in MAPPED[h].values())
    poza = sum(v[1] for v in MAPPED_EX[h]["poza"].values())
    mshp = mshp_sum(h, "map")
    now_ec = sum(v["econ_g"] for v in NOW[h].values()); map_ec = sum(v["econ_g"] for v in MAPPED[h].values())
    og, oa = over_diff(h); eg, ea = econ_diff(h)
    rec = {"n": HN[h], "now": {"ponad_uah": f2(now_over), "of_it_mshp_uah": f2(mshp_sum(h, "now")), "ekonomiya_uah": f2(now_ec)},
           "mapped": {"ponad_uah": f2(map_over_lines + poza + mshp), "ponad_po_strokah_ss_uah": f2(map_over_lines),
                      "poza_ss_uah": f2(poza), "mshp_uah": f2(mshp), "ekonomiya_uah": f2(map_ec)},
           "false_ponad_gone_uah": f2(sum(og.values())), "hidden_ponad_revealed_uah": f2(sum(oa.values())),
           "false_ekonomiya_gone_uah": f2(sum(eg.values())), "hidden_ekonomiya_revealed_uah": f2(sum(ea.values())),
           "top_false_ponad": [{"card": fam_name(k), "uah": f2(v)} for k, v in sorted(og.items(), key=lambda x: -x[1])[:8]],
           "top_revealed_ponad": [{"card": fam_name(k), "uah": f2(v)} for k, v in sorted(oa.items(), key=lambda x: -x[1])[:8]],
           "top_false_ekonomiya": [{"on": ON_NAME.get(k, k), "uah": f2(v)} for k, v in sorted(eg.items(), key=lambda x: -x[1])[:8]],
           "top_revealed_ekonomiya": [{"on": ON_NAME.get(k, k), "uah": f2(v)} for k, v in sorted(ea.items(), key=lambda x: -x[1])[:6]]}
    if h in DONE or h == "30-3":
        rk_over = sum(Dd(r["ПонадСум"]) for r in RAW["rk_ost"] if (house(r["Подр"] + " IR") or house("IR " + r["Подр"])) == h)
        rec["now"]["rk_doc_ponad_uah"] = f2(rk_over)
    if h == "30-3":
        rec["note"] = "№3 частковий: «економія» не рахується (недосписане = потреба), понад — лише те, що вже списано/лежить понад норму"
    if h not in DONE and h != "30-3":
        rec["note"] = "фактів немає — різниця видів лише в «треба / перемістити / докупити»"
    by_house.append(rec)

# ─────────────────────────────── 3. докупити №3–№7 (Шахматка vs связь), FIFO по домам
# «не подтверждено фактом №1–№2»: строка СС, по которой на №1 и №2 не списано и не лежит ни одной единицы (вид «с соответствием»);
# для вида «сейчас» — ОН, у которого ВСЕ строки СС такие (ОН со смешанными строками считается подтверждённой)
zero12_line = {n: all(MAPPED[h][n]["F"] == 0 and MAPPED[h][n]["Sh"] == 0 for h in DONE) for n in MAPPED["30-1"]}
zero12_on = {k: bool(ON_LINES.get(k)) and all(zero12_line.get(n, False) for n in ON_LINES[k]) for k in ON_KEYS}


def unit_of_on(k):
    us = {LINE[n]["unit"] for n in ON_LINES.get(k, [])}
    return "/".join(sorted(us))


def now_items():
    out = []
    for k, v in NOW_BUY.items():
        lp = v["price_last"] if v["price_last"] else v["price_ss"]
        it = {"key": k, "name": ON_NAME.get(k, k), "unit": unit_of_on(k), "sumD": v["sumD"], "Sp": v["Sp"], "move": v["move"],
              "buy": v["buy"], "uah_ss": v["buy"] * v["price_ss"], "uah_last": v["buy"] * lp, "q12": zero12_on.get(k, False),
              "byh": {h: {"buy": x["buy"], "uah": x["buy"] * v["price_ss"]} for h, x in v["byh"].items()}}
        out.append(it)
    return out


def map_items():
    out = []
    for k, v in MAP_BUY.items():
        if k == SEPTIC_LINE:
            it = {"key": k, "name": "Комплект септика", "unit": "компл.", "sumD": v["sumD"], "Sp": v["Sp"], "move": v["move"], "buy": v["buy"],
                  "uah_ss": sum(x["buy_g_ss"] for x in v["byh"].values()), "uah_last": sum(x["buy_g_last"] for x in v["byh"].values()),
                  "q12": False, "parts": {gp: sum(x["buy_parts"][gp] for x in v["byh"].values()) for gp in KIT_COMP},
                  "byh": {h: {"buy": x["buy"], "uah": x["buy_g_ss"]} for h, x in v["byh"].items()}}
        else:
            lp = v["price_last"] if v["price_last"] else v["price_ss"]
            name = LINE[k]["text"] if k in LINE else "Гвинтові кріплення (група)"
            unit = LINE[k]["unit"] if k in LINE else "шт"
            it = {"key": k, "name": name, "unit": unit, "sumD": v["sumD"], "Sp": v["Sp"], "move": v["move"], "buy": v["buy"],
                  "uah_ss": v["buy"] * v["price_ss"], "uah_last": v["buy"] * lp, "q12": zero12_line.get(k, False),
                  "byh": {h: {"buy": x["buy"], "uah": x["buy"] * v["price_ss"]} for h, x in v["byh"].items()}}
        out.append(it)
    return out


NI, MI = now_items(), map_items()


def summ(items, only=None):
    sel = [i for i in items if (only is None or i["q12"] == only)]
    sht = sum(i["buy"] for i in sel if i["unit"] in ("шт", "компл."))
    return {"positions": sum(1 for i in sel if i["buy"] > Decimal("0.0005")), "qty_sht": f3(sht),
            "uah_ss": f2(sum(i["uah_ss"] for i in sel)), "uah_last": f2(sum(i["uah_last"] for i in sel)),
            "move_positions": sum(1 for i in sel if i["move"] > Decimal("0.0005")),
            "by_house_uah_ss": {HN[h]: f2(sum(i["byh"][h]["uah"] for i in sel)) for h in TODO}}


now_all, map_all = summ(NI), summ(MI)
now_conf, map_conf = summ(NI, False), summ(MI, False)
now_q, map_q = summ(NI, True), summ(MI, True)

# ложное понад / экономия — итоги
tot_now_over = sum(Dd(b["now"]["ponad_uah"]) for b in by_house)
tot_map_over = sum(Dd(b["mapped"]["ponad_uah"]) for b in by_house)
totals = {
    "basis": "грн з ПДВ: норма та ціна СС — Комплектующие.Сумма картки 000000009 (з ПДВ); факт і склад — нетто BuhBud × (1 + ПДВ партії), як у бюджеті v3; ціна останньої закупівлі — ПТУ BuhBud (нетто + ПДВ) / кількість",
    "now": {
        "ponad": f2(tot_now_over), "ponad_by_house": {b["n"]: b["now"]["ponad_uah"] for b in by_house[:3]},
        "ekonomiya": f2(sum(Dd(b["now"]["ekonomiya_uah"]) for b in by_house)),
        "dokupyty_3_7_qty": {"positions": now_all["positions"], "sht_only": now_all["qty_sht"]},
        "dokupyty_3_7_uah": now_all["uah_ss"], "dokupyty_3_7_uah_last_price": now_all["uah_last"],
        "dokupyty_3_7_uah_of_it_on_not_issued_on_1_2": now_q["uah_ss"],
        "note_not_issued": "загальні назви, по всіх строках яких на №1–№2 не видано жодної одиниці (бойлер 82, стільці, банер, підвіконик, труба гофра, сигналізація…)",
        "peremistyty_positions": now_all["move_positions"],
    },
    "mapped": {
        "ponad": f2(tot_map_over), "ponad_by_house": {b["n"]: b["mapped"]["ponad_uah"] for b in by_house[:3]},
        "ponad_split": {"po_strokah_ss": f2(sum(Dd(b["mapped"]["ponad_po_strokah_ss_uah"]) for b in by_house)),
                        "poza_ss": f2(sum(Dd(b["mapped"]["poza_ss_uah"]) for b in by_house)),
                        "mshp": f2(sum(Dd(b["mapped"]["mshp_uah"]) for b in by_house))},
        "ekonomiya": f2(sum(Dd(b["mapped"]["ekonomiya_uah"]) for b in by_house)),
        "dokupyty_3_7_qty": {"positions": map_all["positions"], "sht_only": map_all["qty_sht"]},
        "dokupyty_3_7_uah": map_all["uah_ss"], "dokupyty_3_7_uah_last_price": map_all["uah_last"],
        "dokupyty_3_7_uah_of_it_lines_not_issued_on_1_2": map_q["uah_ss"],
        "note_not_issued": "строки СС, по яких на №1–№2 не видано жодної одиниці — вирішити вручну: зв'язати з карткою (напр. ПК 10 ↔ S-8) або виключити зі СС",
        "lines_not_issued_on_1_2": [f"{i['name']}: {fq(i['buy'])} {i['unit']} = {fmt(i['uah_ss'])}" for i in sorted(MI, key=lambda i: -i["uah_ss"]) if i["q12"] and i["uah_ss"] > 0],
        "peremistyty_positions": map_all["move_positions"],
    },
    "false_ponad_gone_1_3": f2(sum(Dd(b["false_ponad_gone_uah"]) for b in by_house)),
    "hidden_ponad_revealed_1_3": f2(sum(Dd(b["hidden_ponad_revealed_uah"]) for b in by_house)),
    "false_ekonomiya_gone_1_2": f2(sum(Dd(b["false_ekonomiya_gone_uah"]) for b in by_house)),
    "hidden_ekonomiya_revealed_1_2": f2(sum(Dd(b["hidden_ekonomiya_revealed_uah"]) for b in by_house)),
}

# по домам — докупить
def top_items(items, h, k=6):
    lst = [i for i in items if i["byh"][h]["uah"] > Decimal("0.005")]
    lst.sort(key=lambda i: -i["byh"][h]["uah"])
    return [{"name": i["name"], "qty": f3(i["byh"][h]["buy"]), "unit": i["unit"], "uah_ss": f2(i["byh"][h]["uah"]),
             "no_fact_on_1_2": i["q12"]} for i in lst[:k]]


dokupyty_by_house = []
for h in TODO:
    dokupyty_by_house.append({
        "n": HN[h], "now_uah": now_all["by_house_uah_ss"][HN[h]], "mapped_uah": map_all["by_house_uah_ss"][HN[h]],
        "mapped_uah_of_it_lines_not_issued_on_1_2": map_q["by_house_uah_ss"][HN[h]],
        "v3_G_h": IRC["deficit"]["folded"]["by_house"][h],
        "top": top_items(MI, h), "top_now": top_items(NI, h, 4)})

# разница по ОН: «сейчас» (Шахматка) -> «с соответствием»
diff_on = collections.defaultdict(lambda: [Z, Z])
for i in NI:
    diff_on[i["key"]][0] += i["uah_ss"]
for i in MI:
    k = "комплект септика" if i["key"] == SEPTIC_LINE else ("гвинтові кріплення" if i["key"] == GROUP_KEY else LINE[i["key"]]["k"])
    diff_on[k][1] += i["uah_ss"]
V3_ADJ = {"гайка": Dd("10948.75"), "шпилька": Dd("72000"), "комплект септика": Dd("48000"), "кріплення ринви": Z, "з'єднувач компресійний": Dd("129.8")}
moves = []
for k, (a, b) in diff_on.items():
    v3 = V3_ADJ.get(k, a)
    if abs(b - a) > Decimal("0.5") or abs(b - v3) > Decimal("0.5"):
        lines = [i for i in MI if (("комплект септика" if i["key"] == SEPTIC_LINE else ("гвинтові кріплення" if i["key"] == GROUP_KEY else LINE[i["key"]]["k"])) == k) and i["uah_ss"] > 0]
        moves.append({"on": ON_NAME.get(k, k), "now_uah": f2(a), "v3_uah": f2(v3), "mapped_uah": f2(b), "delta_mapped_vs_now": f2(b - a),
                      "delta_mapped_vs_v3": f2(b - v3), "no_fact_on_1_2": any(i["q12"] for i in lines) and all(i["q12"] for i in lines),
                      "mapped_lines": [f"{i['name']}: {fq(i['buy'])} {i['unit']} = {fmt(i['uah_ss'])}" + (" (на №1–№2 не списано)" if i["q12"] else "") for i in lines]})
moves.sort(key=lambda x: -abs(x["delta_mapped_vs_now"]))

v3 = IRC["deficit"]["folded"]
vs_budget_v3 = {
    "v3_G_h_total": v3["total"], "v3_by_house": {HN[h]: v for h, v in v3["by_house"].items()},
    "now_shahmatka_total": now_all["uah_ss"], "now_by_house": now_all["by_house_uah_ss"],
    "now_equals_research_cache_f41": abs(Dd(now_all["uah_ss"]) - Dd(IRC["deficit"]["cache"]["total"])) < Decimal("0.01"),
    "v3_adjustments_already_in_budget": IRC["deficit"]["adjust"],
    "mapped_total": map_all["uah_ss"], "mapped_by_house": map_all["by_house_uah_ss"],
    "mapped_of_it_lines_not_issued_on_1_2": map_q["uah_ss"],
    "delta_mapped_vs_v3": f2(Dd(map_all["uah_ss"]) - Dd(str(v3["total"]))),
    "delta_mapped_vs_v3_by_house": {HN[h]: f2(Dd(map_all["by_house_uah_ss"][HN[h]]) - Dd(str(v3["by_house"][h]))) for h in TODO},
    "delta_v3_vs_now": f2(Dd(str(v3["total"])) - Dd(now_all["uah_ss"])),
    "explain": [
        "v3 = Шахматка (1 293 881,78) + 5 ручних поправок (+19 670,20): гайка −19 110,60; шпилька М48 +3 600,00; септик +48 000,00 (4 цілих комплекти); кріплення ринви ← держак −12 300,00; з'єднувач ← муфта −519,20.",
        "Зв'язка відтворює поправки v3 по гайці, шпильці, кріпленню ринви; по септику дає 35 443,83 замість 48 000,00 (запасні деталі складу йдуть у справу); по з'єднувачу +2 725,80 замість −519,20 (американки не зараховуються як компресійні з'єднувачі).",
        "Понад v3 зв'язка системно знімає хибну потребу по кутиках (−20 194,20), вітробар'єру (−11 890,00), кріпленнях труб (−8 240,00), відводах (−4 560,74), болтах, вимикачах — і відкриває приховану потребу по профільній трубі (+55 748,74), паро-гідро бар'єру (+4 875,00) та строках, яких на №1–№2 не видавали (ПК 10 стіни 140 500,00 та ін.).",
    ],
    "by_on": moves[:30],
    "basis": "усі три — за ціною СС з ПДВ, FIFO складу проєкту «ІБ _МД IRC 2026» по домах №3→№7 (як G_h бюджету v3)",
}
