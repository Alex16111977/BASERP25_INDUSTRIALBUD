# -*- coding: utf-8 -*-
"""Z3, «Взгляд бухгалтера»: трудозатраты эталона, вариантов и закупочной 15 м на реальных данных.

Только чтение файлов: снимок BuhBud raw15.json (z3_collect.py, COM ВЫБРАТЬ) + data/demo_specifikaciya_15m.json
+ data/demo_svyaz_ss_30m.json (эталон-прототип 30 м) + data/maprows.json (соответствия 15 м из разбора причин).
В 1С ничего не пишет. Вывод — JSON в Z3_BUH_OUT (по умолчанию рядом, z3_buh_out.json) и сводка в stdout.
"""
import os, sys, re, json, collections, statistics

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.normpath(os.path.join(HERE, "..", "..", "..", "data"))
RAW = json.load(open(os.environ.get("Z3_RAW", os.path.join(HERE, "raw15.json")), encoding="utf-8"))
D15 = json.load(open(os.path.join(DATA, "demo_specifikaciya_15m.json"), encoding="utf-8"))
D30 = json.load(open(os.path.join(DATA, "demo_svyaz_ss_30m.json"), encoding="utf-8"))
MAPR = json.load(open(os.path.join(DATA, "maprows.json"), encoding="utf-8"))
OUT = os.environ.get("Z3_BUH_OUT", os.path.join(HERE, "z3_buh_out.json"))

K = lambda s: " ".join(str(s or "").replace("\xa0", " ").split()).casefold()
SOFT = lambda s: re.sub(r"[^0-9a-zа-яіїєґ]", "", K(s).replace("x", "х").replace("*", "х"))
HOUSES = [f"15 м №{i}" for i in range(1, 7)]
R = {}

# ───────────── 1. Строки СС 008 и ОН: сколько позиций СС у одной ОН (корень ошибок жёлтых)
SS = [l for l in RAW["ss"] if float(l["q"] or 0) != 0]
POS_PER_ON = collections.Counter(l["on_code"] for l in SS)
ONC = collections.defaultdict(list)
for c in RAW["on_cards"]:
    if not c["Пом"]:
        ONC[c["ОНКод"]].append(c["Код"])
LINE = {l["n"]: l for l in RAW["ss"]}

# факт по карточкам на дома 15 м (из закупочной: by_house = списано + лежит)
FACT_Q = {r["code"]: r["by_house"] for r in D15["zakupochnaya"]["rows"]}
USED = {c for c, bh in FACT_Q.items() if any(v for v in bh.values())}

ROWS = D15["etalon"]["rows"]
by_pos = collections.defaultdict(list)
for r in ROWS:
    by_pos[r["pos_n"]].append(r)


def cat_of(n):
    return by_pos[n][0]["auto_category"]


# ───────────── 2. «Вибрати»: сколько кандидатов в подборе и насколько выбор очевиден
size_tok = lambda s: set(re.findall(r"\d+(?:[.,]\d+)?", K(s).replace("*", " ").replace("х", " ").replace("x", " ")))
choose = []
for n in sorted(by_pos):
    cat = cat_of(n)
    if not cat.startswith("вибрати") and not by_pos[n][0]["source"].startswith("Вручну: жовту"):
        continue
    l = LINE[n]
    cards = ONC.get(l["on_code"], [])
    used = [c for c in cards if c in USED]
    chosen = [r["card_code"] for r in by_pos[n] if r.get("card_code")]
    # «топ по факту»: карточка ОН с наибольшим количеством, списанным на 6 домов
    top = max(used, key=lambda c: sum(FACT_Q[c].values()), default=None)
    pt = size_tok(l["pos"])
    by_size = [c for c in cards if pt and pt <= size_tok(next((x["Имя"] for x in RAW["on_cards"] if x["Код"] == c), ""))]
    if not chosen:
        how = "без картки"
    elif chosen[0] == top and len(used) >= 1:
        how = "легко: перша в підборі за фактом"
    elif len(by_size) == 1 and by_size[0] in chosen:
        how = "легко: єдина за розміром з назви"
    elif any(c in ONC.get(l["on_code"], []) for c in chosen):
        how = "середньо: та ж ОН, але не перша за фактом"
    else:
        how = "важко: картка з іншої ОН (призначення, слово прораба/цеху)"
    choose.append({"n": n, "pos": l["pos"], "cat": cat, "cards_in_on": len(cards), "used_15m": len(used),
                   "chosen": chosen, "top_fact": top, "how": how})
R["choose"] = choose
R["choose_how"] = collections.Counter(x["how"] for x in choose)
R["choose_cards_in_on"] = {"median": statistics.median([x["cards_in_on"] for x in choose]),
                           "max": max(x["cards_in_on"] for x in choose),
                           "ge10": sum(1 for x in choose if x["cards_in_on"] >= 10)}

# ───────────── 3. Жёлтые: правило «ОН стоит на 1 позиции СС» ловит ли 9 ошибок
yellow = [n for n in by_pos if cat_of(n).startswith("жовта")]
rej = set(D15["etalon"]["stats"]["yellow_rejected_by_accountant"])
multi = [n for n in yellow if POS_PER_ON[LINE[n]["on_code"]] > 1]
R["yellow"] = {"all": len(yellow), "rejected": sorted(rej), "on_multi_pos": len(multi),
               "rejected_caught_by_rule": sorted(rej & set(multi)),
               "accepted_but_downgraded_by_rule": len(set(multi) - rej),
               "yellow_left_after_rule": len(yellow) - len(multi)}

# ───────────── 4. «Скопіювати картки з еталону 30 м»: тот же этап+текст позиции → та же карточка?
m30 = {}
for m in D30["mapping"]:
    m30[(K(m["etap"]), SOFT(m["text"]))] = [x["code"] for x in m["links"] if x.get("role") != "Виключена"]
cp = collections.Counter()
cp_rows = []
for n in sorted(by_pos):
    l = LINE[n]
    cat = cat_of(n)
    if cat in ("Група", "Вузол (позиція-комплект)"):
        continue
    key = (K(l["etap"]), SOFT(l["pos"]))
    chosen = [r["card_code"] for r in by_pos[n] if r.get("card_code")]
    if key not in m30:
        res = "немає в 30 м"
    elif not chosen:
        res = "у 15 м без картки"
    elif not m30[key]:
        res = "у 30 м без картки"
    elif set(chosen) <= set(m30[key]):
        res = "та сама картка"
    else:
        res = "інша картка"
    manual = cat.startswith("вибрати") or by_pos[n][0]["source"] == "Вручну: жовту відхилено"   # 84 + 9 = 93 ручних вибори
    cp[(res, "ручна" if manual else "авто")] += 1
    if res == "інша картка" and manual:
        cp_rows.append((n, l["pos"], chosen, m30[key]))
R["copy_from_30m"] = {f"{a} / {b}": v for (a, b), v in sorted(cp.items())}
R["copy_from_30m_diff_manual"] = cp_rows

# ───────────── 5. Независимая разметка 15 м из разбора причин (maprows): согласие с выбором прораба
cardname = {r["card_code"]: r["card"] for r in ROWS if r.get("card_code")}
agree = collections.Counter()
for row in MAPR:
    n, on, pos, unit, q, card, conf, note = row
    if n not in by_pos:
        continue
    ch = {K(r["card"]) for r in by_pos[n] if r.get("card")}
    agree[("так" if K(card) in ch else ("ні" if card else "без картки"), conf)] += 1
R["maprows_agree"] = {f"{a} / {b}": v for (a, b), v in sorted(agree.items())}
# по позициям: ни одна карточка разметки причин не совпала с выбором прораба
mp = collections.defaultdict(set)
for row in MAPR:
    if row[5]:
        mp[row[0]].add(K(row[5]))
pos_dis = [n for n in by_pos if any(r.get("card") for r in by_pos[n]) and mp.get(n)
           and not ({K(r["card"]) for r in by_pos[n] if r.get("card")} & mp[n])]
R["maprows_pos_disagree"] = {"positions": len(pos_dis), "list": sorted(pos_dis)}

# деньги по категориям трудности выбора и по отклонённым жёлтым (норма з ПДВ)
sumpos = lambda ns: round(sum(r["sum_norm_vat"] for n in ns for r in by_pos[n]), 2)
R["choose_how_money_vat"] = {h: sumpos([x["n"] for x in choose if x["how"] == h]) for h in R["choose_how"]}
R["yellow"]["rejected_norm_vat"] = sumpos(rej)
R["yellow"]["accepted_norm_vat"] = sumpos(set(yellow) - rej)

# ───────────── 6. Единицы: где нужен k; что видит РК сейчас (розбіжність одиниць)
R["k_rows"] = [(x["pos_n"], x["pos"], x["unit_line"], x["card_unit"], x["k"], x["k_source"]) for x in D15["etalon"]["stats"]["lines_need_k"]]

# ───────────── 7. Узлы: сколько строк перенести, спорных решений
nodes = D15["etalon"]["nodes"]
R["nodes"] = {"total": len(nodes), "lines_moved": sum(x["lines"] for x in nodes), "disputed": len(D15["etalon"]["disputed"]),
              "area": D15["etalon"]["stats"]["area_nodes"]}

# ───────────── 8. Варианты: правки и копии узлов по домам
vb = D15["variants"]["by_house"]
R["variants"] = {h: {"edits": len(v["edits"]), "by_rule": dict(collections.Counter(e["rule"].split()[0] for e in v["edits"])),
                     "node_copies": len(v["copied_nodes"]), "delta_vat": v["delta_vs_etalon_vat"]} for h, v in vb.items()}
R["variants_total"] = {"edits": sum(x["edits"] for x in R["variants"].values()),
                       "node_copies": sum(x["node_copies"] for x in R["variants"].values()),
                       "v2_rejected": len(D15["variants"]["summary"]["v2_rejected_by_builder"])}

# ───────────── 9. Сейчас (РК) против спецификации: сколько строк объяснять на дом
rk_by_house = {}
sv = collections.defaultdict(list)
for r in RAW["rk_sverh"]:
    sv[K(r["Дом"]).replace("мд irc 2026 ", "")].append(r)
pr = collections.Counter(K(r["Дом"]).replace("мд irc 2026 ", "") for r in RAW["rk_prichiny"])
for h in HOUSES:
    lst = sv[K(h)]
    nf = D15["norma_fakt_by_house"][h]["cards"]
    rk_by_house[h] = {
        "rk_rows_sverh": len(lst),
        "rk_rows_no_reason": sum(1 for r in lst if not (r.get("Причина") or "").strip()),
        "rk_prichiny_ceh": pr[K(h)],
        "rk_rows_kolss0": sum(1 for r in lst if not float(r.get("КолСС") or 0)),
        "spec_over_cards": sum(1 for c in nf if c.get("over_s", 0) > 0.004),
        "spec_poza_cards": sum(1 for c in nf if c.get("poza_s", 0) > 0.004),
        "spec_mshp_cards": sum(1 for c in nf if c.get("mshp_s", 0) > 0.004),
        "spec_econ_cards": sum(1 for c in nf if c.get("econ_s", 0) > 0.004),
        "spec_over_ge500": sum(1 for c in nf if c.get("over_s", 0) + c.get("poza_s", 0) >= 500),
    }
R["now_vs_spec_rows"] = rk_by_house
# «поза» на доме, которые повторяются на 5–6 домах = кандидаты в СС (одна правка вместо 6 причин)
poza_h = collections.Counter()
for h in HOUSES:
    for c in D15["norma_fakt_by_house"][h]["cards"]:
        if c.get("poza_s", 0) > 0.004:
            poza_h[c["key"]] += 1
R["poza_repeat"] = {"cards": len(poza_h), "on_5_6_houses": sum(1 for v in poza_h.values() if v >= 5),
                    "on_1_house": sum(1 for v in poza_h.values() if v == 1)}

# ───────────── 10. Закупочная: сколько строк требует решения бухгалтера
zr = D15["zakupochnaya"]["rows"]
pcs = {"шт", "компл", "компл.", "к-т", "пара", "балон", "рул", "пак", "пач", "упак"}
dec = {
    "rows": len(zr),
    "poza": sum(1 for r in zr if r["group"] == "поза еталоном"),
    "mshp": sum(1 for r in zr if r["group"] == "МШП"),
    "only_1_house": sum(1 for r in zr if r["houses_with"] == 1),
    "only_1_house_not_mshp": sum(1 for r in zr if r["houses_with"] == 1 and r["group"] != "МШП"),
    "cv_gt_0_3": sum(1 for r in zr if r["cv"] > 0.3),
    "cv_gt_0_3_not_mshp": sum(1 for r in zr if r["cv"] > 0.3 and r["group"] != "МШП"),
    "etalon_dev_gt_30pct": sum(1 for r in zr if r["group"] == "еталон" and r.get("etalon_norm_q")
                               and abs(r["avg"] - r["etalon_norm_q"]) / r["etalon_norm_q"] > 0.3),
    "piece_fractional_avg": sum(1 for r in zr if K(r["unit"]).strip(".") in {u.strip(".") for u in pcs}
                                and abs(r["avg"] - round(r["avg"])) > 1e-6),
    "price_not_project_ptu": sum(1 for r in zr if r["price_basis"] != "ПТУ проєкту"),
}
# строки к разбору руками: «поза» + (эталон с отклонением >30 %) + (не МШП и только на 1 доме) — без повторов
need = set()
for r in zr:
    if r["group"] == "поза еталоном":
        need.add(r["code"])
    elif r["group"] == "еталон" and r.get("etalon_norm_q") and abs(r["avg"] - r["etalon_norm_q"]) / r["etalon_norm_q"] > 0.3:
        need.add(r["code"])
    elif r["group"] != "МШП" and r["houses_with"] == 1:
        need.add(r["code"])
dec["to_review_manually"] = len(need)
dec["to_review_uah_avg_house"] = round(sum(r["avg_uah_last"] for r in zr if r["code"] in need), 2)
R["zakupochnaya_decisions"] = dec

# ───────────── 11. Модель трудозатрат (минуты на операцию — оценка бухгалтера-практика, а не замер)
NORM = {"заповнити_протокол": 15, "зелена": 0.5, "жовта": 1.5,
        "легко: перша в підборі за фактом": 2, "легко: єдина за розміром з назви": 2,
        "середньо: та ж ОН, але не перша за фактом": 5,
        "важко: картка з іншої ОН (призначення, слово прораба/цеху)": 15, "без картки": 10,
        "k": 3, "група": 5, "вузол": 5, "спірне": 2, "площадний": 10, "септик_з_КН": 10, "ціни_СС_діє": 30,
        "нульовий_з_фактом": 3, "копія_та_сама": 1, "копія_інша_доп": 2,
        "вар_пошук_причини": 60, "вар_пошук_V2_на_дім": 20, "вар_правка": 4, "вар_копія_вузла": 5, "вар_коментар": 1, "вар_діє": 10,
        "вар_пропозиція_кнопкою": 2, "вар_діє_кнопкою": 5,
        "зак_рішення": 2, "зак_масово": 10, "зак_мшп": 5, "зак_діє": 10,
        "рк_рядок_понад": 1, "рк_без_причини": 3, "рк_одиниці": 2, "рк_друк_завантаження": 20}
st_ = D15["etalon"]["stats"]
zero_with_fact = 18
m = collections.OrderedDict()
m["Заповнити з СС + протокол 22 нульових"] = NORM["заповнити_протокол"]
m["зелені 16 × 0,5"] = st_["green"] * NORM["зелена"]
m["жовті 59 × 1,5"] = st_["yellow"] * NORM["жовта"]
m["ручний вибір 93 (легко/середньо/важко/без картки)"] = sum(NORM[x["how"]] for x in choose)
m["коефіцієнти 6 × 3"] = len(R["k_rows"]) * NORM["k"]
m["група кріплення"] = NORM["група"]
m["вузли: 21 × 5 + спірні 32 × 2 + площадні 2 × 10 + септик з КН"] = (R["nodes"]["total"] * NORM["вузол"] + R["nodes"]["disputed"] * NORM["спірне"]
                                                                       + 2 * NORM["площадний"] + NORM["септик_з_КН"])
m["ціни норми, порівняти з СС, Діє, еталон для СС"] = NORM["ціни_СС_діє"]
m["нульові рядки з фактом 18 × 3"] = zero_with_fact * NORM["нульовий_з_фактом"]
etalon_scratch = sum(m.values())
# с копированием карточек и дерева узлов из эталона 30 м
copy_res = {}
for n in sorted(by_pos):
    l = LINE[n]; key = (K(l["etap"]), SOFT(l["pos"]))
    chosen = [r["card_code"] for r in by_pos[n] if r.get("card_code")]
    if key in m30 and chosen and m30[key] and set(chosen) <= set(m30[key]):
        copy_res[n] = "same"
    elif key in m30 and chosen and m30[key]:
        copy_res[n] = "diff"
manual_copy = 0
for x in choose:
    r_ = copy_res.get(x["n"])
    manual_copy += NORM["копія_та_сама"] if r_ == "same" else NORM[x["how"]] + (NORM["копія_інша_доп"] if r_ == "diff" else 0)
m2 = collections.OrderedDict(m)
m2["ручний вибір 93 (легко/середньо/важко/без картки)"] = manual_copy
m2["вузли: 21 × 5 + спірні 32 × 2 + площадні 2 × 10 + септик з КН"] = 60     # дерево копируется, правятся только «Водостік 15 м», площади, септик
etalon_copy = sum(m2.values())
var_full = (NORM["вар_пошук_причини"] + 6 * NORM["вар_пошук_V2_на_дім"] + R["variants_total"]["edits"] * (NORM["вар_правка"] + NORM["вар_коментар"])
            + R["variants_total"]["node_copies"] * NORM["вар_копія_вузла"] + 6 * NORM["вар_діє"])
var_button = (R["variants_total"]["edits"] + R["variants_total"]["v2_rejected"]) * NORM["вар_пропозиція_кнопкою"] + 6 * NORM["вар_діє_кнопкою"]
var_v1 = 5 * (NORM["вар_правка"] + NORM["вар_коментар"]) + 5 * NORM["вар_копія_вузла"] + 3 * NORM["вар_діє"]
zak = dec["to_review_manually"] * NORM["зак_рішення"] + NORM["зак_масово"] + NORM["зак_мшп"] + NORM["зак_діє"]
rk_now = [v["rk_rows_sverh"] * NORM["рк_рядок_понад"] + v["rk_rows_no_reason"] * NORM["рк_без_причини"] + 6.2 * NORM["рк_одиниці"]
          + NORM["рк_друк_завантаження"] for v in rk_by_house.values()]
rk_spec = [(v["spec_over_cards"] + v["spec_poza_cards"]) * NORM["рк_рядок_понад"] + v["spec_over_ge500"] * NORM["рк_без_причини"]
           + NORM["рк_друк_завантаження"] for v in rk_by_house.values()]
# после правки СС: «поза»-карточки, повторяющиеся на 5–6 домах, становятся нормой → объяснять только остальные
rep = {k for k, v in poza_h.items() if v >= 5}
rk_spec_ss = []
for h in HOUSES:
    nf = D15["norma_fakt_by_house"][h]["cards"]
    over = sum(1 for c in nf if c.get("over_s", 0) > 0.004)
    poza_rest = sum(1 for c in nf if c.get("poza_s", 0) > 0.004 and c["key"] not in rep)
    big = sum(1 for c in nf if (c.get("over_s", 0) + (c.get("poza_s", 0) if c["key"] not in rep else 0)) >= 500)
    rk_by_house[h]["after_ss_rows"] = over + poza_rest
    rk_by_house[h]["after_ss_ge500"] = big
    rk_spec_ss.append((over + poza_rest) * NORM["рк_рядок_понад"] + big * NORM["рк_без_причини"] + NORM["рк_друк_завантаження"])
R["time_model"] = {"norms_min": NORM, "etalon_scratch_parts_min": m, "etalon_scratch_h": round(etalon_scratch / 60, 1),
                   "etalon_copy30_parts_min": m2, "etalon_copy30_h": round(etalon_copy / 60, 1),
                   "variants_full_manual_h": round(var_full / 60, 1), "variants_button_h": round(var_button / 60, 1),
                   "variants_only_V1_h": round(var_v1 / 60, 1), "zakupochnaya_h": round(zak / 60, 1),
                   "rk_now_per_house_h": [round(x / 60, 1) for x in rk_now], "rk_now_6_houses_h": round(sum(rk_now) / 60, 1),
                   "rk_spec_per_house_h": [round(x / 60, 1) for x in rk_spec], "rk_spec_6_houses_h": round(sum(rk_spec) / 60, 1),
                   "rk_spec_after_ss_per_house_h": [round(x / 60, 1) for x in rk_spec_ss], "rk_spec_after_ss_6_houses_h": round(sum(rk_spec_ss) / 60, 1),
                   "rows_after_ss": {h: (rk_by_house[h]["after_ss_rows"], rk_by_house[h]["after_ss_ge500"]) for h in HOUSES},
                   "copy30_manual_rows": collections.Counter(copy_res.get(x["n"], "none") for x in choose)}

json.dump(R, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

# ───────────── сводка
print("вибрати + відхилені жовті:", len(choose), dict(R["choose_how"]))
print("карток в ОН (підбір):", R["choose_cards_in_on"])
print("жовті:", R["yellow"])
print("копія з 30 м:", R["copy_from_30m"])
print("  ручні, де 30 м дає іншу картку:", len(cp_rows))
for x in cp_rows[:12]:
    print("   ", x)
print("maprows (розбір причин) vs прораб:", R["maprows_agree"], R["maprows_pos_disagree"])
print("гроші за складністю вибору, з ПДВ:", R["choose_how_money_vat"])
print("k:", R["k_rows"])
print("вузли:", R["nodes"])
print("варіанти:", R["variants_total"])
for h, v in R["variants"].items():
    print("   ", h, v)
print("РК зараз vs специфікація (рядків на дім):")
for h, v in rk_by_house.items():
    print("   ", h, v)
print("поза — повтор по будинках:", R["poza_repeat"])
print("закупівельна — рішення:", dec)
print("ТРУДОЗАТРАТИ:", json.dumps({k: v for k, v in R["time_model"].items() if k != "norms_min"}, ensure_ascii=False, default=str, indent=1))
print("OUT ->", OUT)
