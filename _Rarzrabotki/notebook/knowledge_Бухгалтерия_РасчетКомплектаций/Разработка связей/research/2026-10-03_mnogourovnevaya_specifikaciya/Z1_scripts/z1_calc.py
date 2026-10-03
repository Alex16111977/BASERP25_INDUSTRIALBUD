# -*- coding: utf-8 -*-
"""Z1. Закупочная спецификация типа «МД IRС 2026 30 м» по фактическим списаниям домов №1 и №2
и докупка на №3–№7. ТОЛЬКО ЧТЕНИЕ: база не трогается, читается снимок raw.json прототипа связи (03.10.2026 10:52).

Опирается на код прототипа (base.py / mapping.py / views.py из scripts/03_prototip_svyazi_30m):
  base.cons[c][h]   — списано с карточки со склада дома h (Кт мат -> Дт не мат), кол/нетто/брутто
  base.st[c][cls]   — сальдо карточки по классу склада: P = склад проекта «ІБ _МД IRC 2026», H:30-n = склады дома
  views.ASM         — складові септика, зібрані в напівфабрикат «Комлпект септика» на складі дома
  mapping.MAP       — связь карточки со строкой СС (эталон), вид: Основна / Аналог / Складова / Групова / поза СС / МШП
  views.MAP_BUY     — эталон «со связью»: докупка по строкам СС (для сравнения)
  views.NOW_BUY     — Шахматка (по ОН)
Запуск: python z1_calc.py  (переменная SVYAZ_DIR — папка с raw.json + base/mapping/views; по умолчанию scratchpad сессии).
Выход: data/zakupochnaya_spec_30m.json и z1_out.txt рядом со скриптом."""
import os, sys, json, math, collections
from decimal import Decimal

SVYAZ = os.environ.get("SVYAZ_DIR", r"C:/Users/SUPPOR~1/AppData/Local/Temp/claude/C--Configuration-downloads-BASERP25--claude-worktrees-budget-calc-oct-dec-2026-20c998/74acb5f4-5c54-4380-b2d5-fb369593cb1b/scratchpad/svyaz")
sys.path.insert(0, SVYAZ)
from views import *            # noqa: E402,F401,F403  (base + mapping + views прототипа)
from base import _d            # noqa: E402  (дата ПТУ -> ISO)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))           # «Разработка связей»
DATA = os.path.join(ROOT, "data")
DEMO = json.load(open(os.path.join(DATA, "demo_svyaz_ss_30m.json"), encoding="utf-8"))
OUT_JSON = os.path.join(DATA, "zakupochnaya_spec_30m.json")
OUT_TXT = os.path.join(HERE, "z1_out.txt")

BASE_H = ["30-1", "30-2"]                     # дома-база закупочной спецификации
REST = ["30-3", "30-4", "30-5", "30-6", "30-7"]   # оставшиеся, FIFO склада проекта в этом порядке
NOT_STARTED = ["30-4", "30-5", "30-6", "30-7"]
HN = {f"30-{i}": f"№{i}" for i in range(1, 8)}
PIECE = {"шт", "рул", "пак", "пар", "компл", "компл.", "уп", "упак", "лист", "набір", "балон"}
CV_LIMIT = Decimal("0.3")                     # порог «нестабильно» для коэффициента вариации по двум домам
EPS = Decimal("0.0000001")

LOG = []


def log(*a):
    s = " ".join(str(x) for x in a)
    LOG.append(s)
    print(s)


def f2(x):
    return r2(x)


def f3(x):
    return r3(x)


def unit(c):
    return (CARD[c]["unit"] or "").strip()


def is_piece(c):
    return unit(c) in PIECE


def ceil_q(x):
    """вверх до целого (с допуском на хвосты Decimal)."""
    return Decimal(math.ceil(x - EPS)) if x > 0 else Z


# ───────────────────────────── факт по карточке и дому
def fact_a(c, h):
    """(а) списано на дом: cons + складові, зібрані в напівфабрикат септика на складі дома. Кол, брутто."""
    q, g = cons[c][h][0], cons[c][h][2]
    if c in KIT_CARDS:
        a = ASM[(c, h)]
        q += a[0]; g += a[1]
    return q, g


def stock_h(c, h):
    a = st[c]["H:" + h]
    return a[0], a[2]


def fact_b(c, h):
    qa, ga = fact_a(c, h)
    qs, gs = stock_h(c, h)
    return qa + qs, ga + gs


# ───────────────────────────── цена закупки с проверкой на ошибку единицы
PRICE_K = Decimal(5)          # закупка дороже/дешевле собівартості факту більш ніж у 5 разів — вважаємо помилкою одиниці в ПТУ
# Ціна закупівлі — по ДОКУМЕНТУ ПТУ: середньозважена по всіх рядках картки в ньому (Σ брутто / Σ кількість).
# Виправлення перевірки 03.10: раніше бралась ціна одного рядка, і при кількох рядках картки в останньому ПТУ
# результат залежав від порядку рядків у вибірці (Лист рівний ІБ00-004049: 35 м² по 172,60 і 105 м² по 270,06).
_RCD = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z]))
for _r in RAW["receipts"]:
    if not _r["Проведен"] or not Dd(_r["Кол"]):
        continue
    _net = Dd(_r["Сумма"]) - (Dd(_r["НДС"]) if _r["ВклНДС"] else Z)
    _a = _RCD[_r["Код"]][(_d(_r["Дата"]), _r["Док"])]
    _a[0] += _net + Dd(_r["НДС"]); _a[1] += Dd(_r["Кол"])
RCPT = collections.defaultdict(list)
for _c, _docs in _RCD.items():
    for (_dt, _doc), (_g, _q) in _docs.items():
        if _q:
            RCPT[_c].append({"date": _dt, "price_g": _g / _q, "doc": _doc, "qty": _q})
for _v in RCPT.values():
    _v.sort(key=lambda x: (x["date"], x["doc"]), reverse=True)


def checked_last_price(c, fact_cost):
    """(ціна, позначка, ціна сирої останньої закупівлі): остання закупівля; якщо вона відхиляється від собівартості
    факту > 5 разів — попередня правдоподібна; якщо таких немає — None (далі fallback)."""
    lst = RCPT.get(c, [])
    if not lst:
        return None, None, None
    raw = lst[0]["price_g"]
    if not fact_cost:
        return raw, "остання закупівля", raw
    for i, x in enumerate(lst):
        k = x["price_g"] / fact_cost
        if Decimal(1) / PRICE_K <= k <= PRICE_K:
            return x["price_g"], ("остання закупівля" if i == 0 else "попередня закупівля (остання — помилка одиниці в ПТУ)"), raw
    return None, "усі закупівлі відхиляються від собівартості факту > 5 разів (одиниця в ПТУ)", raw


# ───────────────────────────── классификация и цены СС
def first_link(c):
    m = MAP.get(c)
    if not m or not m["links"]:
        return None
    return sorted(m["links"], key=lambda l: l["prio"])[0]


GROUP_N = sum(LINE[n]["q"] for n in GROUP_LINES)
GROUP_S = sum(LINE[n]["sum"] for n in GROUP_LINES)
SEP = MAP_BUY[SEPTIC_LINE]
KIT_LP = sum(Decimal(str(SEP["lp_comp"][gp])) * per for gp, per in KIT_COMP.items())


def classify(c):
    """группа (эталон / поза еталоном / МШП), ключ сравнения, строка СС, способ связи, норма эталона в ед. карточки, цена СС за ед. карточки."""
    m = MAP.get(c)
    if m is None:
        return {"group": "поза еталоном", "kind": "не в mapping", "key": "C:" + c, "ss_line": None, "method": "не знайдено в mapping",
                "etalon_norm": None, "etalon_shared": 0, "price_ss": None, "note": "картки немає в універсумі прототипу"}
    kind = m["kind"]
    if kind == "МШП":
        return {"group": "МШП", "kind": kind, "key": "C:" + c, "ss_line": None, "method": "МШП", "etalon_norm": Z,
                "etalon_shared": 0, "price_ss": None, "note": m.get("note", "")}
    if kind in ("поза СС", "не визначено"):
        return {"group": "поза еталоном", "kind": kind, "key": "C:" + c, "ss_line": None, "method": m["method"],
                "etalon_norm": Z, "etalon_shared": 0, "price_ss": None, "note": m.get("note", "")}
    if kind == "Групова":
        return {"group": "еталон", "kind": kind, "key": "G", "ss_line": "група «Гвинтові кріплення» (9 строк)",
                "method": "групова норма", "etalon_norm": GROUP_N, "etalon_shared": 8,
                "price_ss": GROUP_S / GROUP_N, "note": m.get("note", "")}
    if kind == "Складова комплекту":
        gp = KIT_CARDS[c]; per = KIT_COMP[gp]
        lpc = Decimal(str(SEP["lp_comp"][gp]))
        shared = sum(1 for x, g in KIT_CARDS.items() if g == gp)
        return {"group": "еталон", "kind": kind, "key": "L:%d" % SEPTIC_LINE,
                "ss_line": "%d %s (%s ×%d)" % (SEPTIC_LINE, LINE[SEPTIC_LINE]["text"], gp, per), "method": "комплект",
                "etalon_norm": Decimal(per) * LINE[SEPTIC_LINE]["q"], "etalon_shared": shared,
                "price_ss": lpc * LINE[SEPTIC_LINE]["price"] / KIT_LP if KIT_LP else None,
                "note": "ціна СС складової = ціна комплекту 12 000 × частка складової в ціні останньої закупівлі комплекту"}
    l = first_link(c)
    n = l["line"]
    shared = len(BY_LINE.get(n, []))
    return {"group": "еталон", "kind": kind, "key": "L:%d" % n, "ss_line": "%d %s" % (n, LINE[n]["text"]),
            "method": m["method"] + ("" if len(m["links"]) == 1 else " (+%d строк)" % (len(m["links"]) - 1)),
            "etalon_norm": LINE[n]["q"] / l["coef"] if l["coef"] else None, "etalon_shared": shared,
            "price_ss": LINE[n]["price"] * l["coef"], "coef": l["coef"], "line_unit": LINE[n]["unit"],
            "note": l.get("note", "")}


# ───────────────────────────── расчёт одного варианта
def compute(variant, override=None, est="avg"):
    """variant 'a' — только списано; 'b' — списано + лежит на складе дома (основной).
    override: {код: {'avg_from': 'a', 'release_house_stock': True}} — сценарий для ЦСП/ОСБ на №2."""
    override = override or {}
    rows = {}
    cards = sorted(c for c in CARD if c != KIT and (any(fact_b(c, h)[0] for h in BASE_H) or any(fact_a(c, h)[0] for h in BASE_H)))
    for c in cards:
        ff = fact_b if variant == "b" else fact_a
        ov = override.get(c, {})
        if ov.get("avg_from") == "a":
            ff = fact_a
        f1q, f1g = ff(c, "30-1"); f2q, f2g = ff(c, "30-2")
        avg = (f1q + f2q) / 2 if est == "avg" else (min(f1q, f2q) if est == "min" else max(f1q, f2q))
        if avg <= 0:
            continue
        # №3: max(0, среднее − (списано №3 + лежит на складе №3))
        f3q = fact_a(c, "30-3")[0]; s3q = stock_h(c, "30-3")[0]
        gap3 = max(Z, avg - (f3q + s3q))
        surplus3 = max(Z, (f3q + s3q) - avg)
        need = {"30-3": gap3}
        for h in NOT_STARTED:
            need[h] = avg
        need_3_7 = sum(need.values())
        sp = max(Z, st[c]["P"][0])
        own = {h: max(Z, stock_h(c, h)[0]) for h in NOT_STARTED}        # склады №4–№7 (сейчас пустые)
        released = Z
        if ov.get("release_house_stock"):
            released = sum(max(Z, stock_h(c, h)[0]) for h in BASE_H)
        pool = sp + released
        byh, rem_pool = {}, pool
        for h in REST:
            d = need[h]
            o = min(d, own.get(h, Z)); d2 = d - o
            if h in own and own[h] > o:
                rem_pool += own[h] - o
            take = min(d2, rem_pool); rem_pool -= take
            byh[h] = {"need": d, "own": o, "from_project": take, "buy_raw": d2 - take}
        buy_raw = sum(v["buy_raw"] for v in byh.values())
        piece = is_piece(c)
        if piece:
            buy = ceil_q(buy_raw)
            cum, prev = Z, Z
            for h in REST:
                cum += byh[h]["buy_raw"]
                ci = ceil_q(cum)
                byh[h]["buy"] = ci - prev; prev = ci
            # построчное округление на каждый дом (иллюстрация)
            need_int = sum(ceil_q(need[h]) for h in REST)
            buy_house_round = max(Z, need_int - sum(own.values()) - pool)
        else:
            buy = buy_raw
            for h in REST:
                byh[h]["buy"] = byh[h]["buy_raw"]
            buy_house_round = buy
        cl = classify(c)
        qf = f1q + f2q
        fact_cost = (f1g + f2g) / qf if qf else None
        lp, lp_flag, lp_raw = checked_last_price(c, fact_cost)
        if lp:
            p_last, b_last = lp, lp_flag
        elif cl["price_ss"]:
            p_last, b_last = cl["price_ss"], "ціна СС (" + (lp_flag or "закупівель немає") + ")"
        else:
            p_last, b_last = fact_cost, "собівартість факту №1–№2 (" + (lp_flag or "немає ні закупівлі, ні ціни СС") + ")"
        p_ss, b_ss = (cl["price_ss"], "ціна СС") if cl["price_ss"] else ((lp, "остання закупівля (немає ціни СС)") if lp else (fact_cost, "собівартість факту №1–№2"))
        # коэффициент вариации по двум домам (по варианту b)
        cv = (Decimal(2).sqrt() * abs(f1q - f2q) / (f1q + f2q)) if (f1q + f2q) > 0 else Z
        rows[c] = {
            "card": CARD[c]["name"], "code": c, "unit": unit(c), "piece": piece, "on_card": CARD[c]["on"],
            "group": cl["group"], "kind": cl["kind"], "key": cl["key"], "ss_line": cl["ss_line"], "method": cl["method"],
            "f1": f1q, "f2": f2q, "f1g": f1g, "f2g": f2g, "avg": avg,
            "fa1": fact_a(c, "30-1")[0], "fa2": fact_a(c, "30-2")[0], "sh1": stock_h(c, "30-1")[0], "sh2": stock_h(c, "30-2")[0],
            "f3": f3q, "s3": s3q, "gap3": gap3, "surplus3": surplus3,
            "etalon_norm": cl["etalon_norm"], "etalon_shared": cl["etalon_shared"], "coef": cl.get("coef"),
            "need_3_7": need_3_7, "stock_project": sp, "stock_4_7": sum(own.values()), "released": released,
            "buy_raw": buy_raw, "buy": buy, "buy_house_round": buy_house_round,
            "price_last": lp, "price_last_raw": lp_raw, "price_ss": cl["price_ss"], "fact_cost": fact_cost,
            "p_last": p_last, "basis_last": b_last, "p_ss": p_ss, "basis_ss": b_ss,
            "uah_last": buy * p_last if p_last else Z, "uah_ss": buy * p_ss if p_ss else Z,
            "uah_last_house_round": buy_house_round * p_last if p_last else Z,
            "cv": cv, "unstable": cv > CV_LIMIT, "note": cl["note"],
            "byh": {h: {"need": v["need"], "own": v["own"], "from_project": v["from_project"], "buy": v["buy"],
                        "uah_last": v["buy"] * p_last if p_last else Z, "uah_ss": v["buy"] * p_ss if p_ss else Z} for h, v in byh.items()},
        }
    return rows


def tot(rows, sel=lambda r: True):
    rs = [r for r in rows.values() if sel(r)]
    return {"cards": len(rs), "cards_to_buy": sum(1 for r in rs if r["buy"] > EPS),
            "uah_last": sum(r["uah_last"] for r in rs), "uah_ss": sum(r["uah_ss"] for r in rs),
            "uah_last_house_round": sum(r["uah_last_house_round"] for r in rs),
            "by_house_last": {h: sum(r["byh"][h]["uah_last"] for r in rs) for h in REST},
            "by_house_ss": {h: sum(r["byh"][h]["uah_ss"] for r in rs) for h in REST},
            "avg_house_uah": sum(r["avg"] * r["p_last"] for r in rs if r["p_last"]),
            "need_uah_last": sum(r["need_3_7"] * r["p_last"] for r in rs if r["p_last"])}


def jt(t):
    return {k: (f2(v) if isinstance(v, Decimal) else ({HN[h]: f2(x) for h, x in v.items()} if isinstance(v, dict) else v)) for k, v in t.items()}


RB = compute("b")
RA = compute("a")
GROUPS = ["еталон", "поза еталоном", "МШП"]

# ───────────────────────────── эталон «со связью» и Шахматка по ключам сравнения
def etalon_items():
    """ключ L:n / G -> {name, unit, buy(ед. строки), uah_ss, uah_last, byh_ss{h}, zero12}"""
    zero12 = {n: all(MAPPED[h][n]["F"] == 0 and MAPPED[h][n]["Sh"] == 0 for h in DONE) for n in MAPPED["30-1"]}
    out = {}
    for k, v in MAP_BUY.items():
        if k == SEPTIC_LINE:
            out["L:%d" % k] = {"name": "Комплект септика", "unit": "компл.", "norm": LINE[k]["q"], "buy": v["buy"],
                               "uah_ss": sum(x["buy_g_ss"] for x in v["byh"].values()),
                               "uah_last": sum(x["buy_g_last"] for x in v["byh"].values()),
                               "byh_ss": {h: x["buy_g_ss"] for h, x in v["byh"].items()},
                               "byh_last": {h: x["buy_g_last"] for h, x in v["byh"].items()}, "zero12": False, "Sp": v["Sp"]}
            continue
        lp = v["price_last"] if v["price_last"] else v["price_ss"]
        key = "G" if k == GROUP_KEY else "L:%d" % k
        out[key] = {"name": LINE[k]["text"] if k in LINE else "Гвинтові кріплення (група)", "unit": LINE[k]["unit"] if k in LINE else "шт",
                    "norm": LINE[k]["q"] if k in LINE else GROUP_N, "buy": v["buy"], "uah_ss": v["buy"] * v["price_ss"],
                    "uah_last": v["buy"] * lp, "byh_ss": {h: x["buy"] * v["price_ss"] for h, x in v["byh"].items()},
                    "byh_last": {h: x["buy"] * lp for h, x in v["byh"].items()}, "zero12": zero12.get(k, False), "Sp": v["Sp"]}
    return out


ET = etalon_items()
ET_TOT_SS = sum(v["uah_ss"] for v in ET.values()); ET_TOT_LAST = sum(v["uah_last"] for v in ET.values())
log("контроль: эталон со связью", f2(ET_TOT_SS), "/", f2(ET_TOT_LAST), " (канон 1 464 265,90 / 1 372 807,82)")
now_ss = sum(v["buy"] * v["price_ss"] for v in NOW_BUY.values())
now_last = sum(v["buy"] * (v["price_last"] if v["price_last"] else v["price_ss"]) for v in NOW_BUY.values())
log("контроль: Шахматка", f2(now_ss), "/", f2(now_last), " (канон 1 293 881,78 / 1 242 856,02)")
V3 = DEMO["vs_budget_v3"]

# ───────────────────────────── сравнение по ключам
def by_key(rows):
    d = collections.defaultdict(lambda: {"cards": [], "uah_last": Z, "uah_ss": Z, "buy_line": Z, "byh_last": collections.defaultdict(lambda: Z)})
    for c, r in rows.items():
        x = d[r["key"]]
        x["cards"].append(c); x["uah_last"] += r["uah_last"]; x["uah_ss"] += r["uah_ss"]
        if r["key"].startswith("L:") and r.get("coef") is not None and r["key"] != "L:%d" % SEPTIC_LINE:
            x["buy_line"] += r["buy"] * r["coef"]
        for h in REST:
            x["byh_last"][h] += r["byh"][h]["uah_last"]
    return d


KB = by_key(RB)
CMP = []
for key in sorted(set(KB) | set(ET)):
    z = KB.get(key); e = ET.get(key)
    zl = z["uah_last"] if z else Z; zs = z["uah_ss"] if z else Z
    el = e["uah_last"] if e else Z; es = e["uah_ss"] if e else Z
    if key.startswith("C:"):
        c = key[2:]; r = RB[c]
        name, grp = r["card"], r["group"]
    else:
        name = e["name"] if e else key; grp = "еталон"
    CMP.append({"key": key, "name": name, "group": grp, "zak_last": zl, "zak_ss": zs, "et_last": el, "et_ss": es,
                "d_last": zl - el, "d_ss": zs - es, "cards": z["cards"] if z else [], "zero12": bool(e and e["zero12"])})

# ───────────────────────────── итоги
TB = {g: tot(RB, lambda r, g=g: r["group"] == g) for g in GROUPS}
TB["всього"] = tot(RB)
TA = {g: tot(RA, lambda r, g=g: r["group"] == g) for g in GROUPS}
TA["всього"] = tot(RA)

log("\n=== ВАРИАНТ (б) списано + лежит — ОСНОВНОЙ")
for g, t in TB.items():
    log(f"{g:15s} карточек {t['cards']:4d} к докупке {t['cards_to_buy']:4d}  по посл. закупке {f2(t['uah_last']):>14,.2f}  по цене СС {f2(t['uah_ss']):>14,.2f}  "
        f"построчно по домам {f2(t['uah_last_house_round']):>14,.2f}  среднее на дом {f2(t['avg_house_uah']):>14,.2f}")
log("по домам (посл.):", {HN[h]: f2(v) for h, v in TB["всього"]["by_house_last"].items()})
log("по домам (СС):  ", {HN[h]: f2(v) for h, v in TB["всього"]["by_house_ss"].items()})
log("\n=== ВАРИАНТ (а) только списано")
for g, t in TA.items():
    log(f"{g:15s} карточек {t['cards']:4d} к докупке {t['cards_to_buy']:4d}  по посл. закупке {f2(t['uah_last']):>14,.2f}  по цене СС {f2(t['uah_ss']):>14,.2f}")

# эталон по группам: всё эталон; плюс доля строк без выдачи на №1–№2
et_zero12_ss = sum(v["uah_ss"] for v in ET.values() if v["zero12"]); et_zero12_last = sum(v["uah_last"] for v in ET.values() if v["zero12"])
log("\nэталон: строки без выдачи на №1–№2", f2(et_zero12_ss), "/", f2(et_zero12_last))

# разница по группам (по последней закупке и по СС)
def grp_delta(base):
    k = "uah_last" if base == "last" else "uah_ss"
    e_issued = sum(v[k] for v in ET.values() if not v["zero12"])
    e_zero = sum(v[k] for v in ET.values() if v["zero12"])
    z = {g: TB[g][k] for g in GROUPS}
    return {"zak": z, "et_issued_lines": e_issued, "et_lines_not_issued_1_2": e_zero,
            "delta_etalon_group_vs_issued_lines": z["еталон"] - e_issued, "delta_total": TB["всього"][k] - (e_issued + e_zero)}


GD_LAST, GD_SS = grp_delta("last"), grp_delta("ss")
log("разница групп (посл.)", {k: (f2(v) if isinstance(v, Decimal) else {a: f2(b) for a, b in v.items()}) for k, v in GD_LAST.items()})
log("разница групп (СС)   ", {k: (f2(v) if isinstance(v, Decimal) else {a: f2(b) for a, b in v.items()}) for k, v in GD_SS.items()})

# сравнение ключей: топ больше / меньше эталона (по цене последней закупки), без строк, не выданных на №1–№2
def key_row(x):
    out = {"key": x["key"], "name": x["name"], "group": x["group"], "zak_uah_last": f2(x["zak_last"]), "et_uah_last": f2(x["et_last"]),
           "delta_uah_last": f2(x["d_last"]), "zak_uah_ss": f2(x["zak_ss"]), "et_uah_ss": f2(x["et_ss"]), "delta_uah_ss": f2(x["d_ss"]),
           "line_not_issued_on_1_2": x["zero12"], "cards": []}
    e = ET.get(x["key"])
    if e:
        out["et_norm_line"] = f3(e["norm"]); out["et_unit"] = e["unit"]; out["et_buy_line_units"] = f3(e["buy"])
    for c in x["cards"]:
        r = RB[c]
        out["cards"].append({"code": c, "card": r["card"], "unit": r["unit"], "fact_1": f3(r["f1"]), "fact_2": f3(r["f2"]), "avg": f3(r["avg"]),
                             "etalon_norm": f3(r["etalon_norm"]) if r["etalon_norm"] is not None else None,
                             "stock_project": f3(r["stock_project"]), "buy_qty": f3(r["buy"]), "uah_last": f2(r["uah_last"]),
                             "unstable": r["unstable"]})
    return out


cmp_sorted = sorted(CMP, key=lambda x: -x["d_last"])
TOP_MORE = [key_row(x) for x in cmp_sorted if x["d_last"] > 0][:15]
TOP_LESS = [key_row(x) for x in sorted(CMP, key=lambda x: x["d_last"]) if x["d_last"] < 0 and not x["zero12"]][:15]
NOT_ISSUED = [key_row(x) for x in sorted(CMP, key=lambda x: x["d_last"]) if x["zero12"] and x["et_last"] > 0]
log("\nTOP больше эталона:")
for x in TOP_MORE:
    log(f"  {x['delta_uah_last']:>12,.2f}  {x['name'][:60]:60s} зак {x['zak_uah_last']:>11,.2f} эт {x['et_uah_last']:>11,.2f}  " +
        "; ".join(f"{c['card'][:30]} f1={c['fact_1']} f2={c['fact_2']} avg={c['avg']} N={c['etalon_norm']} sp={c['stock_project']} buy={c['buy_qty']}" for c in x["cards"][:3]))
log("\nTOP меньше эталона (строки, выданные на №1–№2):")
for x in TOP_LESS:
    log(f"  {x['delta_uah_last']:>12,.2f}  {x['name'][:60]:60s} зак {x['zak_uah_last']:>11,.2f} эт {x['et_uah_last']:>11,.2f}  " +
        "; ".join(f"{c['card'][:30]} f1={c['fact_1']} f2={c['fact_2']} avg={c['avg']} N={c['etalon_norm']} sp={c['stock_project']} buy={c['buy_qty']}" for c in x["cards"][:3]))
log("\nстроки эталона без выдачи на №1–№2 (закупочная = 0):", len(NOT_ISSUED), f2(sum(Dd(x["et_uah_last"]) for x in NOT_ISSUED)))
for x in NOT_ISSUED[:30]:
    log(f"  {x['et_uah_last']:>12,.2f} / СС {x['et_uah_ss']:>12,.2f}  {x['name']}  {x.get('et_buy_line_units')} {x.get('et_unit')}")

# ───────────────────────────── сверка (б) с полем «Остаток» РК №9 / №11
RKH = {"00000000009": "30-1", "00000000011": "30-2", "00000000012": "30-3"}
rk_ost = collections.defaultdict(lambda: Z)
for r in RAW["rk_ost"]:
    num = [x["Номер"] for x in RAW["rk"] if x["Док"] == r["Док"]]
    if num and num[0] in RKH:
        rk_ost[(RKH[num[0]], r["Код"])] += Dd(r["Остаток"])
recon = {}
for h in ("30-1", "30-2"):
    codes = {c for (hh, c) in rk_ost if hh == h} | {c for c in RB if (RB[c]["f1"] if h == "30-1" else RB[c]["f2"])}
    diff = []
    for c in codes:
        mine = (RB[c]["f1"] if h == "30-1" else RB[c]["f2"]) if c in RB else Z
        if c == KIT:
            continue
        rk = rk_ost.get((h, c), Z)
        if abs(mine - rk) > Decimal("0.0005"):
            diff.append((c, CARD.get(c, {}).get("name", "?"), float(mine), float(rk)))
    diff.sort()
    recon[h] = {"cards": len(codes), "diff": diff}
    log(f"сверка с РК {HN[h]}: карточек {len(codes)}, расхождений {len(diff)}", diff[:12])

# ───────────────────────────── септик: детали по факту против состава эталона
SEPTIC = []
for c, gp in KIT_CARDS.items():
    rec = {"code": c, "card": CARD[c]["name"] if c in CARD else c, "part": gp, "etalon_per_house": KIT_COMP[gp]}
    for h in ("30-1", "30-2", "30-3"):
        rec["cons_" + HN[h]] = f3(cons[c][h][0]); rec["asm_" + HN[h]] = f3(ASM[(c, h)][0]); rec["stock_" + HN[h]] = f3(stock_h(c, h)[0])
    rec["stock_project"] = f3(st[c]["P"][0])
    if c in RB:
        r = RB[c]
        rec.update({"avg": f3(r["avg"]), "need_3_7": f3(r["need_3_7"]), "buy": f3(r["buy"]), "uah_last": f2(r["uah_last"]), "uah_ss": f2(r["uah_ss"])})
    SEPTIC.append(rec)
log("\nСЕПТИК:")
for x in SEPTIC:
    log("  ", x)
log("  эталон: септик", ET["L:170"]["buy"], f2(ET["L:170"]["uah_ss"]), f2(ET["L:170"]["uah_last"]), "Sp", ET["L:170"]["Sp"])

# ───────────────────────────── неттинг аналогов по строке эталона (иллюстрация ограничения «плоской по карточкам»)
def analog_netting():
    """Если считать карточки одной строки СС взаимозаменяемыми: потребность и склад проекта — в ед. строки, по строке."""
    res = []
    for key, x in KB.items():
        if not key.startswith("L:") or key == "L:%d" % SEPTIC_LINE:
            continue
        n = int(key[2:])
        linked = [c for c, _, l in BY_LINE.get(n, [])]
        if len(linked) < 2:
            continue
        need_line = sum(RB[c]["need_3_7"] * RB[c]["coef"] for c in x["cards"] if RB[c].get("coef"))
        sp_line = sum(max(Z, st[c]["P"][0]) * l["coef"] for c, _, l in BY_LINE[n])
        buy_line = max(Z, need_line - sp_line)
        flat_line = sum(RB[c]["buy"] * RB[c]["coef"] for c in x["cards"] if RB[c].get("coef"))
        if flat_line - buy_line > Decimal("0.0005"):
            # оценка разницы по цене последней закупки основной карточки
            c0 = sorted(x["cards"], key=lambda c: -RB[c]["need_3_7"])[0]
            p = RB[c0]["p_last"] / RB[c0]["coef"] if RB[c0]["p_last"] and RB[c0]["coef"] else Z
            res.append({"line": n, "name": LINE[n]["text"], "unit": LINE[n]["unit"], "need_line": f3(need_line), "stock_P_all_cards_line": f3(sp_line),
                        "buy_flat_line": f3(flat_line), "buy_netted_line": f3(buy_line), "overstated_uah_last": f2((flat_line - buy_line) * p),
                        "cards": [{"code": c, "card": CARD[c]["name"], "stock_P": f3(st[c]["P"][0])} for c, _, l in BY_LINE[n]]})
    return res


AN = analog_netting()
log("\nаналоги: плоская по карточкам завышает (неттинг по строке):", f2(sum(Dd(x["overstated_uah_last"]) for x in AN)))
for x in sorted(AN, key=lambda x: -x["overstated_uah_last"])[:10]:
    log("   ", x["line"], x["name"], x["overstated_uah_last"], x["buy_flat_line"], "->", x["buy_netted_line"], [(c["card"][:25], c["stock_P"]) for c in x["cards"]])

# ───────────────────────────── ЦСП и ОСБ на №2: сценарий «не нужны на №2, вернуть»
cso = [c for c in RB if RB[c]["sh2"] > 0 and RB[c]["group"] != "МШП"]
log("\nлежит на складе №2:", [(c, RB[c]["card"], float(RB[c]["sh2"]), RB[c]["unit"]) for c in cso])
big = [c for c in cso if RB[c]["sh2"] * (RB[c]["p_last"] or Z) > 1000]
RS = compute("b", {c: {"avg_from": "a", "release_house_stock": True} for c in big})
RS_A = compute("b", {c: {"avg_from": "a"} for c in big})
TS, TS_A = tot(RS), tot(RS_A)
CSP_OSB = []
for c in big:
    b, a, s = RB[c], RA.get(c), RS.get(c)
    CSP_OSB.append({"code": c, "card": b["card"], "unit": b["unit"], "lies_on_2": f3(b["sh2"]), "fact_1": f3(b["f1"]), "fact_2_written": f3(b["fa2"]),
                    "avg_b": f3(b["avg"]), "avg_a": f3(a["avg"]) if a else 0.0, "stock_project": f3(b["stock_project"]),
                    "buy_b": f3(b["buy"]), "uah_b": f2(b["uah_last"]),
                    "buy_avg_a_no_return": f3(RS_A[c]["buy"]) if c in RS_A else 0.0, "uah_avg_a_no_return": f2(RS_A[c]["uah_last"]) if c in RS_A else 0.0,
                    "buy_avg_a_return": f3(s["buy"]) if s else 0.0, "uah_avg_a_return": f2(s["uah_last"]) if s else 0.0,
                    "price_last": f2(b["p_last"])})
    log("  ", CSP_OSB[-1])
log("сценарии ЦСП/ОСБ: (б) основной", f2(TB["всього"]["uah_last"]), "; среднее по (а) для них, остаются на №2", f2(TS_A["uah_last"]),
    "; среднее по (а) + вернуть на склад проекта", f2(TS["uah_last"]))

# ───────────────────────────── №2 позже: что легло на №2 больше, чем на №1
LATER2 = []
for c, r in RB.items():
    d = (r["f2"] - r["f1"]) * (r["p_last"] or Z)
    if d > 500:
        LATER2.append({"code": c, "card": r["card"], "unit": r["unit"], "group": r["group"], "fact_1": f3(r["f1"]), "fact_2": f3(r["f2"]),
                       "delta_uah_last": f2(d), "only_on_2": r["f1"] == 0, "buy_effect_uah_last": f2((r["f2"] - r["avg"]) * 4 * (r["p_last"] or Z))})
LATER2.sort(key=lambda x: -x["delta_uah_last"])
log("\n№2 > №1 (топ):", len(LATER2), f2(sum(Dd(x["delta_uah_last"]) for x in LATER2)))
for x in LATER2[:15]:
    log("  ", x)
ONLY1 = sorted([{"code": c, "card": r["card"], "unit": r["unit"], "group": r["group"], "fact_1": f3(r["f1"]), "uah_avg": f2(r["avg"] * (r["p_last"] or Z))}
                for c, r in RB.items() if r["f2"] == 0 and r["f1"] > 0], key=lambda x: -x["uah_avg"])
ONLY2 = sorted([{"code": c, "card": r["card"], "unit": r["unit"], "group": r["group"], "fact_2": f3(r["f2"]), "uah_avg": f2(r["avg"] * (r["p_last"] or Z))}
                for c, r in RB.items() if r["f1"] == 0 and r["f2"] > 0], key=lambda x: -x["uah_avg"])
log("только на №1:", len(ONLY1), f2(sum(Dd(x["uah_avg"]) for x in ONLY1)), ONLY1[:8])
log("только на №2:", len(ONLY2), f2(sum(Dd(x["uah_avg"]) for x in ONLY2)), ONLY2[:8])

# нестабильные
UNST = sorted([r for r in RB.values() if r["unstable"]], key=lambda r: -r["uah_last"])
log("\nнестабильных карточек:", len(UNST), "из", len(RB), "; их докупка (посл.)", f2(sum(r["uah_last"] for r in UNST)),
    "; их среднее на дом", f2(sum(r["avg"] * (r["p_last"] or Z) for r in UNST)))
for r in UNST[:15]:
    log(f"   {r['card'][:50]:50s} {r['unit']:4s} f1={float(r['f1']):g} f2={float(r['f2']):g} cv={float(r['cv']):.2f} buy={float(r['buy']):g} uah={f2(r['uah_last']):,.2f} {r['group']}")

# №3: излишек на складе №3 сверх среднего
S3 = sorted([r for r in RB.values() if r["surplus3"] > 0], key=lambda r: -(r["surplus3"] * (r["p_last"] or Z)))
log("\n№3 сверх среднего (списано+лежит):", len(S3), f2(sum(r["surplus3"] * (r["p_last"] or Z) for r in S3)))
for r in S3[:10]:
    log(f"   {r['card'][:50]:50s} {r['unit']} avg={float(r['avg']):g} f3={float(r['f3']):g} s3={float(r['s3']):g} surplus={float(r['surplus3']):g} uah={f2(r['surplus3'] * (r['p_last'] or Z)):,.2f}")

# округление
rnd_q = sum(r["buy_house_round"] - r["buy"] for r in RB.values() if r["piece"])
log("\nокругление: построчно по домам vs один раз на итог: штук", float(rnd_q), "грн",
    f2(TB["всього"]["uah_last_house_round"] - TB["всього"]["uah_last"]),
    "; без округления вовсе", f2(sum(r["buy_raw"] * (r["p_last"] or Z) for r in RB.values())))
RND_TOP = sorted([r for r in RB.values() if r["piece"] and r["uah_last_house_round"] > r["uah_last"]],
                 key=lambda r: -(r["uah_last_house_round"] - r["uah_last"]))[:10]
for r in RND_TOP:
    log(f"   {r['card'][:50]:50s} avg={float(r['avg']):g} buy_once={float(r['buy']):g} buy_per_house={float(r['buy_house_round']):g} +{f2(r['uah_last_house_round'] - r['uah_last']):,.2f}")

# цены: откуда
basis_cnt = collections.Counter((r["basis_last"]) for r in RB.values() if r["buy"] > EPS)
basis_uah = collections.defaultdict(lambda: Z)
for r in RB.values():
    if r["buy"] > EPS:
        basis_uah[r["basis_last"]] += r["uah_last"]
log("\nоснова цены (посл.):", dict(basis_cnt), {k: f2(v) for k, v in basis_uah.items()})
basis_ss_uah = collections.defaultdict(lambda: Z)
for r in RB.values():
    if r["buy"] > EPS:
        basis_ss_uah[r["basis_ss"]] += r["uah_ss"]
log("основа цены (СС):", {k: f2(v) for k, v in basis_ss_uah.items()})

# среднее на дом против нормы эталона на дом
avg_house_last = TB["всього"]["avg_house_uah"]
log("\nсреднее на дом по закупочной (посл. закупка):", f2(avg_house_last), "; норма эталона на дом СС 1 009 864,74")
log("среднее на дом по факту брутто партий: №1", f2(sum(r["f1g"] for r in RB.values())), "№2", f2(sum(r["f2g"] for r in RB.values())))

# ───────────────────────────── JSON
def jrow(r):
    return {"card": r["card"], "code": r["code"], "unit": r["unit"], "piece": r["piece"], "on_card": r["on_card"], "group": r["group"], "kind": r["kind"],
            "ss_line": r["ss_line"], "method": r["method"], "compare_key": r["key"],
            "fact_1": f3(r["f1"]), "fact_2": f3(r["f2"]), "fact_1_spysano": f3(r["fa1"]), "fact_2_spysano": f3(r["fa2"]),
            "lies_house_1": f3(r["sh1"]), "lies_house_2": f3(r["sh2"]), "fact_1_uah": f2(r["f1g"]), "fact_2_uah": f2(r["f2g"]),
            "avg": f3(r["avg"]), "avg_variant_a": f3(RA[r["code"]]["avg"]) if r["code"] in RA else 0.0,
            "cv": round(float(r["cv"]), 3), "unstable": r["unstable"],
            "etalon_norm": f3(r["etalon_norm"]) if r["etalon_norm"] is not None else None, "etalon_norm_shared_by_cards": r["etalon_shared"],
            "coef_to_ss_unit": float(r["coef"]) if r.get("coef") is not None else None,
            "fact_3_spysano": f3(r["f3"]), "lies_house_3": f3(r["s3"]), "gap_3": f3(r["gap3"]), "surplus_3": f3(r["surplus3"]),
            "need_3_7": f3(r["need_3_7"]), "stock_project": f3(r["stock_project"]), "stock_houses_4_7": f3(r["stock_4_7"]),
            "buy_qty_raw": f3(r["buy_raw"]), "buy_qty": f3(r["buy"]), "buy_qty_if_rounded_per_house": f3(r["buy_house_round"]),
            "price_last": f2(r["p_last"]) if r["p_last"] else None, "price_last_raw_ptu": f2(r["price_last_raw"]) if r["price_last_raw"] else None,
            "price_ss": f2(r["price_ss"]) if r["price_ss"] else None, "fact_cost_1_2": f2(r["fact_cost"]) if r["fact_cost"] else None,
            "price_basis_last": r["basis_last"], "price_basis_ss": r["basis_ss"],
            "buy_uah_last": f2(r["uah_last"]), "buy_uah_ss": f2(r["uah_ss"]),
            "by_house": {HN[h]: {"need": f3(v["need"]), "from_project": f3(v["from_project"]), "own_house_stock": f3(v["own"]), "buy": f3(v["buy"]),
                                 "uah_last": f2(v["uah_last"]), "uah_ss": f2(v["uah_ss"])} for h, v in r["byh"].items()},
            "note": r["note"]}


rows_json = [jrow(r) for r in sorted(RB.values(), key=lambda r: (GROUPS.index(r["group"]), -r["uah_last"], r["card"]))]

shah_h = DEMO["vs_budget_v3"]["now_by_house"]; et_h = DEMO["vs_budget_v3"]["mapped_by_house"]; v3_h = DEMO["vs_budget_v3"]["v3_by_house"]
zak_h_last = {HN[h]: f2(v) for h, v in TB["всього"]["by_house_last"].items()}
zak_h_ss = {HN[h]: f2(v) for h, v in TB["всього"]["by_house_ss"].items()}
et_h_last = {HN[h]: f2(sum(v["byh_last"][h] for v in ET.values())) for h in REST}
now_h_last = {HN[h]: f2(sum(x["buy"] * (v["price_last"] if v["price_last"] else v["price_ss"]) for v in NOW_BUY.values() for hh, x in v["byh"].items() if hh == h)) for h in REST}

ALLC = [c for c in CARD if c != KIT]
res = {
    "snapshot": {"datetime": RAW["stamp"], "source": "raw.json прототипу зв'язки (BuhBud bas_industrialbud, лише читання); рухи 20/22/26/28 по складах IRC — %d рядків, приходи — %d" % (len(MV), len(RAW["receipts"])),
                 "card": "СС 000000009 «МД IRС 2026 30 м», 7 домів", "project_warehouse": "ІБ _МД IRC 2026",
                 "basis": "грн з ПДВ; факт і склад — кількості карток BuhBud; ціна останньої закупівлі — ПТУ BuhBud (нетто + ПДВ) / кількість; ціна СС — Комплектующие.Сумма / Кол строки СС × коефіцієнт зв'язки"},
    "method": {
        "spec": "Закупівельна специфікація — плоска, по картках, без вузлів: кількість на дім = середнє по списаних домах типу (№1, №2); дім без списання картки дає 0. Включено все реально списане: позиції еталона, «поза еталоном», МШП (окрема група).",
        "fact": "Факт картки на дім, варіант (б) ОСНОВНИЙ = списано на дім (Кт матеріального рахунку зі складу дому на не матеріальний; для складових септика + зібране в напівфабрикат «Комлпект септика» на складі дому) + лежить на складах дому (= поле «Остаток» РК №9/№11). Варіант (а) — лише списано.",
        "avg": "Середнє = (факт №1 + факт №2) / 2 по кожній картці; карточка «Комлпект септика» (напівфабрикат) виключена — рахуються її складові.",
        "need": "Потреба №3–№7 = середнє × 4 (№4–№7) + по №3 max(0, середнє − (списано №3 + лежить на складі №3)).",
        "netting": "Докупка = max(0, потреба − склад проєкту «ІБ _МД IRC 2026» (сальдо картки на знімок) − склади №4–№7 (зараз порожні)). Неттинг по картці, аналоги не підміняють одне одного (див. compare.analog_netting).",
        "rounding": "Штучні картки (шт, рул, пак, пар) — округлення вгору ОДИН раз на підсумок картки; по домах — через накопичений підсумок (ceil(cum) − ceil(cum_prev)), щоб сума по домах = підсумку. Вагові/метрові — без округлення.",
        "price": "Оцінка: ціна останньої закупівлі картки з ПДВ (ПТУ BuhBud). Немає закупівель → ціна СС через зв'язку з позицією еталона (позначка price_basis_last). Немає ні того, ні іншого → собівартість факту №1–№2 з ПДВ партії. Друга оцінка — за ціною СС (для «поза еталоном» і МШП ціни СС немає → остання закупівля).",
        "fifo": "Розкладка по домах №3→№7: склад проєкту віддається FIFO у порядку №3, №4, №5, №6, №7 (як G_h бюджету v3 і прототип).",
        "compare": "Порівняння з еталоном «зі зв'язкою» — по ключу: строка СС (картки еталона через першу зв'язку; групова норма — одна група; складові септика — строка 170) або картка («поза еталоном», МШП).",
    },
    "houses_base": {
        "base": [HN[h] for h in BASE_H], "rest": [HN[h] for h in REST],
        "houses": [{"n": HN[h], "fact_b_uah": f2(sum(fact_b(c, h)[1] for c in ALLC)), "spysano_uah": f2(sum(fact_a(c, h)[1] for c in ALLC)),
                    "lies_uah": f2(sum(stock_h(c, h)[1] for c in ALLC)), "cards": sum(1 for c in ALLC if fact_b(c, h)[0] > 0),
                    "not_in_spec_uah": f2(sum(fact_b(c, h)[1] for c in ALLC if c not in RB))} for h in ("30-1", "30-2", "30-3")],
        "note": "№1 і №2 відвантажені й списані (РК №9, №11); на №2 на складі дому лежить 99 649,86 з ПДВ партії: ЦСП 9 і ОСБ 50 (98 816,96; переміщені 10.09 ІБ00-002211 на склад «ІС МД IRC 2026 30 м №2 Виробничий») та дрібниця — шуруп 750, свердло 10 мм, кліпса 10, стрічка 1. №3 частково: списано 120 727,54, лежить 197 681,57. №4–№7 не розпочаті. Усе — на знімок 03.10 10:52 (зміни після знімка — verification.after_snapshot).",
        "rk_reconciliation": {HN[h]: {"cards": v["cards"], "diff_count": len(v["diff"]), "diff": [{"code": a, "card": b, "mine": c, "rk_ostatok": d} for a, b, c, d in v["diff"]]} for h, v in recon.items()},
    },
    "rows": rows_json,
    "totals": {
        "variant_b_main": {g: jt(t) for g, t in TB.items()},
        "variant_a_spysano_only": {g: jt(t) for g, t in TA.items()},
        "delta_b_minus_a_uah_last": f2(TB["всього"]["uah_last"] - TA["всього"]["uah_last"]),
        "rounding": {"once_on_total_uah_last": f2(TB["всього"]["uah_last"]), "per_house_uah_last": f2(TB["всього"]["uah_last_house_round"]),
                     "delta_uah_last": f2(TB["всього"]["uah_last_house_round"] - TB["всього"]["uah_last"]), "delta_pieces": f3(rnd_q),
                     "no_rounding_uah_last": f2(sum(r["buy_raw"] * (r["p_last"] or Z) for r in RB.values())),
                     "top": [{"card": r["card"], "avg": f3(r["avg"]), "buy_once": f3(r["buy"]), "buy_per_house": f3(r["buy_house_round"]),
                              "delta_uah": f2(r["uah_last_house_round"] - r["uah_last"])} for r in RND_TOP]},
        "price_basis_last": {k: {"cards": basis_cnt[k], "uah": f2(v)} for k, v in basis_uah.items()},
        "price_basis_ss": {k: f2(v) for k, v in basis_ss_uah.items()},
        "avg_per_house_uah_last": f2(avg_house_last),
        "by_house_uah_last": zak_h_last, "by_house_uah_ss": zak_h_ss,
    },
    "compare": {
        "dokupyty_3_7": {
            "zakupochnaya_uah_last": f2(TB["всього"]["uah_last"]), "zakupochnaya_uah_ss": f2(TB["всього"]["uah_ss"]),
            "shahmatka_uah_ss": DEMO["totals"]["now"]["dokupyty_3_7_uah"], "shahmatka_uah_last": DEMO["totals"]["now"]["dokupyty_3_7_uah_last_price"],
            "etalon_mapped_uah_ss": DEMO["totals"]["mapped"]["dokupyty_3_7_uah"], "etalon_mapped_uah_last": DEMO["totals"]["mapped"]["dokupyty_3_7_uah_last_price"],
            "budget_v3_uah": V3["v3_G_h_total"],
            "delta_vs_shahmatka_last": f2(TB["всього"]["uah_last"] - Dd(DEMO["totals"]["now"]["dokupyty_3_7_uah_last_price"])),
            "delta_vs_shahmatka_ss": f2(TB["всього"]["uah_ss"] - Dd(DEMO["totals"]["now"]["dokupyty_3_7_uah"])),
            "delta_vs_etalon_last": f2(TB["всього"]["uah_last"] - Dd(DEMO["totals"]["mapped"]["dokupyty_3_7_uah_last_price"])),
            "delta_vs_etalon_ss": f2(TB["всього"]["uah_ss"] - Dd(DEMO["totals"]["mapped"]["dokupyty_3_7_uah"])),
            "delta_vs_v3_ss": f2(TB["всього"]["uah_ss"] - Dd(V3["v3_G_h_total"])),
            "delta_vs_v3_last": f2(TB["всього"]["uah_last"] - Dd(V3["v3_G_h_total"])),
        },
        "by_house": {"zakupochnaya_uah_last": zak_h_last, "zakupochnaya_uah_ss": zak_h_ss, "shahmatka_uah_ss": shah_h, "shahmatka_uah_last": now_h_last,
                     "etalon_mapped_uah_ss": et_h, "etalon_mapped_uah_last": et_h_last, "budget_v3_uah": v3_h},
        "by_group": {"basis_last": {k: (f2(v) if isinstance(v, Decimal) else {a: f2(b) for a, b in v.items()}) for k, v in GD_LAST.items()},
                     "basis_ss": {k: (f2(v) if isinstance(v, Decimal) else {a: f2(b) for a, b in v.items()}) for k, v in GD_SS.items()},
                     "note": "Еталон «зі зв'язкою» весь лежить у групі «еталон»; et_lines_not_issued_1_2 — строки СС, по яких на №1–№2 не списано і не лежить жодної одиниці (у закупівельній їх немає)."},
        "top15_more_than_etalon": TOP_MORE, "top15_less_than_etalon": TOP_LESS,
        "etalon_lines_not_issued_on_1_2": NOT_ISSUED,
        "septic": {"parts": SEPTIC, "etalon_kit": {"per_house": {gp: per for gp, per in KIT_COMP.items()}, "etalon_buy_kits": f3(ET["L:170"]["buy"]),
                                                    "etalon_uah_ss": f2(ET["L:170"]["uah_ss"]), "etalon_uah_last": f2(ET["L:170"]["uah_last"])},
                   "zakupochnaya_uah_last": f2(sum(RB[c]["uah_last"] for c in KIT_CARDS if c in RB)),
                   "zakupochnaya_uah_ss": f2(sum(RB[c]["uah_ss"] for c in KIT_CARDS if c in RB))},
        "analog_netting": {"overstated_uah_last": f2(sum(Dd(x["overstated_uah_last"]) for x in AN)), "lines": AN},
    },
    "caveats": [],
}

# сценарии для оговорок
# чувствительность к выборке из двух домов: мин / макс по №1–№2 вместо среднего
R_MIN, R_MAX = compute("b", est="min"), compute("b", est="max")
T_MIN, T_MAX = tot(R_MIN), tot(R_MAX)
res["compare"]["sensitivity_two_houses"] = {"min_f1_f2_uah_last": f2(T_MIN["uah_last"]), "avg_uah_last": f2(TB["всього"]["uah_last"]),
                                            "max_f1_f2_uah_last": f2(T_MAX["uah_last"]), "min_f1_f2_uah_ss": f2(T_MIN["uah_ss"]),
                                            "avg_uah_ss": f2(TB["всього"]["uah_ss"]), "max_f1_f2_uah_ss": f2(T_MAX["uah_ss"]),
                                            "note": "якби кількість на дім брали мінімум / максимум із №1 і №2 замість середнього — діапазон невизначеності вибірки з двох домів"}
log("\nчувствительность мин/сред/макс (посл.):", f2(T_MIN["uah_last"]), f2(TB["всього"]["uah_last"]), f2(T_MAX["uah_last"]),
    "; СС:", f2(T_MIN["uah_ss"]), f2(TB["всього"]["uah_ss"]), f2(T_MAX["uah_ss"]))
# гибрид: закупочная + строки эталона, по которым на №1–№2 не выдано ни одной единицы
ZERO_KEYS = [k for k, v in ET.items() if v["zero12"]]
HYB_LAST = TB["всього"]["uah_last"] + et_zero12_last
HYB_SS = TB["всього"]["uah_ss"] + et_zero12_ss
res["compare"]["hybrid_zakupochnaya_plus_etalon_lines_not_issued"] = {
    "uah_last": f2(HYB_LAST), "uah_ss": f2(HYB_SS),
    "by_house_uah_last": {HN[h]: f2(TB["всього"]["by_house_last"][h] + sum(ET[k]["byh_last"][h] for k in ZERO_KEYS)) for h in REST},
    "by_house_uah_ss": {HN[h]: f2(TB["всього"]["by_house_ss"][h] + sum(ET[k]["byh_ss"][h] for k in ZERO_KEYS)) for h in REST},
    "etalon_lines_not_issued_uah_last": f2(et_zero12_last), "etalon_lines_not_issued_uah_ss": f2(et_zero12_ss),
    "note": "Закупівельна не бачить строк СС, яких на №1–№2 не видали (ПК 10 стіни, бойлер 82 л, банер, стільці, підвіконня…): або їх ставлять пізніше/купують напряму, або вони зайві в СС. До рішення людини — додавати в бюджет за еталоном."}
res["totals"]["plan_month"] = {"№3": "2026-10", "№4": "2026-10", "№5": "2026-10", "№6": "2026-10", "№7": "2026-11"}
res["totals"]["price_flags"] = [{"code": c, "card": r["card"], "unit": r["unit"], "price_used": f2(r["p_last"]) if r["p_last"] else None,
                                 "price_last_raw_ptu": f2(r["price_last_raw"]) if r["price_last_raw"] else None,
                                 "fact_cost_1_2": f2(r["fact_cost"]) if r["fact_cost"] else None, "basis": r["basis_last"], "buy_qty": f3(r["buy"]),
                                 "uah_last": f2(r["uah_last"])} for c, r in RB.items() if r["basis_last"] != "остання закупівля"]
_ssfb = [r for r in RB.values() if r["basis_last"].startswith("ціна СС") and r["buy"] > EPS]
res["totals"]["ss_fallback_vs_fact_cost"] = {"cards": len(_ssfb), "uah_at_ss": f2(sum(r["uah_last"] for r in _ssfb)),
                                             "uah_at_fact_cost": f2(sum(r["buy"] * r["fact_cost"] for r in _ssfb if r["fact_cost"])),
                                             "note": "картки без закупівель, оцінені за ціною СС (рішення 03.10); якби оцінювали собівартістю факту №1–№2 — друге число"}
log("\nSS-fallback:", res["totals"]["ss_fallback_vs_fact_cost"])
log("\nгибрид (закупочная + строки эталона без выдачи):", f2(HYB_LAST), "/ СС", f2(HYB_SS))
for x in res["totals"]["price_flags"]:
    log("  цена:", x)
res["compare"]["csp_osb_on_2"] = {"cards": CSP_OSB, "total_b_main_uah_last": f2(TB["всього"]["uah_last"]),
                                  "total_avg_a_for_them_stay_on_2_uah_last": f2(TS_A["uah_last"]),
                                  "total_avg_a_for_them_return_to_project_uah_last": f2(TS["uah_last"])}
res["compare"]["house_2_more_than_1"] = {"count": len(LATER2), "uah_last": f2(sum(Dd(x["delta_uah_last"]) for x in LATER2)), "top": LATER2[:20],
                                         "only_on_1": ONLY1[:20], "only_on_1_count": len(ONLY1), "only_on_2": ONLY2[:20], "only_on_2_count": len(ONLY2)}
res["compare"]["unstable"] = {"cv_limit": float(CV_LIMIT), "count": len(UNST), "of_cards": len(RB), "buy_uah_last": f2(sum(r["uah_last"] for r in UNST)),
                              "avg_house_uah_last": f2(sum(r["avg"] * (r["p_last"] or Z) for r in UNST)),
                              "top": [{"card": r["card"], "unit": r["unit"], "fact_1": f3(r["f1"]), "fact_2": f3(r["f2"]), "cv": round(float(r["cv"]), 2),
                                       "buy_qty": f3(r["buy"]), "uah_last": f2(r["uah_last"]), "group": r["group"]} for r in UNST[:25]]}
res["compare"]["house_3_surplus"] = {"cards": len(S3), "uah_last": f2(sum(r["surplus3"] * (r["p_last"] or Z) for r in S3)),
                                     "top": [{"card": r["card"], "unit": r["unit"], "avg": f3(r["avg"]), "spysano_3": f3(r["f3"]), "lies_3": f3(r["s3"]),
                                              "surplus": f3(r["surplus3"]), "uah_last": f2(r["surplus3"] * (r["p_last"] or Z))} for r in S3[:15]]}

def F(x):
    return f"{f2(x):,.2f}".replace(",", " ").replace(".", ",")


_door = RB.get("00-00010822"); _osb = RB.get("00000013734"); _dn = RB.get("00-00010526")
res["caveats"] = [
    "Розрахунок у пам'яті Python по знімку raw.json (BuhBud 03.10.2026 10:52); у базу нічого не записано, конфігурацію не змінено.",
    f"Вибірка — два доми. Нестабільних карток (коефіцієнт варіації №1/№2 > {float(CV_LIMIT)}): {len(UNST)} з {len(RB)}; їх докупка {F(sum(r['uah_last'] for r in UNST))} за останньою закупівлею. "
    f"Діапазон невизначеності: якби брали мінімум із №1/№2 — {F(T_MIN['uah_last'])}, максимум — {F(T_MAX['uah_last'])} (середнє {F(TB['всього']['uah_last'])}).",
    f"№2 оформлено пізніше (комплектації ІБ00-000154/155 від 01.10): на ньому дві вхідні двері (на №1 і в еталоні — одні) — середнє 1,5 дає докупку 4 двері = {F(_door['uah_last']) if _door else '—'}; "
    f"якщо друга двері помилкова — 1 двері = 22 037,15. На №2 є кондиціонер, лист рівний 44 м², лист ПВЛ, плівка пароізоляційна, яких немає на №1; на №1 є 7 обігрівачів, пінопласт 50 мм 132 шт, ґрунт, кутник 50×50, яких немає на №2 — це «незавершеність» одного з домів, а не норма. Після знімка (переміщення ІБ00-002378 від 01.10 20:00, проведене пізніше 10:52) 7 обігрівачів передано зі складу проєкту на №2 — на живій базі обігрівачі є на обох домах (див. verification.after_snapshot).",
    f"ЦСП 9 і ОСБ 50 лежать на складі №2 «ІС МД IRC 2026 30 м №2 Виробничий» (98 816,96 з ПДВ; разом із дрібницею на №2 лежить 99 649,86). Варіант (б) рахує їх як факт №2 (ОСБ 50 = факт №1 і норма еталона 50 шт): докупка ОСБ 104 шт = {F(_osb['uah_last']) if _osb else '—'}. "
    f"Якщо на №2 вони не потрібні: середнє за (а) — разом {F(TS_A['uah_last'])}; за (а) і повернути на склад проєкту — {F(TS['uah_last'])}. ЦСП докупки не дає в жодному варіанті (на складі проєкту 37 шт).",
    f"Септик: на №1 днище зібране в напівфабрикат двічі (ІБ00-000174 і ІБ00-000175) — середнє 1,5 днища на дім і докупка 6 днищ = {F(_dn['uah_last']) if _dn else '—'}; при 1 днищі на дім — 3 днища = 9 644,98, і закупівельна по септику дорівнює еталону (39 830,02 за останньою закупівлею).",
    "Закупівельна не бачить строк СС, яких на №1–№2 не видали жодної одиниці (25 строк: ПК 10 стіни 140 500,00, бойлер 82 л 46 000,00, банер 32 000,00, стільці, підвіконня, гофра, сигналізація…). Для бюджету до рішення людини їх треба додати за еталоном (гібрид).",
    "Склад проєкту «ІБ _МД IRC 2026» неттується повністю під 30 м (як у Шахматці, прототипі і v3). Будинки 15 м №1–№6 вже відвантажені, тому конфлікт малий, але склад проєкту спільний.",
    f"Ціни: остання закупівля перевірена на помилку одиниці в ПТУ (відхилення від собівартості факту > 5 разів): скоба для степлера (остання ПТУ 15 «шт» за 1 605 грн — фактично упаковки) — взята попередня закупівля; свердло 5,0×86 — собівартість факту. "
    f"Картки без закупівель за ціною СС: {res['totals']['ss_fallback_vs_fact_cost']['cards']} шт, {F(Dd(res['totals']['ss_fallback_vs_fact_cost']['uah_at_ss']))}; собівартістю факту було б {F(Dd(res['totals']['ss_fallback_vs_fact_cost']['uah_at_fact_cost']))} (ціни СС на дрібницю — шайби, гайки М12, кліпси — у рази вищі за ринок).",
    "Порівняння з еталоном «зі зв'язкою» успадковує одиничні помилки прототипу: «Ґрунт 10кг» у картці в кг, але закупівля і списання — відрами (≈560 грн за одиницю), тому еталон для строки «Грунт під клей» (40 кг на дім) дає 78 470,78 докупки, а за фактом склад 60 «кг» (відер) закриває все. Кутник 30×30: ціна СС 18,67 грн/м проти 68 грн/м останньої закупівлі.",
    f"Плоска специфікація по картках не підміняє аналоги: якщо картки однієї строки СС взаємозамінні, неттинг по строці зменшив би докупку на {F(Dd(res['compare']['analog_netting']['overstated_uah_last']))} (найбільше — з'єднувачі ринви/труби, які насправді не взаємозамінні, тож це верхня межа).",
    "МШП (інструмент, свердла, диски, валики) включено окремою групою за рішенням 03.10: інструмент багаторазовий (пістолет скобозабивний, коронка), тому його середнє на дім — завищення, а не потреба кожного дому.",
    f"Округлення: один раз на підсумок картки дає {F(TB['всього']['uah_last'])}; якби округлювали по кожному дому — {F(TB['всього']['uah_last_house_round'])} (+{F(TB['всього']['uah_last_house_round'] - TB['всього']['uah_last'])}); без округлення — {F(Dd(res['totals']['rounding']['no_rounding_uah_last']))}.",
]

_PV = os.path.join(HERE, "z1_proverka.json")          # результати незалежної перевірки 03.10 (статичні факти, COM BuhBud)
if os.path.exists(_PV):
    res["verification"] = json.load(open(_PV, encoding="utf-8"))
json.dump(res, open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
open(OUT_TXT, "w", encoding="utf-8").write("\n".join(LOG))
log("\nJSON:", OUT_JSON)
