# -*- coding: utf-8 -*-
"""Z6. Закупки проекта «МД IRC 2026» целиком (15 м №1–№6 + 30 м №1–№7) по карточкам. ТОЛЬКО ЧТЕНИЕ: база не трогается,
читается свежий снимок raw.json (z6_collect.py), прототип связи (base / mapping / views) работает на нём.

Что считается (по каждой карточке номенклатуры):
  need_15      — потребность 15 м = факт 6 домов 15 м (списано со складов дома + лежит на складах дома); тип закрыт, РК не пересчитываются;
  need_30      — потребность 30 м = факт №1 + факт №2 (списано + лежит; складові септика по прототипу) + №3: max(середнє, факт №3)
                 + середнє × 4 (№4–№7) + 25 строк эталона, не выданных на №1–№2, по норме СС на №3–№7 (решение 7);
  written_off  — списано на дома (Кт 20/22/28 со склада дома → Дт не материальный счёт), по домам;
  stock        — лежит: склады домов (по дому) и склад проекта (класс P прототипа: «ІБ _МД IRC 2026», «ІС _МД IRC 2026», …);
  purchased    — закупки на склады проекта и домов: ПТУ («Надходження товарів і послуг»), «Надходження дод. витрат» (только сумма),
                 «Авансовий звіт»; кол-во, без ПДВ (Хозрасчетный) и з ПДВ (ПТУ / ПДВ партії);
  other_in/out — прочие приходы/расходы системы складов IRC: перемещения с/на склады вне проекта (ІБ_МШП, склады ІНДЕПТ СТІЛ, другие объекты),
                 списания со склада проекта не на дом, возвраты, перетворення карток (комплектація);
  balance_diff — сверка замкнутости: закуплено + прочие приходы − списано − прочие расходы − остаток Хозрасчетный.Остатки (независимый запрос);
  to_buy_30    — ещё докупить для 30 м №3–№7 = max(0, остаток потребности − склад проекта − склады №4–№7), штучные — округление вверх
                 один раз на итог карточки, по домам — накопленным итогом, FIFO склада проекта №3 → №7 (как Z1); месяцы: №3–№6 — 2026-10, №7 — 2026-11;
  surplus      — останется лишним = склад проекта (+ склады №4–№7) после покрытия потребности 30 м.
Цены: последняя закупка — средневзвешенная по последнему ПТУ карточки с проверкой «×5 к себестоимости факта» (правило Z1 §1.4);
нет закупок → цена СС через связь (з ПДВ; без ПДВ = ÷1,2); нет ни того ни другого → себестоимость факта. Вторая оценка — цена СС.
Запуск: python z6_calc.py  (Z6_RAW_DIR — папка с raw.json и копиями base/mapping/views; по умолчанию z6_tmp/calc).
Выход: data/proekt_irc2026_zakupki.json и z6_out.txt рядом со скриптом."""
import os, sys, json, math, collections, re
from decimal import Decimal

RAWDIR = os.environ.get("Z6_RAW_DIR", r"C:/Users/SUPPOR~1/AppData/Local/Temp/claude/z6_tmp/calc")
sys.path.insert(0, RAWDIR)
from views import *            # noqa: E402,F401,F403  (base + mapping + views прототипа на свежем снимке)
from base import _d            # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))           # «Разработка связей»
DATA = os.path.join(ROOT, "data")
Z1J = json.load(open(os.path.join(DATA, "zakupochnaya_spec_30m.json"), encoding="utf-8"))
Z3J = json.load(open(os.path.join(DATA, "demo_specifikaciya_15m.json"), encoding="utf-8"))
OUT_JSON = os.path.join(DATA, "proekt_irc2026_zakupki.json")
OUT_TXT = os.path.join(HERE, "z6_out.txt")

VATK = Decimal("1.2")
H15 = [f"15-{i}" for i in range(1, 7)]
H30 = [f"30-{i}" for i in range(1, 8)]
BASE_H = ["30-1", "30-2"]
REST = ["30-3", "30-4", "30-5", "30-6", "30-7"]
NOT_STARTED = ["30-4", "30-5", "30-6", "30-7"]
MONTH = {"30-3": "2026-10", "30-4": "2026-10", "30-5": "2026-10", "30-6": "2026-10", "30-7": "2026-11"}
HN = {**{f"30-{i}": f"30 м №{i}" for i in range(1, 8)}, **{f"15-{i}": f"15 м №{i}" for i in range(1, 7)}}
PIECE = {"шт", "рул", "пак", "пар", "компл", "компл.", "уп", "упак", "лист", "набір", "балон"}
CV_LIMIT = Decimal("0.3")
EPS = Decimal("0.0000001")
MSHP_NODE = "МШП та інструмент (накладні, не вузол будинку)"
LOG = []


def log(*a):
    s = " ".join(str(x) for x in a)
    LOG.append(s)
    print(s)


f2, f3 = r2, r3


def f6(x):
    return float(Dd(x).quantize(Decimal("0.000001")))


# ───────────────────────────── карточки: справочник снимка (единица, ставка)
CARDX = {c["Код"]: c for c in RAW.get("cards", [])}


def cname(c):
    return (CARD.get(c, {}).get("name") or CARDX.get(c, {}).get("Имя") or c).strip()


def unit(c):
    return ((CARD.get(c, {}).get("unit") or "") or (CARDX.get(c, {}).get("Ед") or "")).strip()


def is_piece(c):
    return unit(c) in PIECE


def ceil_q(x):
    return Decimal(math.ceil(x - EPS)) if x > 0 else Z


S_CLS = lambda cl: cl == "P" or cl.startswith("H:")


# ═════════════════════════════ ЧАСТЬ A. Закупочная 30 м (повтор правил Z1 на свежем снимке)
def fact_a(c, h):
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


PRICE_K = Decimal(5)
_RCD = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z]))
for _r in RAW["receipts"]:
    if not _r["Проведен"] or not Dd(_r["Кол"]):
        continue
    _net = Dd(_r["Сумма"]) - (Dd(_r["НДС"]) if _r["ВклНДС"] else Z)
    _a = _RCD[_r["Код"]][(_d(_r["Дата"]), _r["Док"])]
    _a[0] += _net + Dd(_r["НДС"]); _a[1] += Dd(_r["Кол"]); _a[2] += _net
RCPT = collections.defaultdict(list)
for _c, _docs in _RCD.items():
    for (_dt, _doc), (_g, _q, _n) in _docs.items():
        if _q:
            RCPT[_c].append({"date": _dt, "price_g": _g / _q, "price_n": _n / _q, "doc": _doc, "qty": _q})
for _v in RCPT.values():
    _v.sort(key=lambda x: (x["date"], x["doc"]), reverse=True)


def checked_last_price(c, fact_cost):
    """(запись ПТУ | None, позначка, ціна сирої останньої закупівлі з ПДВ) — правило Z1 §1.4 (fact_cost — з ПДВ)."""
    lst = RCPT.get(c, [])
    if not lst:
        return None, None, None
    raw = lst[0]["price_g"]
    if not fact_cost:
        return lst[0], "остання закупівля", raw
    for i, x in enumerate(lst):
        k = x["price_g"] / fact_cost
        if Decimal(1) / PRICE_K <= k <= PRICE_K:
            return x, ("остання закупівля" if i == 0 else "попередня закупівля (остання — помилка одиниці в ПТУ)"), raw
    return None, "усі закупівлі відхиляються від собівартості факту > 5 разів (одиниця в ПТУ)", raw


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
    m = MAP.get(c)
    if m is None:
        return {"group": "поза еталоном", "kind": "не в mapping", "key": "C:" + c, "ss_line": None, "method": "не знайдено в mapping",
                "etalon_norm": None, "etalon_shared": 0, "price_ss": None, "note": "картки немає в універсумі прототипу 30 м"}
    kind = m["kind"]
    if kind == "МШП":
        return {"group": "МШП", "kind": kind, "key": "C:" + c, "ss_line": None, "method": "МШП", "etalon_norm": Z,
                "etalon_shared": 0, "price_ss": None, "note": m.get("note", "")}
    if kind in ("поза СС", "не визначено", "Виключена"):
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
                "price_ss": lpc * LINE[SEPTIC_LINE]["price"] / KIT_LP if KIT_LP else None, "note": "складова септика"}
    l = first_link(c)
    n = l["line"]
    shared = len(BY_LINE.get(n, []))
    return {"group": "еталон", "kind": kind, "key": "L:%d" % n, "ss_line": "%d %s" % (n, LINE[n]["text"]),
            "method": m["method"] + ("" if len(m["links"]) == 1 else " (+%d строк)" % (len(m["links"]) - 1)),
            "etalon_norm": LINE[n]["q"] / l["coef"] if l["coef"] else None, "etalon_shared": shared,
            "price_ss": LINE[n]["price"] * l["coef"], "coef": l["coef"], "line_unit": LINE[n]["unit"], "line": n,
            "note": l.get("note", "")}


def compute_z1():
    """Закупочная 30 м (вариант б Z1): строки по карточкам с фактом на №1/№2."""
    rows = {}
    cards = sorted(c for c in CARD if c != KIT and (any(fact_b(c, h)[0] for h in BASE_H) or any(fact_a(c, h)[0] for h in BASE_H)))
    for c in cards:
        f1q, f1g = fact_b(c, "30-1"); f2q, f2g = fact_b(c, "30-2")
        avg = (f1q + f2q) / 2
        if avg <= 0:
            continue
        f3q = fact_a(c, "30-3")[0]; s3q = stock_h(c, "30-3")[0]
        cv = (Decimal(2).sqrt() * abs(f1q - f2q) / (f1q + f2q)) if (f1q + f2q) > 0 else Z
        rows[c] = {"f1": f1q, "f2": f2q, "f1g": f1g, "f2g": f2g, "avg": avg, "f3": f3q, "s3": s3q, "cv": cv, "unstable": cv > CV_LIMIT}
    return rows


Z1R = compute_z1()
log("снимок:", RAW["stamp_full"], "движений", len(MV), "последнее", RAW.get("moves_max_period"), "строк ПТУ проекта", len(RAW["receipts"]))
log("закупочная 30 м: карточек с фактом №1/№2:", len(Z1R))

# ═════════════════════════════ ЧАСТЬ B. Потоки по системе складов IRC (класс P и H:* прототипа)
WCLS = collections.Counter()
FL = collections.defaultdict(lambda: {"purch": collections.defaultdict(lambda: [Z, Z, Z]), "in": collections.defaultdict(lambda: [Z, Z, Z]),
                                      "spys": collections.defaultdict(lambda: [Z, Z, Z]), "out": collections.defaultdict(lambda: [Z, Z, Z]),
                                      "spys_doc": collections.defaultdict(lambda: [Z, Z, Z])})
LEAK = collections.defaultdict(lambda: [0, Z])
DOCS = {"purch": collections.defaultdict(set), "in": collections.defaultdict(set), "out": collections.defaultdict(set)}
ORG_IS = 'ТОВ "ІНДЕПТ СТІЛ"'


def org_tag(r):
    return " [ІНДЕПТ СТІЛ]" if (r.get("Орг") or "") == ORG_IS else ""


def add(a, q, n, g):
    a[0] += q; a[1] += n; a[2] += g


for r in MV:
    dt_, kt_ = r["СчДт"] or "", r["СчКт"] or ""
    s = Dd(r["Сумма"]); typ = r.get("ТипРег") or "?"
    cd, ck = wcls(r["СклДт"]), wcls(r["СклКт"])
    if is_mat(dt_):
        WCLS[(r["СклДт"], cd)] += 1
    if is_mat(kt_):
        WCLS[(r["СклКт"], ck)] += 1
    din = is_mat(dt_) and S_CLS(cd)
    kin = is_mat(kt_) and S_CLS(ck)
    if din:
        c = r["КодДт"]; q = Dd(r["КолДт"] or 0); g = s * (1 + VAT.ratio(r["ПартДт"], c, r["СклДт"]))
    if kin:
        k = r["КодКт"]; qk = Dd(r["КолКт"] or 0); gk = s * (1 + VAT.ratio(r["ПартКт"], k, r["СклКт"]))
    if din and kin:
        if r["КодДт"] == r["КодКт"]:
            continue                                    # перемещение внутри проекта — не закупка и не расход
        add(FL[k]["out"]["перетворення в іншу картку: " + typ + org_tag(r)], qk, s, gk)
        add(FL[c]["in"]["перетворення з іншої картки: " + typ + org_tag(r)], q, s, g)
        DOCS["out"]["перетворення в іншу картку: " + typ + org_tag(r)].add(r["Рег"])
        DOCS["in"]["перетворення з іншої картки: " + typ + org_tag(r)].add(r["Рег"])
    elif din:
        if typ == "Надходження товарів і послуг":
            key = "ПТУ"
        elif typ == "Надходження дод. витрат":
            key = "дод. витрати (лише сума)"
        elif typ == "Авансовий звіт":
            key = "авансовий звіт"
        else:
            key = None
        if key:
            add(FL[c]["purch"][key + org_tag(r)], q, s, g); DOCS["purch"][key + org_tag(r)].add(r["Рег"])
        else:
            kk = ("переміщення з «%s»" % r["СклКт"]) if is_mat(kt_) else ("%s (Кт %s)" % (typ, kt_))
            add(FL[c]["in"][kk + org_tag(r)], q, s, g); DOCS["in"][kk + org_tag(r)].add(r["Рег"])
    elif kin:
        if not is_mat(dt_) and ck.startswith("H:"):
            add(FL[k]["spys"][ck[2:]], qk, s, gk)
            add(FL[k]["spys_doc"][(ck[2:], typ + org_tag(r))], qk, s, gk)
        elif is_mat(dt_):
            kk = "переміщення на «%s»" % r["СклДт"]
            add(FL[k]["out"][kk + org_tag(r)], qk, s, gk); DOCS["out"][kk + org_tag(r)].add(r["Рег"])
        else:
            pre = "повернення постачальнику" if typ.startswith("Повернення") else (
                ("списання зі складу проєкту не на дім: " if ck == "P" else "списання: ") + typ + " (Дт " + dt_ + ")")
            add(FL[k]["out"][pre + org_tag(r)], qk, s, gk); DOCS["out"][pre + org_tag(r)].add(r["Рег"])
    elif is_mat(dt_) or is_mat(kt_):
        LEAK[(typ, r["СклДт"] if is_mat(dt_) else r["СклКт"], dt_, kt_)][0] += 1
        LEAK[(typ, r["СклДт"] if is_mat(dt_) else r["СклКт"], dt_, kt_)][1] += s

log("\nклассы складов в движениях (матеріальні рахунки):")
for (w, cl), n in sorted(WCLS.items(), key=lambda x: (x[0][1], x[0][0])):
    log(f"   {cl:8s} {n:5d}  {w}")
log("движения вне системы складов IRC (партия IRC, склад не IRC):", sum(v[0] for v in LEAK.values()), f2(sum(v[1] for v in LEAK.values())))
for kk, v in sorted(LEAK.items(), key=lambda x: -x[1][1]):
    log("   ", kk, v[0], f2(v[1]))

# остатки Хозрасчетный (независимый запрос) по классу склада
OST = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z]))
OST_W = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z]))
for r in RAW["ostatki_now"]:
    cl = wcls(r["Склад"])
    a = OST[r["Код"]][cl]; a[0] += Dd(r["Кол"]); a[1] += Dd(r["Сумма"])
    b = OST_W[r["Код"]][r["Склад"]]; b[0] += Dd(r["Кол"]); b[1] += Dd(r["Сумма"])
log("остатки Хозрасчетный сейчас:", len(RAW["ostatki_now"]), "строк; на 01.01.2026:", len(RAW["ostatki_2026"]), "строк")
ost_cls = collections.Counter()
for c, d in OST.items():
    for cl, a in d.items():
        ost_cls[cl] += float(a[1])
log("   остаток по классам, без ПДВ:", {k: round(v, 2) for k, v in ost_cls.items()})

# закупки на подразделение IRC мимо складов проекта (ІБ_МШП, «ІС Основной склад»)
PUR_SIDE = collections.defaultdict(lambda: [Z, Z, Z, set()])
for r in RAW["receipts"]:
    if not r["Проведен"] or not Dd(r["Кол"]):
        continue
    if S_CLS(wcls(r["Склад"])):
        continue
    net = Dd(r["Сумма"]) - (Dd(r["НДС"]) if r["ВклНДС"] else Z)
    a = PUR_SIDE[r["Код"]]; a[0] += Dd(r["Кол"]); a[1] += net; a[2] += net + Dd(r["НДС"]); a[3].add(r["Номер"] + " → " + r["Склад"])
# куда ушли партии этих ПТУ (проверка Z6): Кт 20/22/28 по партии ПТУ → склад-получатель; остаток = закуплено − ушло
SIDE_DOCS = {r["Док"] for r in RAW["receipts"] if r["Проведен"] and Dd(r["Кол"]) and not S_CLS(wcls(r["Склад"]))}
SIDE_FLOW = collections.defaultdict(lambda: collections.defaultdict(lambda: Z))     # откуда -> куда -> без ПДВ
for r in MV:
    if r["ПартКт"] in SIDE_DOCS and is_mat(r["СчКт"]) and r["СклДт"] != r["СклКт"]:
        dest = r["СклДт"] if is_mat(r["СчДт"]) else "списано (Дт %s)" % r["СчДт"]
        SIDE_FLOW[r["СклКт"]][dest] += Dd(r["Сумма"])
SIDE_INTO = sum(v for src in SIDE_FLOW.values() for w, v in src.items() if S_CLS(wcls(w)))
SIDE_OTHER = {src: {w: v for w, v in d.items() if not S_CLS(wcls(w)) and wcls(w) != "MSP"} for src, d in SIDE_FLOW.items()}
SIDE_MSP = sum(v for src in SIDE_FLOW.values() for w, v in src.items() if wcls(w) == "MSP")
SIDE_OUT = sum(v for src in SIDE_FLOW.values() for v in src.values())
SIDE_LEFT = sum(v[1] for v in PUR_SIDE.values()) - SIDE_OUT

# ═════════════════════════════ ЧАСТЬ C. 25 строк эталона 30 м, не выданных на №1–№2 (решение 7)
Z1_NI = [x for x in Z1J["compare"]["etalon_lines_not_issued_on_1_2"] if x.get("line_not_issued_on_1_2") and x["et_uah_last"] > 0]
ZL = []                     # строки: {n, text, unit, norm, price_ss_g, cards[(c, coef, prio)], target, still_zero}
ZL_CARD = collections.defaultdict(lambda: {"per_house": Z, "lines": []})      # карточка -> норма на дом (ед. карточки) из нулевых строк
ZL_PSEUDO = []
for x in Z1_NI:
    n = int(x["key"][2:])
    l = LINE[n]
    still_zero = all(MAPPED[h][n]["F"] == 0 and MAPPED[h][n]["Sh"] == 0 for h in BASE_H)
    on3 = MAPPED["30-3"][n]["F"] + MAPPED["30-3"][n]["Sh"]
    linked = [(c, l2["coef"], l2["prio"]) for c, _, l2 in BY_LINE.get(n, []) if l2["coef"]]
    rec = {"n": n, "text": l["text"], "unit": l["unit"], "norm": l["q"], "price_ss_g": l["price"], "sum_ss_g": l["sum"], "linked": linked,
           "still_zero_on_1_2": still_zero, "on_3_line_units": on3, "z1_et_uah_last": x["et_uah_last"], "z1_et_uah_ss": x["et_uah_ss"],
           "z1_buy_line_units": x.get("et_buy_line_units")}
    if linked:
        # основная карточка: приоритет связи, затем та, что лежит на складе проекта, затем последняя закупка
        tgt = sorted(linked, key=lambda t: (t[2], -float(max(Z, st[t[0]]["P"][0])), -(1 if t[0] in RCPT else 0)))[0]
        rec["target"] = tgt[0]
        ZL_CARD[tgt[0]]["per_house"] += l["q"] / tgt[1]
        ZL_CARD[tgt[0]]["lines"].append(n)
    else:
        rec["target"] = None
        ZL_PSEUDO.append(rec)
    ZL.append(rec)
log("\nнулевые строки 30 м (решение 7):", len(ZL), "; с картками:", sum(1 for x in ZL if x["target"]), "; без картки:", len(ZL_PSEUDO),
    "; всё ещё не выдано на №1–№2:", sum(1 for x in ZL if x["still_zero_on_1_2"]))
for x in ZL:
    log("   ", x["n"], x["text"], f3(x["norm"]), x["unit"], "картки:", [(c, cname(c)[:30], float(k), float(st[c]["P"][0])) for c, k, p in x["linked"]],
        "→", x["target"], "на №3:", float(x["on_3_line_units"]))

# ═════════════════════════════ ЧАСТЬ D. Строки проекта по карточкам
ALLC = set(FL) | set(Z1R) | set(ZL_CARD) | set(OST)
ALLC = sorted(c for c in ALLC if c)


def cost_fact(c):
    """себестоимость единицы: всё списанное на дома + лежит на складах домов (з ПДВ / без ПДВ)."""
    q = sum(FL[c]["spys"][h][0] for h in FL[c]["spys"]) + sum(st[c][cl][0] for cl in st[c] if cl.startswith("H:"))
    n = sum(FL[c]["spys"][h][1] for h in FL[c]["spys"]) + sum(st[c][cl][1] for cl in st[c] if cl.startswith("H:"))
    g = sum(FL[c]["spys"][h][2] for h in FL[c]["spys"]) + sum(st[c][cl][2] for cl in st[c] if cl.startswith("H:"))
    if q > 0 and g > 0:
        return g / q, n / q
    sp = st[c]["P"]
    if sp[0] > 0 and sp[2] > 0:
        return sp[2] / sp[0], sp[1] / sp[0]
    return None, None


ROWS = {}
for c in ALLC:
    fl = FL[c]
    cl = classify(c)
    # потребность 15 м = факт 6 домов
    f15 = {h: (fl["spys"][h][0] + st[c]["H:" + h][0], fl["spys"][h][1] + st[c]["H:" + h][1], fl["spys"][h][2] + st[c]["H:" + h][2]) for h in H15}
    need15 = sum(v[0] for v in f15.values())
    # 30 м
    z = Z1R.get(c)
    avg = z["avg"] if z else Z
    zl = ZL_CARD[c]["per_house"] if c in ZL_CARD else Z
    per_house = avg + zl
    f30 = {h: fact_b(c, h) for h in H30}
    f1q, f2q = f30["30-1"][0], f30["30-2"][0]
    f3q = fact_a(c, "30-3")[0] + stock_h(c, "30-3")[0]
    rem = {"30-3": max(Z, per_house - f3q)}
    for h in NOT_STARTED:
        rem[h] = per_house
    need3 = f3q + rem["30-3"]
    need30 = f1q + f2q + need3 + per_house * 4
    rem_total = sum(rem.values())
    # склад: проект + дома №4–№7 (FIFO по домам; свой склад дома первым)
    sp = max(Z, st[c]["P"][0])
    own = {h: max(Z, stock_h(c, h)[0]) for h in NOT_STARTED}
    byh, pool = {}, sp
    for h in REST:
        d = rem[h]
        o = min(d, own.get(h, Z)); d2 = d - o
        if h in own and own[h] > o:
            pool += own[h] - o
        take = min(d2, pool); pool -= take
        byh[h] = {"need": d, "own": o, "from_project": take, "buy_raw": d2 - take}
    buy_raw = sum(v["buy_raw"] for v in byh.values())
    piece = is_piece(c)
    if piece:
        buy = ceil_q(buy_raw)
        cum, prev = Z, Z
        for h in REST:
            cum += byh[h]["buy_raw"]; ci = ceil_q(cum)
            byh[h]["buy"] = ci - prev; prev = ci
    else:
        buy = buy_raw
        for h in REST:
            byh[h]["buy"] = byh[h]["buy_raw"]
    surplus_q = pool                                     # склад проекта + свободные склады №4–№7 после покрытия
    # цены
    fc_g, fc_n = cost_fact(c)
    if z:
        qf = z["f1"] + z["f2"]
        fc_z1 = (z["f1g"] + z["f2g"]) / qf if qf else None       # Z1: себестоимость факта №1–№2 з ПДВ
    else:
        fc_z1 = fc_g
    lp, lp_flag, lp_raw = checked_last_price(c, fc_z1)
    if lp:
        p_last_g, p_last_n, b_last = lp["price_g"], lp["price_n"], lp_flag
    elif cl["price_ss"]:
        p_last_g, p_last_n, b_last = cl["price_ss"], cl["price_ss"] / VATK, "ціна СС (" + (lp_flag or "закупівель немає") + ")"
    elif fc_z1:
        p_last_g, p_last_n, b_last = fc_z1, (fc_n if fc_n else fc_z1 / VATK), "собівартість факту (" + (lp_flag or "немає ні закупівлі, ні ціни СС") + ")"
    else:
        p_last_g, p_last_n, b_last = None, None, "ціни немає"
    if cl["price_ss"]:
        p_ss_g, p_ss_n, b_ss = cl["price_ss"], cl["price_ss"] / VATK, "ціна СС"
    else:
        p_ss_g, p_ss_n, b_ss = p_last_g, p_last_n, "немає ціни СС → як «посл.»"
    # закуплено / прочее
    pq = sum(v[0] for v in fl["purch"].values()); pn = sum(v[1] for v in fl["purch"].values()); pg = sum(v[2] for v in fl["purch"].values())
    iq = sum(v[0] for v in fl["in"].values()); in_ = sum(v[1] for v in fl["in"].values()); ig = sum(v[2] for v in fl["in"].values())
    oq = sum(v[0] for v in fl["out"].values()); on_ = sum(v[1] for v in fl["out"].values()); og = sum(v[2] for v in fl["out"].values())
    wq = sum(v[0] for v in fl["spys"].values()); wn = sum(v[1] for v in fl["spys"].values()); wg = sum(v[2] for v in fl["spys"].values())
    stk = {cl_: st[c][cl_] for cl_ in st[c] if S_CLS(cl_)}
    sq = sum(v[0] for v in stk.values()); sn = sum(v[1] for v in stk.values()); sg = sum(v[2] for v in stk.values())
    oq_ost = sum(a[0] for cl_, a in OST[c].items() if S_CLS(cl_)); on_ost = sum(a[1] for cl_, a in OST[c].items() if S_CLS(cl_))
    ROWS[c] = {
        "c": c, "cl": cl, "piece": piece, "fl": fl, "f15": f15, "need15": need15, "f30": f30, "avg": avg, "zl": zl, "per_house": per_house,
        "f3": f3q, "need3": need3, "need30": need30, "rem": rem, "rem_total": rem_total, "sp": sp, "own": own, "byh": byh,
        "buy_raw": buy_raw, "buy": buy, "surplus_q": surplus_q, "z": z,
        "p_last_g": p_last_g, "p_last_n": p_last_n, "b_last": b_last, "p_ss_g": p_ss_g, "p_ss_n": p_ss_n, "b_ss": b_ss,
        "lp_raw": lp_raw, "fc_g": fc_g, "fc_n": fc_n,
        "pq": pq, "pn": pn, "pg": pg, "iq": iq, "in": in_, "ig": ig, "oq": oq, "on": on_, "og": og, "wq": wq, "wn": wn, "wg": wg,
        "stk": stk, "sq": sq, "sn": sn, "sg": sg, "oq_ost": oq_ost, "on_ost": on_ost,
        "diff_q": pq + iq - wq - oq - oq_ost, "diff_n": pn + in_ - wn - on_ - on_ost,
        "diff_mv_q": pq + iq - wq - oq - sq,
    }

# псевдо-строки: нулевые строки 30 м без картки
PSEUDO = {}
for x in ZL_PSEUDO:
    code = "СС30:%d" % x["n"]
    norm = x["norm"]
    rem = {h: norm for h in REST}
    byh = {}
    piece = (x["unit"] or "").strip() in PIECE
    tot_raw = norm * 5
    buy = ceil_q(tot_raw) if piece else tot_raw
    cum, prev = Z, Z
    for h in REST:
        cum += norm
        ci = ceil_q(cum) if piece else cum
        byh[h] = {"need": norm, "own": Z, "from_project": Z, "buy_raw": norm, "buy": ci - prev}; prev = ci
    PSEUDO[code] = {"x": x, "rem": rem, "byh": byh, "buy": buy, "piece": piece, "p_g": x["price_ss_g"], "p_n": x["price_ss_g"] / VATK}

# ═════════════════════════════ ЧАСТЬ E. Узлы строителя
Z3ROWS = {r["code"]: r for r in Z3J["zakupochnaya"]["rows"]}
Z3NODES = list(Z3J["zakupochnaya"]["by_node"].keys())
POS_NODE = {}
for r in Z3J["etalon"]["rows"]:
    POS_NODE.setdefault(K(r["pos"]), r["node_path"])
NODE_PATHS = sorted(set(r["node_path"] for r in Z3J["etalon"]["rows"]) | set(Z3NODES))

# решения прораба для карточек 30 м, которых нет в 15 м (и для строк 30 м без аналога в 15 м):
# ключ — код карточки или «СС30:n»; значение — узел дерева 15 м. Обоснование — в комментарии.
PRORAB = {
    "00-00005140": "Водопостачання і сантехніка",          # американка PP-R д20×1/2 Н — внутрішній водопровід (як Z3: «американк» → Водопостачання)
    "00-00005141": "Водопостачання і сантехніка",          # американка PP-R д20×1/2 В — те саме
    "00000019652": "Каналізація всередині",                 # муфта d110 — каналізаційна труба д110
    "00-00001357": "Зовнішні мережі",                        # муфта д.25-3/4 (аналог «З'єднувач компресійний 25») — зовнішній водопровід ПЕ 25 (як Z3 §10.4 п. 2)
    "00-00005019": "Зовнішні мережі",                        # оцинкована решітка 1000×1000 — вхідна група / благоустрій (як плитка тротуарна в Z3); ніде не використана
    "00-00006134": "Покрівля",                               # стрічка бутилкаучукова 150 мм — герметизація примикань покрівлі
    "00-00022793": "Каналізація всередині",                 # труба каналізаційна d50 — внутрішня каналізація (Z3: «каналізаційна, d=50» → Каналізація)
    "СС30:144": "Зовнішні мережі",                           # «Підключення ємкості до труби ПЕ» — у дереві 15 м разом з ємністю і трубою ПЕ д25
}
KW30 = [  # (подстрока названия карточки/строки, узел) — правила прораба для 30 м (в дополнение к правилам Z3)
    ("ринв", "Покрівля / Водостік 15 м"), ("водост", "Покрівля / Водостік 15 м"), ("держак", "Покрівля / Водостік 15 м"),
    ("лійк", "Покрівля / Водостік 15 м"), ("заглушка ринви", "Покрівля / Водостік 15 м"), ("з'єднувач ринви", "Покрівля / Водостік 15 м"),
    ("днище", "Зовнішні мережі / Септик"), ("кільце", "Зовнішні мережі / Септик"), ("кцд", "Зовнішні мережі / Септик"),
    ("люк", "Зовнішні мережі / Септик"), ("плита перекриття", "Зовнішні мережі / Септик"), ("кришка пп", "Зовнішні мережі / Септик"),
    ("плитка тротуарна", "Зовнішні мережі"), ("бетонна суміш", "Зовнішні мережі"), ("суміш бетонна", "Зовнішні мережі"),
    ("ємніст", "Зовнішні мережі"), ("насосна станц", "Зовнішні мережі"), ("труба пе", "Зовнішні мережі"), ("муфта пе", "Зовнішні мережі"),
    ("генератор", "Електрика / Резервне живлення (генератор)"), ("авр", "Електрика / Резервне живлення (генератор)"),
    ("автомат", "Електрика / Електрощит"), ("щрн", "Електрика / Електрощит"), ("шина", "Електрика / Електрощит"),
    ("захист від перенапруги", "Електрика / Електрощит"), ("бокс", "Електрика / Електрощит"), ("щит", "Електрика / Електрощит"),
    ("кабель", "Електрика"), ("ввгнг", "Електрика"), ("дріт пвз", "Електрика"), ("розетк", "Електрика"), ("вимикач", "Електрика"),
    ("світильник", "Електрика"), ("коробк", "Електрика"), ("гофр", "Електрика"), ("клем", "Електрика"), ("рамк", "Електрика"),
    ("лампа", "Електрика"), ("ізострічк", "Електрика"), ("сигналізац", "Оснащення"),
    ("кондиціонер", "Опалення, вентиляція і клімат"), ("обігрівач", "Опалення, вентиляція і клімат"), ("конвектор", "Опалення, вентиляція і клімат"),
    ("вентилятор", "Опалення, вентиляція і клімат"), ("тепла підлога", "Опалення, вентиляція і клімат"), ("теплої підлоги", "Опалення, вентиляція і клімат"),
    ("канал d100", "Опалення, вентиляція і клімат"), ("гратк", "Опалення, вентиляція і клімат"), ("радіатор", "Опалення, вентиляція і клімат"),
    ("бойлер", "Водопостачання і сантехніка"), ("водонагрів", "Водопостачання і сантехніка"), ("pp-r", "Водопостачання і сантехніка"),
    ("ппр", "Водопостачання і сантехніка"), ("кліпса", "Водопостачання і сантехніка"), ("кран", "Водопостачання і сантехніка"),
    ("змішувач", "Водопостачання і сантехніка"), ("американк", "Водопостачання і сантехніка"), ("умивальник", "Водопостачання і сантехніка"),
    ("унітаз", "Водопостачання і сантехніка"), ("шланг", "Водопостачання і сантехніка"), ("душ", "Водопостачання і сантехніка"),
    ("піддон", "Водопостачання і сантехніка"), ("мийк", "Водопостачання і сантехніка"), ("клапан", "Водопостачання і сантехніка"),
    ("трійник латун", "Водопостачання і сантехніка"), ("футорк", "Водопостачання і сантехніка"), ("перехідник для радіатора", "Опалення, вентиляція і клімат"),
    ("сифон", "Каналізація всередині"), ("каналіз", "Каналізація всередині"), ("д110", "Каналізація всередині"), ("д32", "Каналізація всередині"),
    ("d50", "Каналізація всередині"), ("муфта d110", "Каналізація всередині"), ("редукц", "Каналізація всередині"), ("ревізій", "Каналізація всередині"),
    ("хомут", "Каналізація всередині"), ("коліно", "Каналізація всередині"), ("відвід пвх", "Каналізація всередині"), ("трійник д", "Каналізація всередині"),
    ("перехід ексцентр", "Каналізація всередині"), ("перехід ексентр", "Каналізація всередині"),
    ("вікн", "Вікна й двері"), ("двер", "Вікна й двері"), ("підвікон", "Вікна й двері"), ("відкіс", "Вікна й двері"), ("замок", "Вікна й двері"),
    ("піна", "Вікна й двері"), ("герметик", "Вікна й двері"), ("жалюзі", "Вікна й двері"), ("штора", "Вікна й двері"), ("петл", "Вікна й двері"),
    ("ручк", "Вікна й двері"), ("поріг", "Вікна й двері"), ("доводчик", "Вікна й двері"),
    ("лінолеум", "Підлога (на 1 м²)"), ("цсп", "Підлога (на 1 м²)"), ("плінтус підлог", "Підлога (на 1 м²)"), ("профіль алюмін", "Підлога (на 1 м²)"),
    ("osb", "Стіни і перегородки / Внутрішня обшивка стін (на 1 м²)"), ("осб", "Стіни і перегородки / Внутрішня обшивка стін (на 1 м²)"),
    ("пінопласт", "Стіни і перегородки"), ("пінополістирол", "Стіни і перегородки"), ("мінват", "Стіни і перегородки"),
    ("ekoboard", "Стіни і перегородки"), ("утеплювач", "Стіни і перегородки"), ("брус", "Стіни і перегородки"), ("мембран", "Стіни і перегородки"),
    ("плівка", "Стіни і перегородки"), ("клей", "Стіни і перегородки"), ("грунт", "Стіни і перегородки"), ("ґрунт під", "Стіни і перегородки"),
    ("кутик", "Стіни і перегородки"), ("кутник 50", "Стіни і перегородки"), ("профіль т", "Стіни і перегородки"), ("профіль ат", "Стіни і перегородки"),
    ("плінтус на стільницю", "Стіни і перегородки"), ("strotex", "Стіни і перегородки"), ("вагонк", "Стеля"),
    ("профнастил", "Фасад"), ("порфлист", "Фасад"), ("пк 10", "Фасад"), ("пс-10", "Фасад"), ("лист рівний", "Фасад"), ("гладкий лист", "Фасад"),
    ("пк 8", "Фасад"), ("s-8", "Фасад"), ("заклепк", "Фасад"), ("пк 20", "Покрівля"), ("т-18", "Покрівля"), ("мембрана дахова", "Покрівля"),
    ("труба профільна", "Каркас модуля"), ("профтруб", "Каркас модуля"), ("труба ст3", "Каркас модуля"), ("труба металева", "Каркас модуля"),
    ("лист 4", "Каркас модуля"), ("лист 8", "Каркас модуля"), ("лист 14", "Каркас модуля"), ("лист пвл", "Каркас модуля"), ("кутник 32", "Каркас модуля"),
    ("профіль 40", "Каркас модуля"), ("труба 40*20", "Каркас модуля"), ("40х20", "Каркас модуля"), ("дошка", "Підлога (на 1 м²)"),
    ("гайк", "Каркас модуля / Болтові з'єднання і анкери каркаса"), ("болт", "Каркас модуля / Болтові з'єднання і анкери каркаса"),
    ("шпильк", "Каркас модуля / Болтові з'єднання і анкери каркаса"), ("шайб", "Каркас модуля / Болтові з'єднання і анкери каркаса"),
    ("гровер", "Каркас модуля / Болтові з'єднання і анкери каркаса"), ("анкер", "Каркас модуля / Болтові з'єднання і анкери каркаса"),
    ("електрод", "Каркас модуля / Зварювання і фарбування каркаса"), ("дріт 1,2", "Каркас модуля / Зварювання і фарбування каркаса"),
    ("проволка", "Каркас модуля / Зварювання і фарбування каркаса"), ("зварн", "Каркас модуля / Зварювання і фарбування каркаса"),
    ("грунт-емаль", "Каркас модуля / Зварювання і фарбування каркаса"), ("ґрунт-емаль", "Каркас модуля / Зварювання і фарбування каркаса"),
    ("фарб", "Каркас модуля / Зварювання і фарбування каркаса"), ("емаль", "Каркас модуля / Зварювання і фарбування каркаса"),
    ("розчинник", "Каркас модуля / Зварювання і фарбування каркаса"), ("corgon", "Каркас модуля / Зварювання і фарбування каркаса"),
    ("мікс балон", "Каркас модуля / Зварювання і фарбування каркаса"), ("газ", "Каркас модуля / Зварювання і фарбування каркаса"),
    ("саморіз", "Кріплення загального призначення"), ("шуруп", "Кріплення загального призначення"), ("дюбел", "Кріплення загального призначення"),
    ("вогнегасник", "Оснащення"), ("ковдр", "Оснащення"), ("килим", "Оснащення"), ("стіл", "Оснащення"), ("стільц", "Оснащення"), ("стілець", "Оснащення"),
    ("шаф", "Оснащення"), ("кушетк", "Оснащення"), ("диспенсер", "Оснащення"), ("дозатор", "Оснащення"), ("банер", "Оснащення"),
    ("табличк", "Оснащення"), ("поручен", "Оснащення"), ("дзеркал", "Оснащення"), ("смітник", "Оснащення"), ("відро", "Оснащення"),
    ("медичн", "Оснащення"), ("тримач", "Оснащення"), ("вішалк", "Оснащення"), ("тумб", "Оснащення"), ("крісл", "Оснащення"),
]
UNRESOLVED = []


def node_of(c, text=None, on=None, mshp=False):
    """узел строителя: 1) узел карточки в закупочной 15 м (Z3); 2) решение прораба; 3) МШП;
    4) позиция СС 30 м с тем же текстом, что позиция эталона 15 м; 5) правила прораба по названию; 6) «Інше»."""
    if c in Z3ROWS:
        return Z3ROWS[c]["node"], "як у закупівельній 15 м (Z3)"
    if c in PRORAB:
        return PRORAB[c], "рішення прораба Z6"
    if mshp:
        return MSHP_NODE, "МШП (як Z1/Z3)"
    if text and K(text) in POS_NODE:
        return POS_NODE[K(text)], "позиція СС 30 м = позиція еталона 15 м"
    nm = K(text or cname(c))
    for kw, nd in KW30:
        if kw in nm:
            return nd, "правило прораба Z6 за назвою («%s»)" % kw
    if on and K(on) in POS_NODE:
        return POS_NODE[K(on)], "ОН = позиція еталона 15 м"
    UNRESOLVED.append((c, text or cname(c)))
    return "Інше (прораб: уточнити)", "не визначено"


for c, r in ROWS.items():
    cl = r["cl"]
    line_text = None
    if cl.get("line") and cl["line"] in LINE:
        line_text = LINE[cl["line"]]["text"]
    mshp = cl["group"] == "МШП" or K(CARD.get(c, {}).get("on")) == "мшп"
    nd, how = node_of(c, None, None, mshp) if c in Z3ROWS or c in PRORAB or mshp else (None, None)
    if nd is None:
        # сначала по тексту позиции СС 30 м, затем по названию карточки
        if line_text and K(line_text) in POS_NODE:
            nd, how = POS_NODE[K(line_text)], "позиція СС 30 м «%s» = позиція еталона 15 м" % line_text
        elif cl["kind"] == "Складова комплекту":
            nd, how = "Зовнішні мережі / Септик", "складова септика"
        elif cl["kind"] == "Групова":
            nd, how = "Кріплення загального призначення", "група «Гвинтові кріплення»"
        else:
            nd, how = node_of(c, None, CARD.get(c, {}).get("on"))
            if how == "не визначено" and line_text:
                UNRESOLVED.pop()
                nd, how = node_of("", line_text)
    r["node"], r["node_how"] = nd, how
for code, p in PSEUDO.items():
    nd, how = (PRORAB[code], "рішення прораба Z6") if code in PRORAB else node_of("", p["x"]["text"])
    p["node"], p["node_how"] = nd, how
log("\nузел не определён:", len(UNRESOLVED))
for c, t in UNRESOLVED:
    r = ROWS.get(c)
    log("   ", c, t, "|", (r["cl"]["group"], r["cl"]["ss_line"]) if r else "", "| need30", f3(r["need30"]) if r else "", "sp", f3(r["sp"]) if r else "")

# группа для отчёта: 30 м (Z1) → иначе 15 м (Z3) → иначе по классификации
for c, r in ROWS.items():
    g30 = r["cl"]["group"] if (r["z"] or c in ZL_CARD or r["sp"] > 0) and r["cl"]["kind"] != "не в mapping" else None
    g15 = Z3ROWS[c]["group"].split(" (")[0] if c in Z3ROWS else None
    r["group"] = g30 or g15 or r["cl"]["group"]
    if c == KIT:
        r["group"] = "напівфабрикат (не рахується)"
    if r["node"] == MSHP_NODE:
        r["group"] = "МШП"

# ═════════════════════════════ ЧАСТЬ F. Деньги по строке
def money_row(r):
    c = r["c"]
    pg, pn = r["p_last_g"] or Z, r["p_last_n"] or Z
    sg, sn = r["p_ss_g"] or Z, r["p_ss_n"] or Z
    r["buy_last_g"] = r["buy"] * pg; r["buy_last_n"] = r["buy"] * pn
    r["buy_ss_g"] = r["buy"] * sg; r["buy_ss_n"] = r["buy"] * sn
    for h in REST:
        b = r["byh"][h]
        b["uah_last_g"] = b["buy"] * pg; b["uah_last_n"] = b["buy"] * pn; b["uah_ss_g"] = b["buy"] * sg; b["uah_ss_n"] = b["buy"] * sn
    # лишнее — по себестоимости склада проекта (Хозрасчетный)
    spq, spn, spg = st[c]["P"]
    unit_n = spn / spq if spq > 0 else (r["fc_n"] or pn)
    unit_g = spg / spq if spq > 0 else (r["fc_g"] or pg)
    r["surplus_n"] = r["surplus_q"] * unit_n; r["surplus_g"] = r["surplus_q"] * unit_g
    # потребность в деньгах: съеденное — по себестоимости факта, остаток 30 м — по цене последней закупки
    r["need15_n"] = sum(v[1] for v in r["f15"].values()); r["need15_g"] = sum(v[2] for v in r["f15"].values())
    cons30_n = sum(FL[c]["spys"][h][1] + st[c]["H:" + h][1] for h in ("30-1", "30-2", "30-3"))
    cons30_g = sum(FL[c]["spys"][h][2] + st[c]["H:" + h][2] for h in ("30-1", "30-2", "30-3"))
    r["need30_cons_n"], r["need30_cons_g"] = cons30_n, cons30_g
    r["need30_rem_n"], r["need30_rem_g"] = r["rem_total"] * pn, r["rem_total"] * pg
    r["need30_n"] = cons30_n + r["need30_rem_n"]; r["need30_g"] = cons30_g + r["need30_rem_g"]


for r in ROWS.values():
    money_row(r)
for code, p in PSEUDO.items():
    p["buy_last_g"] = p["buy"] * p["p_g"]; p["buy_last_n"] = p["buy"] * p["p_n"]
    for h in REST:
        b = p["byh"][h]
        b["uah_last_g"] = b["buy"] * p["p_g"]; b["uah_last_n"] = b["buy"] * p["p_n"]; b["uah_ss_g"] = b["uah_last_g"]; b["uah_ss_n"] = b["uah_last_n"]

# ═════════════════════════════ ЧАСТЬ G. Контроли
# G1. закупочная 30 м по Z1-правилам на свежем снимке (без нулевых строк) — против Z1 §11.4 «живая база» 1 044 351,93 / 1 158 389,78
z1_live_last = z1_live_ss = Z
for c, r in ROWS.items():
    z = r["z"]
    if not z:
        continue
    # повтор Z1: потребность без нулевых строк
    avg = z["avg"]
    rem = max(Z, avg - r["f3"]) + avg * 4
    sp = r["sp"] + sum(r["own"].values())
    br = max(Z, rem - sp)
    b = ceil_q(br) if r["piece"] else br
    z1_live_last += b * (r["p_last_g"] or Z); z1_live_ss += b * (r["p_ss_g"] or Z)
log("\nКОНТРОЛЬ 1: закупочная 30 м по правилам Z1 на свежем снимке:", f2(z1_live_last), "/", f2(z1_live_ss),
    "(Z1 §11.4 живая база: 1 044 351,93 / 1 158 389,78)")
# G2. 15 м: факт против Z3
f15_spys_n = sum(sum(r["fl"]["spys"][h][1] for h in H15) for r in ROWS.values())
f15_lie_n = sum(sum(st[r["c"]]["H:" + h][1] for h in H15) for r in ROWS.values())
log("КОНТРОЛЬ 2: факт 15 м без ПДВ: списано", f2(f15_spys_n), "+ лежит", f2(f15_lie_n), "=", f2(f15_spys_n + f15_lie_n),
    "(Z3: 2 944 362,31 + 30 768,30 = 2 975 130,61)")
# G3. замкнутость (по движениям — тождество; по независимым остаткам — сверка)
bad_mv = [r for r in ROWS.values() if abs(r["diff_mv_q"]) > Decimal("0.0005")]
bad = [r for r in ROWS.values() if abs(r["diff_q"]) > Decimal("0.0005") or abs(r["diff_n"]) > Decimal("0.05")]
log("КОНТРОЛЬ 3: замкнутость: по движениям расхождений", len(bad_mv), "; против Хозрасчетный.Остатки расхождений", len(bad))
for r in bad[:30]:
    log("   ", r["c"], cname(r["c"])[:40], "q", f3(r["diff_q"]), "грн", f2(r["diff_n"]), "ост", f3(r["oq_ost"]), "движ", f3(r["sq"]))
tot_ost_n = sum(r["on_ost"] for r in ROWS.values()); tot_mv_n = sum(r["sn"] for r in ROWS.values())
log("   остаток по движениям", f2(tot_mv_n), "; Хозрасчетный.Остатки", f2(tot_ost_n))

# ═════════════════════════════ ЧАСТЬ H. Итоги
def T(sel=lambda r: True):
    rs = [r for r in ROWS.values() if sel(r)]
    t = collections.OrderedDict()
    t["cards"] = len(rs)
    for k in ("need15_n", "need15_g", "need30_n", "need30_g", "need30_cons_n", "need30_rem_n", "need30_rem_g", "pn", "pg", "in", "ig", "on", "og",
              "wn", "wg", "sn", "sg", "buy_last_g", "buy_last_n", "buy_ss_g", "buy_ss_n", "surplus_n", "surplus_g", "on_ost"):
        t[k] = sum(r[k] for r in rs)
    t["spys_15_n"] = sum(sum(r["fl"]["spys"][h][1] for h in H15) for r in rs)
    t["spys_30_n"] = sum(sum(r["fl"]["spys"][h][1] for h in H30) for r in rs)
    t["spys_15_g"] = sum(sum(r["fl"]["spys"][h][2] for h in H15) for r in rs)
    t["spys_30_g"] = sum(sum(r["fl"]["spys"][h][2] for h in H30) for r in rs)
    t["lie_houses_n"] = sum(sum(v[1] for cl_, v in r["stk"].items() if cl_.startswith("H:")) for r in rs)
    t["lie_houses_g"] = sum(sum(v[2] for cl_, v in r["stk"].items() if cl_.startswith("H:")) for r in rs)
    t["lie_project_n"] = sum(r["stk"].get("P", [Z, Z, Z])[1] for r in rs)
    t["lie_project_g"] = sum(r["stk"].get("P", [Z, Z, Z])[2] for r in rs)
    t["cards_to_buy"] = sum(1 for r in rs if r["buy"] > EPS)
    t["cards_surplus"] = sum(1 for r in rs if r["surplus_q"] > EPS)
    t["by_house_last_g"] = {h: sum(r["byh"][h]["uah_last_g"] for r in rs) for h in REST}
    t["by_house_last_n"] = {h: sum(r["byh"][h]["uah_last_n"] for r in rs) for h in REST}
    t["by_house_ss_g"] = {h: sum(r["byh"][h]["uah_ss_g"] for r in rs) for h in REST}
    return t


TOT = T()
PS_T = {"buy_last_g": sum(p["buy_last_g"] for p in PSEUDO.values()), "buy_last_n": sum(p["buy_last_n"] for p in PSEUDO.values()),
        "by_house_last_g": {h: sum(p["byh"][h]["uah_last_g"] for p in PSEUDO.values()) for h in REST},
        "by_house_last_n": {h: sum(p["byh"][h]["uah_last_n"] for p in PSEUDO.values()) for h in REST}}
BUY30_LAST_G = TOT["buy_last_g"] + PS_T["buy_last_g"]
BUY30_SS_G = TOT["buy_ss_g"] + PS_T["buy_last_g"]
BUY30_LAST_N = TOT["buy_last_n"] + PS_T["buy_last_n"]
BUY30_SS_N = TOT["buy_ss_n"] + PS_T["buy_last_n"]
BYH_LAST_G = {h: TOT["by_house_last_g"][h] + PS_T["by_house_last_g"][h] for h in REST}
BYH_LAST_N = {h: TOT["by_house_last_n"][h] + PS_T["by_house_last_n"][h] for h in REST}
BYH_SS_G = {h: TOT["by_house_ss_g"][h] + PS_T["by_house_last_g"][h] for h in REST}
BYM_LAST_G = collections.OrderedDict((m, sum(v for h, v in BYH_LAST_G.items() if MONTH[h] == m)) for m in ("2026-10", "2026-11"))
BYM_LAST_N = collections.OrderedDict((m, sum(v for h, v in BYH_LAST_N.items() if MONTH[h] == m)) for m in ("2026-10", "2026-11"))
BYM_SS_G = collections.OrderedDict((m, sum(v for h, v in BYH_SS_G.items() if MONTH[h] == m)) for m in ("2026-10", "2026-11"))
log("\n=== ИТОГИ ПРОЕКТА")
for k, v in TOT.items():
    log("  ", k, {h: f2(x) for h, x in v.items()} if isinstance(v, dict) else (f2(v) if isinstance(v, Decimal) else v))
log("   нулевые строки без картки:", len(PSEUDO), f2(PS_T["buy_last_g"]), "з ПДВ")
log("   ЕЩЁ ДОКУПИТЬ 30 м, з ПДВ: посл.", f2(BUY30_LAST_G), "СС", f2(BUY30_SS_G), "; без ПДВ: посл.", f2(BUY30_LAST_N), "СС", f2(BUY30_SS_N))
log("   по домам посл. з ПДВ:", {HN[h]: f2(v) for h, v in BYH_LAST_G.items()}, "по месяцам:", {m: f2(v) for m, v in BYM_LAST_G.items()})

# разложение разницы «ещё докупить» с Z1-гибридом
Z1_HYB_LAST, Z1_HYB_SS = Decimal("1314870.57"), Decimal("1427219.78")
Z1_ZAK_LAST, Z1_ZAK_SS = Decimal("1001174.97"), Decimal("1098389.78")
V3 = Decimal("1313551.98")          # G_h v3: норма СС × цена СС, з ПДВ
V3_MSHP = Decimal("12002.98")       # МШП_h v3 — отдельный член бюджета (в Z6 МШП входит в «ещё докупить»)
# доля нулевых строк в моём расчёте: разница «с нулевыми строками − без них» по карточкам + псевдо-строки
zl_effect_last = zl_effect_ss = Z
for c, r in ROWS.items():
    if c not in ZL_CARD:
        continue
    z = r["z"]
    avg = z["avg"] if z else Z
    rem0 = max(Z, avg - r["f3"]) + avg * 4
    sp = r["sp"] + sum(r["own"].values())
    br0 = max(Z, rem0 - sp)
    b0 = ceil_q(br0) if r["piece"] else br0
    zl_effect_last += (r["buy"] - b0) * (r["p_last_g"] or Z); zl_effect_ss += (r["buy"] - b0) * (r["p_ss_g"] or Z)
zl_total_last = zl_effect_last + PS_T["buy_last_g"]
zl_total_ss = zl_effect_ss + PS_T["buy_last_g"]
log("   нулевые строки в моём расчёте: посл.", f2(zl_total_last), "СС", f2(zl_total_ss), "(Z1: 313 695,60 / 328 830,00)")
log("   закупочная 30 м (без нулевых строк) на свежем снимке:", f2(BUY30_LAST_G - zl_total_last), "/", f2(BUY30_SS_G - zl_total_ss))

# ═════════════════════════════ ЧАСТЬ I. Узлы
NODE_ORDER = ["Каркас модуля", "Каркас модуля / Болтові з'єднання і анкери каркаса", "Каркас модуля / Зварювання і фарбування каркаса",
              "Підлога (на 1 м²)", "Стіни і перегородки", "Стіни і перегородки / Внутрішня обшивка стін (на 1 м²)", "Стеля", "Покрівля",
              "Покрівля / Водостік 15 м", "Фасад", "Вікна й двері", "Електрика", "Електрика / Електрощит", "Електрика / Резервне живлення (генератор)",
              "Опалення, вентиляція і клімат", "Водопостачання і сантехніка", "Каналізація всередині", "Зовнішні мережі", "Зовнішні мережі / Септик",
              "Оснащення", "Кріплення загального призначення", MSHP_NODE, "Інше (прораб: уточнити)"]
BYN = collections.OrderedDict()
for nd in NODE_ORDER:
    BYN[nd] = collections.defaultdict(lambda: Z)
for r in ROWS.values():
    if r["c"] == KIT:
        nd = "Зовнішні мережі / Септик"
    else:
        nd = r["node"]
    a = BYN.setdefault(nd, collections.defaultdict(lambda: Z))
    a["cards"] += 1
    for k in ("need15_n", "need30_n", "need30_g", "pn", "pg", "wn", "sn", "buy_last_g", "buy_last_n", "buy_ss_g", "surplus_n", "in", "on"):
        a[k] += r[k]
    a["spys_15_n"] += sum(r["fl"]["spys"][h][1] for h in H15)
    a["spys_30_n"] += sum(r["fl"]["spys"][h][1] for h in H30)
    a["lie_houses_n"] += sum(v[1] for cl_, v in r["stk"].items() if cl_.startswith("H:"))
    a["lie_project_n"] += r["stk"].get("P", [Z, Z, Z])[1]
for code, p in PSEUDO.items():
    a = BYN.setdefault(p["node"], collections.defaultdict(lambda: Z))
    a["lines_no_card"] += 1
    a["buy_last_g"] += p["buy_last_g"]; a["buy_last_n"] += p["buy_last_n"]; a["buy_ss_g"] += p["buy_last_g"]
    a["need30_n"] += p["buy"] * p["p_n"]; a["need30_g"] += p["buy"] * p["p_g"]
log("\n=== ПО УЗЛАМ (без ПДВ; докупить — з ПДВ)")
for nd, a in BYN.items():
    if not a:
        continue
    log(f"  {nd[:60]:60s} карт {int(a['cards']):3d} п15 {f2(a['need15_n']):>13,.2f} п30 {f2(a['need30_n']):>13,.2f} закуп {f2(a['pn']):>13,.2f} "
        f"спис {f2(a['wn']):>13,.2f} лежит {f2(a['sn']):>12,.2f} докуп(зПДВ) {f2(a['buy_last_g']):>12,.2f} лишн {f2(a['surplus_n']):>11,.2f}")

# ═════════════════════════════ ЧАСТЬ J. Топы и прочие приходы/расходы
TOP_BUY = sorted(list(ROWS.values()), key=lambda r: -r["buy_last_g"])
TOP_BUY_ALL = sorted([("card", r["c"], r["buy_last_g"]) for r in ROWS.values() if r["buy"] > EPS] +
                     [("pseudo", code, p["buy_last_g"]) for code, p in PSEUDO.items()], key=lambda x: -x[2])[:20]
TOP_SUR = sorted([r for r in ROWS.values() if r["surplus_q"] > EPS], key=lambda r: -r["surplus_n"])[:20]
OTH_IN = collections.defaultdict(lambda: [Z, Z, Z, 0]); OTH_OUT = collections.defaultdict(lambda: [Z, Z, Z, 0]); PUR = collections.defaultdict(lambda: [Z, Z, Z, 0])
for r in ROWS.values():
    for k, v in r["fl"]["in"].items():
        a = OTH_IN[k]; a[0] += v[0]; a[1] += v[1]; a[2] += v[2]; a[3] += 1
    for k, v in r["fl"]["out"].items():
        a = OTH_OUT[k]; a[0] += v[0]; a[1] += v[1]; a[2] += v[2]; a[3] += 1
    for k, v in r["fl"]["purch"].items():
        a = PUR[k]; a[0] += v[0]; a[1] += v[1]; a[2] += v[2]; a[3] += 1
log("\nзакуплено по видам:")
for k, v in sorted(PUR.items(), key=lambda x: -x[1][1]):
    log(f"   {k:60s} карт {v[3]:3d} без ПДВ {f2(v[1]):>14,.2f} з ПДВ {f2(v[2]):>14,.2f} док {len(DOCS['purch'][k])}")
log("прочие приходы:")
for k, v in sorted(OTH_IN.items(), key=lambda x: -x[1][1]):
    log(f"   {k[:90]:90s} карт {v[3]:3d} без ПДВ {f2(v[1]):>12,.2f} док {len(DOCS['in'][k])}")
log("прочие расходы:")
for k, v in sorted(OTH_OUT.items(), key=lambda x: -x[1][1]):
    log(f"   {k[:90]:90s} карт {v[3]:3d} без ПДВ {f2(v[1]):>12,.2f} док {len(DOCS['out'][k])}")
SPYS_DOC = collections.defaultdict(lambda: [Z, Z, Z])
for r in ROWS.values():
    for (h, typ), v in r["fl"]["spys_doc"].items():
        a = SPYS_DOC[(h[:2] + " м", typ)]; a[0] += v[0]; a[1] += v[1]; a[2] += v[2]
log("списано на дома по видам документов:")
for k, v in sorted(SPYS_DOC.items()):
    log(f"   {k}  без ПДВ {f2(v[1]):>14,.2f}")
log("закуплено на подразделение IRC мимо складов проекта (ПТУ):", len(PUR_SIDE), "карт.", f2(sum(v[1] for v in PUR_SIDE.values())), "без ПДВ",
    f2(sum(v[2] for v in PUR_SIDE.values())), "з ПДВ")
log("   из них: в систему IRC", f2(SIDE_INTO), "; на другие объекты", {s: {w: f2(v) for w, v in d.items()} for s, d in SIDE_OTHER.items() if d},
    "; на МШП-склады вне системы", f2(SIDE_MSP), "; осталось на складе закупки", f2(SIDE_LEFT))

# топ-20 печать
log("\nТОП-20 ДОКУПИТЬ (з ПДВ, посл.):")
for kind, code, v in TOP_BUY_ALL:
    if kind == "card":
        r = ROWS[code]
        log(f"   {f2(v):>12,.2f}  {cname(code)[:50]:50s} {unit(code):5s} потр.ост {f3(r['rem_total']):>9} склад {f3(r['sp']):>8} купить {f3(r['buy']):>8}  {r['node']}")
    else:
        p = PSEUDO[code]
        log(f"   {f2(v):>12,.2f}  [без картки] {p['x']['text'][:38]:38s} {p['x']['unit']:5s} купить {f3(p['buy']):>8}  {p['node']}")
log("\nТОП-20 ЛИШНЕЕ (без ПДВ, по себестоимости склада):")
for r in TOP_SUR:
    log(f"   {f2(r['surplus_n']):>12,.2f}  {cname(r['c'])[:50]:50s} {unit(r['c']):5s} склад {f3(r['sp']):>8} потр.ост {f3(r['rem_total']):>8} лишн {f3(r['surplus_q']):>8}  {r['node']}")

# ═════════════════════════════ ЧАСТЬ K. Дополнительные проверки и пояснения
# K1. снимки Z1 (10:52) и Z3 (20:03) против свежего: движение за движением
Z1_RAW = os.environ.get("Z1_RAW", r"C:/Users/SUPPOR~1/AppData/Local/Temp/claude/C--Configuration-downloads-BASERP25--claude-worktrees-budget-calc-oct-dec-2026-20c998/74acb5f4-5c54-4380-b2d5-fb369593cb1b/scratchpad/svyaz/raw.json")
Z3_RAW = os.environ.get("Z3_RAW", r"C:/Users/SUPPOR~1/AppData/Local/Temp/claude/z15_tmp/z3/raw15.json")


def mkey(m):
    return (m["Рег"], m["СчДт"], m["СчКт"], m["КодДт"], m["КодКт"], m["СклДт"], m["СклКт"], round(m["Сумма"], 2),
            round(m["КолДт"] or 0, 4), round(m["КолКт"] or 0, 4), m["ПартДт"], m["ПартКт"])


SNAP_DIFF = {}
for nm, pth in (("Z1", Z1_RAW), ("Z3", Z3_RAW)):
    if not os.path.exists(pth):
        SNAP_DIFF[nm] = {"note": "файл снимка не найден: " + pth}
        continue
    old = json.load(open(pth, encoding="utf-8"))
    A = collections.Counter(mkey(m) for m in MV); B_ = collections.Counter(mkey(m) for m in old["moves"])
    plus, minus = A - B_, B_ - A
    docs = collections.defaultdict(lambda: {"added_lines": 0, "removed_lines": 0, "sum_delta": Z})
    for k, v in plus.items():
        docs[k[0]]["added_lines"] += v; docs[k[0]]["sum_delta"] += Dd(k[7]) * v
    for k, v in minus.items():
        docs[k[0]]["removed_lines"] += v; docs[k[0]]["sum_delta"] -= Dd(k[7]) * v
    rk = lambda r: (r["Док"], r["Код"], round(r["Кол"], 4), round(r["Сумма"], 2), r["Проведен"])
    ra = collections.Counter(rk(r) for r in RAW["receipts"]); rb = collections.Counter(rk(r) for r in old["receipts"])
    SNAP_DIFF[nm] = {"stamp": old.get("stamp"), "moves_old": len(old["moves"]), "moves_new": len(MV), "lines_added": sum(plus.values()),
                     "lines_removed": sum(minus.values()), "receipts_added": sum((ra - rb).values()), "receipts_removed": sum((rb - ra).values()),
                     "docs": {d: {"added_lines": v["added_lines"], "removed_lines": v["removed_lines"], "sum_delta_net": f2(v["sum_delta"])} for d, v in sorted(docs.items())}}
    log("\nснимок", nm, SNAP_DIFF[nm]["stamp"], "против свежего:", SNAP_DIFF[nm]["lines_added"], "+ /", SNAP_DIFF[nm]["lines_removed"], "−; ПТУ +",
        SNAP_DIFF[nm]["receipts_added"], "−", SNAP_DIFF[nm]["receipts_removed"])
    for d, v in SNAP_DIFF[nm]["docs"].items():
        log("   ", d, v)

# K2. закупочная 30 м по карточкам: свежий снимок против JSON Z1 (что дало разницу)
Z1ROWS = {r["code"]: r for r in Z1J["rows"]}
Z1_DELTA = []
for c in sorted(set(Z1ROWS) | {c for c, r in ROWS.items() if r["z"]}):
    r = ROWS.get(c)
    if r and r["z"]:
        rem0 = max(Z, r["z"]["avg"] - r["f3"]) + r["z"]["avg"] * 4
        br0 = max(Z, rem0 - r["sp"] - sum(r["own"].values()))
        b0 = ceil_q(br0) if r["piece"] else br0
        mine_l, mine_s = b0 * (r["p_last_g"] or Z), b0 * (r["p_ss_g"] or Z)
        mine_q, mine_avg, mine_sp = b0, r["z"]["avg"], r["sp"]
    else:
        mine_l = mine_s = mine_q = mine_avg = mine_sp = Z
    z = Z1ROWS.get(c)
    zl_, zs_ = Dd(z["buy_uah_last"]) if z else Z, Dd(z["buy_uah_ss"]) if z else Z
    if abs(mine_l - zl_) > Decimal("0.005") or abs(mine_s - zs_) > Decimal("0.005") or (z is None) != (r is None or not r["z"]):
        Z1_DELTA.append({"code": c, "card": cname(c), "z1_avg": z["avg"] if z else None, "avg_now": f3(mine_avg),
                         "z1_stock_project": z["stock_project"] if z else None, "stock_project_now": f3(mine_sp),
                         "z1_buy": z["buy_qty"] if z else None, "buy_now": f3(mine_q), "z1_uah_last": z["buy_uah_last"] if z else None,
                         "uah_last_now": f2(mine_l), "delta_last": f2(mine_l - zl_), "delta_ss": f2(mine_s - zs_), "in_z1": z is not None})
log("\nзакупочная 30 м: карточки с отличием от Z1:", len(Z1_DELTA), "; Δ посл.", f2(sum(Dd(x["delta_last"]) for x in Z1_DELTA)),
    "; Δ СС", f2(sum(Dd(x["delta_ss"]) for x in Z1_DELTA)))
for x in Z1_DELTA:
    log("   ", x)

# K3. тождество проекта по количеству: закуплено + прочие приходы − прочие расходы + докупить = потребность 15 + потребность 30 + лишнее + округление
ID4 = []
for c, r in ROWS.items():
    lhs = r["pq"] + r["iq"] - r["oq"] + r["buy"]
    rhs = r["need15"] + r["need30"] + r["surplus_q"] + (r["buy"] - r["buy_raw"])
    if abs(lhs - rhs) > Decimal("0.0005"):
        ID4.append({"code": c, "card": cname(c), "lhs": f3(lhs), "rhs": f3(rhs), "diff": f3(lhs - rhs), "stock_P": f3(st[c]["P"][0]),
                    "asm_30": f3(sum(ASM[(c, h)][0] for h in H30)) if c in KIT_CARDS else 0.0})
log("\nКОНТРОЛЬ 4: тождество проекта по количеству — карточек с невязкой:", len(ID4))
for x in ID4:
    log("   ", x)

# K4. по домам: списано и лежит (без ПДВ)
BY_HOUSE = collections.OrderedDict()
for h in H15 + H30:
    BY_HOUSE[HN[h]] = {"written_off_uah_net": f2(sum(r["fl"]["spys"][h][1] for r in ROWS.values())),
                       "written_off_uah_vat": f2(sum(r["fl"]["spys"][h][2] for r in ROWS.values())),
                       "lies_uah_net": f2(sum(st[r["c"]]["H:" + h][1] for r in ROWS.values())),
                       "cards_written_off": sum(1 for r in ROWS.values() if r["fl"]["spys"][h][0] or r["fl"]["spys"][h][1]),
                       "cards_lying": sum(1 for r in ROWS.values() if st[r["c"]]["H:" + h][0] or st[r["c"]]["H:" + h][1])}
log("\nпо домам:", json.dumps(BY_HOUSE, ensure_ascii=False))

# K5. нулевые строки 30 м: карточка-кандидат по опыту 15 м (позиция эталона 15 м с тем же текстом)
POS_CARD15 = {}
for r_ in Z3J["etalon"]["rows"]:
    if r_.get("card_code"):
        POS_CARD15.setdefault(K(r_["pos"]), (r_["card_code"], r_["card"], r_["q_card_house"]))
for x in ZL:
    cand = POS_CARD15.get(K(x["text"]))
    x["cand15"] = None
    if cand:
        cc = cand[0]; rr = ROWS.get(cc)
        lp = RCPT.get(cc, [None])[0]
        x["cand15"] = {"code": cc, "card": cand[1], "etalon_15_per_house": cand[2], "need_15_fact": f3(rr["need15"]) if rr else 0.0,
                       "fact_30_1_2": f3(rr["f30"]["30-1"][0] + rr["f30"]["30-2"][0]) if rr else 0.0,
                       "stock_project": f3(st[cc]["P"][0]), "price_last_vat": f2(lp["price_g"]) if lp else None}
log("\nнулевые строки 30 м — карточка по опыту 15 м:")
for x in ZL:
    log("   ", x["n"], x["text"], "→", x["cand15"])

# K6. 15 м по цене последней закупки (сравнение с Z3: средний дом 490 895,25 × 6 = 2 945 371,50 без ПДВ)
need15_last_n = sum(r["need15"] * (r["p_last_n"] or Z) for r in ROWS.values())
log("\n15 м по цене последней закупки (без ПДВ, правило Z1):", f2(need15_last_n), "; Z3:", 2945371.50)

# K7. лишнее и возможные замены (решение 3: взаимозаменяемость пока только «Група» — здесь только подсветка)
S8 = ROWS.get("00-00016723"); PC10 = next((p for code, p in PSEUDO.items() if p["x"]["n"] == 65), None)

# K8. нулевые строки: выдавались ли на 15 м (опыт 6 домов) или не выдавались ни на одном доме IRC
ZL_SPLIT = {"issued_on_15m": {"lines": [], "uah_last_vat": Z, "uah_ss_vat": Z}, "never_issued": {"lines": [], "uah_last_vat": Z, "uah_ss_vat": Z}}
for x in ZL:
    if x["target"]:
        r = ROWS[x["target"]]
        used15 = r["need15"] > 0
        ul, us = r["buy_last_g"], r["buy_ss_g"]
    else:
        p = next(p for p in PSEUDO.values() if p["x"]["n"] == x["n"])
        used15 = bool(x["cand15"] and x["cand15"]["need_15_fact"] > 0)
        ul = us = p["buy_last_g"]
    k = "issued_on_15m" if used15 else "never_issued"
    ZL_SPLIT[k]["lines"].append("%d %s" % (x["n"], x["text"])); ZL_SPLIT[k]["uah_last_vat"] += ul; ZL_SPLIT[k]["uah_ss_vat"] += us
log("\nнулевые строки: выдавались на 15 м —", len(ZL_SPLIT["issued_on_15m"]["lines"]), f2(ZL_SPLIT["issued_on_15m"]["uah_last_vat"]),
    "; не выдавались ни на одном доме —", len(ZL_SPLIT["never_issued"]["lines"]), f2(ZL_SPLIT["never_issued"]["uah_last_vat"]))
S8_SCEN = None
if S8 and PC10:
    s8_unit_n = S8["surplus_n"] / S8["surplus_q"] if S8["surplus_q"] else Z
    take = min(S8["surplus_q"], PC10["buy"])
    S8_SCEN = {"s8_surplus_m2": f3(S8["surplus_q"]), "s8_surplus_uah_net": f2(S8["surplus_n"]), "pk10_need_m2": f3(PC10["buy"]),
               "pk10_to_buy_uah_vat": f2(PC10["buy_last_g"]), "covered_m2": f3(take),
               "to_buy_reduction_uah_vat": f2(PC10["p_g"] * take), "surplus_reduction_uah_net": f2(s8_unit_n * take)}
    log("сценарий «стены ПК 10 из S-8 склада проекта»:", S8_SCEN)


def fm(x, d=2):
    v = float(x)
    s_ = f"{abs(v):,.{d}f}".replace(",", " ").replace(".", ",")
    return ("−" if v < 0 else "") + s_


# ═════════════════════════════ JSON
def q3(x):
    return f3(x) if x is not None else None


def g2(x):
    return f2(x) if x is not None else None


rows_json = []
for c in sorted(ROWS, key=lambda c: (NODE_ORDER.index(ROWS[c]["node"]) if ROWS[c]["node"] in NODE_ORDER else 99, cname(c))):
    r = ROWS[c]; fl = r["fl"]
    rows_json.append({
        "card": cname(c), "code": c, "unit": unit(c), "piece": r["piece"], "node": r["node"], "node_how": r["node_how"], "group": r["group"],
        "ss_line_30": r["cl"]["ss_line"], "kind_30": r["cl"]["kind"],
        "need_15": q3(r["need15"]), "need_15_by_house": {HN[h]: q3(v[0]) for h, v in r["f15"].items() if v[0]},
        "need_15_uah_net": g2(r["need15_n"]),
        "need_30": q3(r["need30"]),
        "need_30_parts": {"fact_1": q3(r["f30"]["30-1"][0]), "fact_2": q3(r["f30"]["30-2"][0]), "need_3": q3(r["need3"]), "fact_3_so_far": q3(r["f3"]),
                          "avg_1_2": q3(r["avg"]), "etalon_lines_not_issued_per_house": q3(r["zl"]), "per_house_3_7": q3(r["per_house"]),
                          "remaining_3_7": q3(r["rem_total"]), "unstable_cv_gt_0_3": bool(r["z"]["unstable"]) if r["z"] else None},
        "need_30_uah_net": g2(r["need30_n"]), "need_30_uah_vat": g2(r["need30_g"]),
        "written_off": {"total": q3(r["wq"]), "uah_net": g2(r["wn"]), "uah_vat": g2(r["wg"]),
                        "by_house": {HN[h]: {"qty": q3(v[0]), "uah_net": g2(v[1])} for h, v in sorted(fl["spys"].items()) if v[0] or v[1]}},
        "stock": {"project": q3(r["stk"].get("P", [Z])[0]), "project_uah_net": g2(r["stk"].get("P", [Z, Z])[1]),
                  "houses": {HN[cl_[2:]]: q3(v[0]) for cl_, v in sorted(r["stk"].items()) if cl_.startswith("H:") and (v[0] or v[1])},
                  "houses_uah_net": g2(sum(v[1] for cl_, v in r["stk"].items() if cl_.startswith("H:"))),
                  "by_warehouse_ostatki": {w: q3(a[0]) for w, a in OST_W[c].items() if a[0] or a[1]}},
        "purchased": {"qty": q3(r["pq"]), "uah": g2(r["pn"]), "uah_net": g2(r["pn"]), "uah_vat": g2(r["pg"]),
                      "by_kind": {k: {"qty": q3(v[0]), "uah_net": g2(v[1])} for k, v in fl["purch"].items()}},
        "purchased_irc_side_warehouse": ({"qty": q3(PUR_SIDE[c][0]), "uah_net": g2(PUR_SIDE[c][1]), "docs": sorted(PUR_SIDE[c][3])} if c in PUR_SIDE else None),
        "other_in": {"qty": q3(r["iq"]), "uah_net": g2(r["in"]), "by_kind": {k: {"qty": q3(v[0]), "uah_net": g2(v[1])} for k, v in fl["in"].items()}},
        "other_out": {"qty": q3(r["oq"]), "uah_net": g2(r["on"]), "by_kind": {k: {"qty": q3(v[0]), "uah_net": g2(v[1])} for k, v in fl["out"].items()}},
        "stock_ostatki_now": {"qty": q3(r["oq_ost"]), "uah_net": g2(r["on_ost"])},
        "balance_diff": {"qty": q3(r["diff_q"]), "uah_net": g2(r["diff_n"]),
                         "formula": "закуплено + прочие приходы − списано − прочие расходы − остаток Хозрасчетный.Остатки (склады IRC)"},
        "to_buy_30": {"qty": q3(r["buy"]), "qty_raw": q3(r["buy_raw"]), "uah_last": g2(r["buy_last_g"]), "uah_ss": g2(r["buy_ss_g"]),
                      "uah_last_net": g2(r["buy_last_n"]), "uah_ss_net": g2(r["buy_ss_n"]),
                      "price_last_vat": g2(r["p_last_g"]), "price_last_net": g2(r["p_last_n"]), "price_basis_last": r["b_last"],
                      "price_ss_vat": g2(r["p_ss_g"]), "price_basis_ss": r["b_ss"],
                      "by_house": {HN[h]: {"need": q3(b["need"]), "own_house_stock": q3(b["own"]), "from_project": q3(b["from_project"]),
                                           "buy": q3(b["buy"]), "uah_last": g2(b["uah_last_g"]), "uah_ss": g2(b["uah_ss_g"]), "month": MONTH[h]}
                                   for h, b in r["byh"].items()}},
        "surplus": {"qty": q3(r["surplus_q"]), "uah": g2(r["surplus_n"]), "uah_net": g2(r["surplus_n"]), "uah_vat": g2(r["surplus_g"])},
    })
for code, p in PSEUDO.items():
    x = p["x"]
    rows_json.append({
        "card": "— без картки: рядок СС 30 м %d «%s»" % (x["n"], x["text"]), "code": code, "unit": x["unit"], "piece": p["piece"], "node": p["node"],
        "node_how": p["node_how"], "group": "еталон (рядок не видано на №1–№2, рішення 7)", "ss_line_30": "%d %s" % (x["n"], x["text"]), "kind_30": "без картки",
        "need_15": 0.0, "need_30": q3(p["buy"]),
        "need_30_parts": {"fact_1": 0.0, "fact_2": 0.0, "need_3": q3(x["norm"]), "etalon_lines_not_issued_per_house": q3(x["norm"]), "remaining_3_7": q3(x["norm"] * 5)},
        "need_30_uah_net": g2(p["buy"] * p["p_n"]), "need_30_uah_vat": g2(p["buy"] * p["p_g"]),
        "written_off": {"total": 0.0, "uah_net": 0.0, "by_house": {}}, "stock": {"project": 0.0, "houses": {}},
        "purchased": {"qty": 0.0, "uah": 0.0, "uah_net": 0.0, "uah_vat": 0.0}, "other_in": {"qty": 0.0, "uah_net": 0.0}, "other_out": {"qty": 0.0, "uah_net": 0.0},
        "balance_diff": {"qty": 0.0, "uah_net": 0.0},
        "to_buy_30": {"qty": q3(p["buy"]), "uah_last": g2(p["buy_last_g"]), "uah_ss": g2(p["buy_last_g"]), "uah_last_net": g2(p["buy_last_n"]),
                      "uah_ss_net": g2(p["buy_last_n"]), "price_last_vat": g2(p["p_g"]), "price_basis_last": "ціна СС (картки немає)",
                      "price_ss_vat": g2(p["p_g"]), "price_basis_ss": "ціна СС",
                      "by_house": {HN[h]: {"need": q3(b["need"]), "buy": q3(b["buy"]), "uah_last": g2(b["uah_last_g"]), "uah_ss": g2(b["uah_ss_g"]),
                                           "month": MONTH[h]} for h, b in p["byh"].items()}},
        "surplus": {"qty": 0.0, "uah": 0.0, "uah_net": 0.0, "uah_vat": 0.0},
        "note": "жовтий рядок: за еталоном до рішення бухгалтера/технолога (рішення 7)",
    })

by_node_json = collections.OrderedDict()
for nd, a in BYN.items():
    if not a:
        continue
    by_node_json[nd] = {"cards": int(a["cards"]), "lines_no_card": int(a["lines_no_card"]),
                        "need_15_uah_net": f2(a["need15_n"]), "need_30_uah_net": f2(a["need30_n"]), "need_30_uah_vat": f2(a["need30_g"]),
                        "purchased_uah_net": f2(a["pn"]), "purchased_uah_vat": f2(a["pg"]),
                        "written_off_15_uah_net": f2(a["spys_15_n"]), "written_off_30_uah_net": f2(a["spys_30_n"]), "written_off_uah_net": f2(a["wn"]),
                        "lies_houses_uah_net": f2(a["lie_houses_n"]), "lies_project_uah_net": f2(a["lie_project_n"]), "lies_uah_net": f2(a["sn"]),
                        "other_in_uah_net": f2(a["in"]), "other_out_uah_net": f2(a["on"]),
                        "to_buy_30_uah_last_vat": f2(a["buy_last_g"]), "to_buy_30_uah_last_net": f2(a["buy_last_n"]), "to_buy_30_uah_ss_vat": f2(a["buy_ss_g"]),
                        "surplus_uah_net": f2(a["surplus_n"])}

grp_tot = collections.OrderedDict()
for g in sorted({r["group"] for r in ROWS.values()}):
    t = T(lambda r, g=g: r["group"] == g)
    grp_tot[g] = {"cards": t["cards"], "to_buy_uah_last_vat": f2(t["buy_last_g"]), "to_buy_uah_ss_vat": f2(t["buy_ss_g"]),
                  "to_buy_uah_last_net": f2(t["buy_last_n"]), "surplus_uah_net": f2(t["surplus_n"]), "need_15_uah_net": f2(t["need15_n"]),
                  "need_30_uah_net": f2(t["need30_n"]), "purchased_uah_net": f2(t["pn"])}

top_buy_json = []
for kind, code, v in TOP_BUY_ALL:
    if kind == "card":
        r = ROWS[code]
        top_buy_json.append({"code": code, "card": cname(code), "unit": unit(code), "node": r["node"], "group": r["group"],
                             "remaining_need_3_7": f3(r["rem_total"]), "stock_project": f3(r["sp"]), "to_buy": f3(r["buy"]),
                             "uah_last_vat": f2(r["buy_last_g"]), "uah_ss_vat": f2(r["buy_ss_g"]), "uah_last_net": f2(r["buy_last_n"]),
                             "price_basis": r["b_last"], "unstable": bool(r["z"]["unstable"]) if r["z"] else None,
                             "from_etalon_lines_not_issued": c in ZL_CARD if False else (code in ZL_CARD)})
    else:
        p = PSEUDO[code]
        top_buy_json.append({"code": code, "card": "— без картки: «%s»" % p["x"]["text"], "unit": p["x"]["unit"], "node": p["node"],
                             "group": "еталон (не видано)", "to_buy": f3(p["buy"]), "uah_last_vat": f2(p["buy_last_g"]), "uah_ss_vat": f2(p["buy_last_g"]),
                             "uah_last_net": f2(p["buy_last_n"]), "price_basis": "ціна СС", "from_etalon_lines_not_issued": True})
top_sur_json = [{"code": r["c"], "card": cname(r["c"]), "unit": unit(r["c"]), "node": r["node"], "group": r["group"], "stock_project": f3(r["sp"]),
                 "remaining_need_3_7": f3(r["rem_total"]), "surplus": f3(r["surplus_q"]), "uah_net": f2(r["surplus_n"]), "uah_vat": f2(r["surplus_g"]),
                 "need_15": f3(r["need15"]), "need_30": f3(r["need30"])} for r in TOP_SUR]

RES = {
    "snapshot": {"datetime": RAW["stamp_full"], "base": 'Srvr="localhost";Ref="bas_industrialbud" (лише читання, ВЫБРАТЬ)',
                 "moves": len(MV), "moves_max_period": RAW.get("moves_max_period"), "receipts_project_lines": len(RAW["receipts"]),
                 "ostatki_now_lines": len(RAW["ostatki_now"]), "ostatki_on_2026_01_01_lines": len(RAW["ostatki_2026"]),
                 "cards": len(ROWS), "pseudo_rows_no_card": len(PSEUDO), "collector": "Z6_scripts/z6_collect.py", "calc": "Z6_scripts/z6_calc.py"},
}
RES["snapshot"]["diff_vs_z1"] = SNAP_DIFF.get("Z1")
RES["snapshot"]["diff_vs_z3"] = SNAP_DIFF.get("Z3")
RES["method"] = {
    "system": "Система складов проекта = класс P прототипа («ІБ _МД IRC 2026», «ІС _МД IRC 2026», «МД IRС 2026», «Цех МД IRS 2026», «МД IRS 2026») "
              "+ склады домов (имя «… IRC/IRС 2026 15|30 м №N …», обе организации). Счета 20/22/28 (материальные). Движения 2026 года: "
              "до 2026 на складах IRC двигался только «Цех МД IRC» (другой проект «МД IRS»), остатков на 01.01.2026 нет.",
    "purchased": "Закуплено = Дт 20/22/28 на склад системы с нематериального счёта документами «Надходження товарів і послуг» (ПТУ), "
                 "«Надходження дод. витрат» (только сумма, без количества), «Авансовий звіт». Количество и сумма без ПДВ — Хозрасчетный; з ПДВ — ПДВ партии (ПТУ).",
    "other_in_out": "Прочие приходы/расходы = перемещения с/на склады вне системы (ІБ_МШП, «ІБ Загальний склад Виробництво», «ІС Основной склад» "
                    "ТОВ «ІНДЕПТ СТІЛ», другие объекты), списания со склада проекта не на дом, возвраты поставщику, перетворення карточек "
                    "(комплектация между карточками). Перемещения внутри системы (склад проекта ↔ склады домов) — не закупка и не расход.",
    "written_off": "Списано на дом = Кт 20/22/28 со склада дома → Дт нематериальный счёт (КН «Амбулаторія модульна» на 26, ПМА на 231, "
                   "списания) — правило cons прототипа. ПМА ТОВ «ІНДЕПТ СТІЛ» со складов «ІС … №N Виробничий» — тоже на дом.",
    "stock": "Лежит = сальдо карточек по движениям на складах системы; независимо сверено с Хозрасчетный.Остатки на текущий момент.",
    "need_15": "Потребность 15 м = факт 6 домов (списано + лежит на складах дома) = закупочная 15 м × 6. Тип закрыт, РК не пересчитываются.",
    "need_30": "Потребность 30 м = факт №1 + №2 (списано + лежит; складові септика — как Z1) + №3: max(среднее, факт №3) + среднее × 4 (№4–№7) "
               "+ 25 строк эталона 30 м, не выданных на №1–№2, по норме СС × 5 домов (решение 7; 1 строка с карточкой — стулья, 24 — без карточки).",
    "to_buy_30": "Ещё докупить = max(0, остаток потребности №3–№7 − склад проекта − склады №4–№7); штучные (шт, рул, пак, пар, компл, уп, лист, балон) — "
                 "вверх один раз на итог карточки; по домам — накопленным итогом, FIFO склада проекта №3 → №7. Месяцы: №3–№6 — 2026-10, №7 — 2026-11. "
                 "Неттинг по карточке (аналоги не подменяют друг друга, решение 3).",
    "surplus": "Останется лишним = склад проекта + склады №4–№7 после покрытия потребности 30 м (до округления закупки), по себестоимости склада.",
    "price": "Цена посл. — последняя закупка карточки, средневзвешенная по документу ПТУ, с проверкой ×5 к себестоимости факта (правило Z1 §1.4); "
             "закупок нет — цена СС через связь (з ПДВ; без ПДВ = ÷1,2); нет ни того, ни другого — себестоимость факта. "
             "Вторая оценка — цена СС. Строки без карточки — цена СС. Остатки и списания — себестоимость Хозрасчетный (без ПДВ) и з ПДВ партии.",
    "nodes": "Узел строителя: карточка есть в закупочной 15 м — её узел (Z3: 21 узел + группа МШП); иначе — позиция СС 30 м с тем же текстом, "
             "что позиция эталона 15 м; складові септика — «Септик»; группа «Гвинтові кріплення» — «Кріплення»; дальше правила прораба Z6 по названию "
             "и 8 ручных решений (PRORAB в z6_calc.py).",
    "basis": "Суммы отчётов — без ПДВ (решение 8); «ещё докупить» дополнительно з ПДВ для сравнения с Z1 и бюджетом v3 (они з ПДВ).",
}
RES["rows"] = rows_json
RES["by_node"] = by_node_json
RES["totals"] = {
    "basis": "без ПДВ — Хозрасчетный (собівартість); з ПДВ — ПДВ партії (ПТУ) / ціна з ПДВ; ціна посл. — остання закупівля (середньозважена по ПТУ)",
    "need_total_uah_net": f2(TOT["need15_n"] + TOT["need30_n"] + sum(p["buy"] * p["p_n"] for p in PSEUDO.values())),
    "need_15_uah_net": f2(TOT["need15_n"]), "need_15_uah_vat": f2(TOT["need15_g"]),
    "need_30_uah_net": f2(TOT["need30_n"] + sum(p["buy"] * p["p_n"] for p in PSEUDO.values())),
    "need_30_uah_vat": f2(TOT["need30_g"] + sum(p["buy"] * p["p_g"] for p in PSEUDO.values())),
    "need_30_consumed_1_3_uah_net": f2(TOT["need30_cons_n"]),
    "need_30_remaining_3_7_uah_net_last": f2(TOT["need30_rem_n"] + sum(p["buy"] * p["p_n"] for p in PSEUDO.values())),
    "purchased_uah_net": f2(TOT["pn"]), "purchased_uah_vat": f2(TOT["pg"]),
    "purchased_by_kind": {k: {"cards": v[3], "uah_net": f2(v[1]), "uah_vat": f2(v[2]), "docs": len(DOCS["purch"][k])} for k, v in PUR.items()},
    "purchased_irc_subdivision_other_warehouse": {"cards": len(PUR_SIDE), "uah_net": f2(sum(v[1] for v in PUR_SIDE.values())),
                                                 "uah_vat": f2(sum(v[2] for v in PUR_SIDE.values())),
                                                 "flow_uah_net": {"into_irc_system": f2(SIDE_INTO),
                                                                  "to_other_objects": {s: {w: f2(v) for w, v in d.items()} for s, d in SIDE_OTHER.items() if d},
                                                                  "to_other_mshp_warehouses": f2(SIDE_MSP),
                                                                  "still_on_purchase_warehouse": f2(SIDE_LEFT)}},
    "other_in_uah_net": f2(TOT["in"]), "other_in_by_kind": {k: {"cards": v[3], "uah_net": f2(v[1]), "docs": len(DOCS["in"][k])} for k, v in OTH_IN.items()},
    "other_out_uah_net": f2(TOT["on"]), "other_out_by_kind": {k: {"cards": v[3], "uah_net": f2(v[1]), "docs": len(DOCS["out"][k])} for k, v in OTH_OUT.items()},
    "written_off_uah_net": f2(TOT["wn"]), "written_off_15_uah_net": f2(TOT["spys_15_n"]), "written_off_30_uah_net": f2(TOT["spys_30_n"]),
    "written_off_uah_vat": f2(TOT["wg"]),
    "written_off_by_doc_type": {"%s | %s" % k: f2(v[1]) for k, v in sorted(SPYS_DOC.items())},
    "lies_uah_net": f2(TOT["sn"]), "lies_houses_uah_net": f2(TOT["lie_houses_n"]), "lies_project_uah_net": f2(TOT["lie_project_n"]),
    "lies_uah_vat": f2(TOT["sg"]), "lies_ostatki_now_uah_net": f2(TOT["on_ost"]),
    "to_buy_30_uah_last_vat": f2(BUY30_LAST_G), "to_buy_30_uah_ss_vat": f2(BUY30_SS_G),
    "to_buy_30_uah_last_net": f2(BUY30_LAST_N), "to_buy_30_uah_ss_net": f2(BUY30_SS_N),
    "to_buy_30_cards": TOT["cards_to_buy"], "to_buy_30_lines_no_card": len(PSEUDO),
    "to_buy_30_by_house_last_vat": {HN[h]: f2(v) for h, v in BYH_LAST_G.items()},
    "to_buy_30_by_house_ss_vat": {HN[h]: f2(v) for h, v in BYH_SS_G.items()},
    "to_buy_30_by_house_last_net": {HN[h]: f2(v) for h, v in BYH_LAST_N.items()},
    "to_buy_30_by_month_last_vat": {m: f2(v) for m, v in BYM_LAST_G.items()},
    "to_buy_30_by_month_ss_vat": {m: f2(v) for m, v in BYM_SS_G.items()},
    "to_buy_30_by_month_last_net": {m: f2(v) for m, v in BYM_LAST_N.items()},
    "to_buy_30_etalon_lines_not_issued_last_vat": f2(zl_total_last), "to_buy_30_etalon_lines_not_issued_ss_vat": f2(zl_total_ss),
    "surplus_uah_net": f2(TOT["surplus_n"]), "surplus_uah_vat": f2(TOT["surplus_g"]), "surplus_cards": TOT["cards_surplus"],
    "by_group": grp_tot,
    "identity_check": {
        "money_net": "закуплено + прочие приходы − прочие расходы = списано + лежит",
        "lhs": f2(TOT["pn"] + TOT["in"] - TOT["on"]), "rhs": f2(TOT["wn"] + TOT["sn"]),
        "balance_diff_cards": len(bad), "balance_diff_qty_abs_sum": f3(sum(abs(r["diff_q"]) for r in ROWS.values())),
        "balance_diff_uah_net_abs_sum": f2(sum(abs(r["diff_n"]) for r in ROWS.values()))},
    "top20_to_buy": top_buy_json, "top20_surplus": top_sur_json,
    "control_z1_rules_on_fresh_snapshot": {"uah_last_vat": f2(z1_live_last), "uah_ss_vat": f2(z1_live_ss),
                                           "z1_live_base_11_4": {"uah_last_vat": 1044351.93, "uah_ss_vat": 1158389.78}},
    "control_15m_fact_net": {"written_off": f2(f15_spys_n), "lies": f2(f15_lie_n), "total": f2(f15_spys_n + f15_lie_n), "z3": 2975130.61},
}
RES["compare"] = {
    "to_buy_30_vs_z1_hybrid": {"mine_last_vat": f2(BUY30_LAST_G), "mine_ss_vat": f2(BUY30_SS_G), "z1_hybrid_last_vat": 1314870.57,
                               "z1_hybrid_ss_vat": 1427219.78, "delta_last": f2(BUY30_LAST_G - Z1_HYB_LAST), "delta_ss": f2(BUY30_SS_G - Z1_HYB_SS),
                               "decomposition_last": {
                                   "zakupochnaya_snapshot_change": f2(z1_live_last - Z1_ZAK_LAST),
                                   "etalon_lines_not_issued_card_netting": f2(zl_total_last - Decimal("313695.60")),
                                   "other_rounding": f2((BUY30_LAST_G - Z1_HYB_LAST) - (z1_live_last - Z1_ZAK_LAST) - (zl_total_last - Decimal("313695.60")))},
                               "decomposition_ss": {
                                   "zakupochnaya_snapshot_change": f2(z1_live_ss - Z1_ZAK_SS),
                                   "etalon_lines_not_issued_card_netting": f2(zl_total_ss - Decimal("328830.00")),
                                   "other_rounding": f2((BUY30_SS_G - Z1_HYB_SS) - (z1_live_ss - Z1_ZAK_SS) - (zl_total_ss - Decimal("328830.00")))}},
    "to_buy_30_vs_budget_v3": {"v3": 1313551.98,
                               "v3_basis": "G_h бюджета v3: Σ max(0, норма СС 000000009 − списано − склад дома − склад проекта FIFO) × цена СС, з ПДВ, "
                                           "30 м №3–№7. МШП в v3 — отдельный член МШП_h 12 002,98; позиции вне сметы ВС_h = 0 "
                                           "(«Как считался сценарий … мой расчет»/12_КОРРЕКТИРОВКА_МАТЕРИАЛОВ_IRC_v3.md)",
                               "delta_last": f2(BUY30_LAST_G - V3), "delta_last_note": "разные базы цены: посл. закупка Z6 против цены СС v3",
                               "delta_ss": f2(BUY30_SS_G - V3), "delta_ss_note": "одна база цены (СС); Z6 дополнительно включает МШП и «поза еталоном»",
                               "v3_g_plus_mshp": f2(V3 + V3_MSHP),
                               "delta_last_vs_g_plus_mshp": f2(BUY30_LAST_G - V3 - V3_MSHP), "delta_ss_vs_g_plus_mshp": f2(BUY30_SS_G - V3 - V3_MSHP),
                               "budget_replacement_note": "1 358 047,53 Z6 включает МШП (группа «МШП»): в бюджете заменять им G_h + МШП_h, а не только G_h, "
                                                          "иначе МШП посчитается дважды"},
    "z1_by_month_last_vat": {"2026-10": 799569.91, "2026-11": 515300.67},
    "mine_by_month_last_vat": {m: f2(v) for m, v in BYM_LAST_G.items()},
}
RES["etalon_lines_not_issued"] = [{"line": x["n"], "text": x["text"], "unit": x["unit"], "norm_per_house": f3(x["norm"]), "price_ss_vat": f2(x["price_ss_g"]),
                                   "card": x["target"], "card_name": cname(x["target"]) if x["target"] else None,
                                   "still_not_issued_on_1_2": x["still_zero_on_1_2"], "on_3_line_units": f3(x["on_3_line_units"]),
                                   "z1_uah_last": x["z1_et_uah_last"], "z1_uah_ss": x["z1_et_uah_ss"], "z1_buy_line_units": x["z1_buy_line_units"]} for x in ZL]
RES["totals"]["by_house"] = BY_HOUSE
RES["totals"]["control_identity_qty"] = {"rule": "закуплено + прочие приходы − прочие расходы + докупить = потребность 15 + потребность 30 + лишнее + округление",
                                         "cards_with_residual": len(ID4), "rows": ID4}
RES["totals"]["etalon_lines_not_issued_split"] = {k: {"lines": v["lines"], "uah_last_vat": f2(v["uah_last_vat"]), "uah_ss_vat": f2(v["uah_ss_vat"])}
                                                  for k, v in ZL_SPLIT.items()}
RES["totals"]["scenario_pk10_from_s8_surplus"] = S8_SCEN
RES["compare"]["z1_card_deltas_zakupochnaya"] = Z1_DELTA
RES["compare"]["need_15_at_last_price_net"] = {"mine_z1_price_rule": f2(need15_last_n), "z3_avg_house_x6": 2945371.50,
                                               "delta": f2(need15_last_n - Decimal("2945371.50")),
                                               "why": "Z3: цена — последняя строка ПТУ проекта по дате, затем ПТУ 2025–2026 любого подразделения, затем цена нормы; "
                                                      "Z6/Z1: средневзвешенная по последнему ПТУ с проверкой ×5, затем цена СС, затем себестоимость факта"}
for x, j in zip(ZL, RES["etalon_lines_not_issued"]):
    j["candidate_card_by_15m_experience"] = x["cand15"]
OUT_S = sum(Dd(v["uah_net"]) for v in RES["totals"]["other_out_by_kind"].values())
IS_IN = OTH_IN.get("переміщення з «ІС Основной склад» [ІНДЕПТ СТІЛ]", [Z, Z])[1]
# решение 3: «Група» крепежа взаимозаменяема. Расчёт неттингует по карточке; сколько дал бы неттинг склада внутри группы (проверка Z6)
GRP = [r for r in ROWS.values() if r["cl"]["kind"] == "Групова"]
GRP_SUR = sum(r["surplus_q"] for r in GRP)
GRP_BUY = [(r["p_last_g"] or Z, r["buy"]) for r in GRP if r["buy"] > EPS]


def _grp_take(order):
    left, s = GRP_SUR, Z
    for p, b in order:
        t = min(left, b); s += t * p; left -= t
    return s


GROUP_NET = {"cards": len(GRP), "surplus_pcs": f3(GRP_SUR), "to_buy_pcs": f3(sum(b for _, b in GRP_BUY)),
             "nettable_pcs": f3(min(GRP_SUR, sum(b for _, b in GRP_BUY))),
             "reduction_uah_last_vat_min": f2(_grp_take(sorted(GRP_BUY))), "reduction_uah_last_vat_max": f2(_grp_take(sorted(GRP_BUY, reverse=True))),
             "surplus_cards": [{"card": cname(r["c"]), "surplus_pcs": f3(r["surplus_q"]), "uah_net": f2(r["surplus_n"])} for r in GRP if r["surplus_q"] > EPS],
             "note": "решение 3 разрешает взаимозаменяемость внутри «Групи» крепежа; Z6 неттингует склад по карточке, поэтому «ещё докупить» "
                     "может быть завышено на min…max (зависит от того, какую карточку группы закрывает лишнее)"}
RES["totals"]["group_fasteners_netting_check"] = GROUP_NET
RES["caveats"] = [
    "Снимок %s (BuhBud, только ВЫБРАТЬ): %d движений, последнее %s; совпадает со снимком Z3 (20:03:41) движение в движение; против Z1 (10:52) "
    "+7/−7 строк: распроведены служебные КН септика ІБ00-000174/175, проведены переміщення ІБ00-002378/002379 (обогреватели 7 шт и дошка на №1/№2) "
    "и ІБ00-002376/002377 (на ІБ_МШП, вне системы), перепроведено ІБ00-002347 (продукция 26, без влияния)." % (RAW["stamp_full"], len(MV), RAW.get("moves_max_period")),
    "Вся разница «ещё докупить 30 м» с гибридом Z1 (+%s по посл. закупке, +%s по СС) — обогреватели TC395: после ІБ00-002378 они есть на обоих домах, "
    "среднее 3,5 → 7 (= норма эталона), склад проекта 14 → 7, докупка 4 → 28 шт. 25 строк эталона дали ровно столько же, сколько в Z1 (313 695,60 / 328 830,00)."
    % (fm(BUY30_LAST_G - Z1_HYB_LAST), fm(BUY30_SS_G - Z1_HYB_SS)),
    "Склад проекта неттуется целиком под 30 м: тип 15 м закрыт, других потребителей у склада нет. Лишнее %s без ПДВ (%d карточек) — запас или неликвид "
    "на момент снимка; часть его — аналоги карточек, которые докупаются (решение 3: аналоги не неттингуются)." % (fm(TOT["surplus_n"]), TOT["cards_surplus"]),
    "Строки эталона 30 м без выдачи на №1–№2 (решение 7) взяты по норме СС на 5 домов, жёлтые: %d строк, %s з ПДВ. Из них %d строк выдавались на 15 м "
    "(%s) и %d не выдавались ни на одном из 8 домов IRC (%s) — главные кандидаты в «не постачається»."
    % (len(ZL), fm(zl_total_last), len(ZL_SPLIT["issued_on_15m"]["lines"]), fm(ZL_SPLIT["issued_on_15m"]["uah_last_vat"]),
       len(ZL_SPLIT["never_issued"]["lines"]), fm(ZL_SPLIT["never_issued"]["uah_last_vat"])),
    ("ПК 10 стены (500 м², %s з ПДВ) и лишний профнастил S-8 на складе проекта (%s м², %s без ПДВ): если стены 30 м закрываются S-8, "
     "докупка −%s з ПДВ и лишнее −%s без ПДВ. Решает технолог." % (fm(S8_SCEN["pk10_to_buy_uah_vat"]), fm(S8_SCEN["s8_surplus_m2"], 3),
                                                                   fm(S8_SCEN["s8_surplus_uah_net"]), fm(S8_SCEN["to_buy_reduction_uah_vat"]),
                                                                   fm(S8_SCEN["surplus_reduction_uah_net"]))) if S8_SCEN else "",
    "Закупки проекта = ПТУ на склады системы (все на «ІБ _МД IRC 2026»), плюс дод. витрати и авансовые отчёты. Ещё %s без ПДВ закуплено ПТУ с подразделением IRC "
    "на ІБ_МШП и «ІС Основной склад» (исправлено проверкой Z6): в систему IRC переміщенням пришло %s (учтено в «прочих приходах»); на другие объекты ушло %s "
    "(%s); на «ІБ_МШП виробничий» — %s; на складе закупки осталось %s."
    % (fm(sum(v[1] for v in PUR_SIDE.values())), fm(SIDE_INTO), fm(sum(v for d in SIDE_OTHER.values() for v in d.values())),
       "; ".join("с «%s» %s: %s" % (s, fm(sum(d.values())), ", ".join("%s %s" % (w, fm(v)) for w, v in sorted(d.items(), key=lambda x: -x[1])))
                 for s, d in SIDE_OTHER.items() if d), fm(SIDE_MSP), fm(SIDE_LEFT)),
    "Бюджет v3 (G_h 1 313 551,98) посчитан по норме СС и цене СС; МШП в нём — отдельный член МШП_h 12 002,98. Поэтому «Z6 − v3 = +%s» сравнивает разные "
    "базы цены; по одной базе (цена СС) разница +%s. Z6 включает МШП, поэтому в бюджете он заменяет G_h + МШП_h = %s: +%s по посл. закупке, "
    "+%s по цене СС (исправлено проверкой Z6)." % (fm(BUY30_LAST_G - V3), fm(BUY30_SS_G - V3), fm(V3 + V3_MSHP), fm(BUY30_LAST_G - V3 - V3_MSHP),
                                                  fm(BUY30_SS_G - V3 - V3_MSHP)),
    "Крепёж «Гвинтові кріплення» (решение 3: «Група» взаимозаменяема) неттингуется по карточке: лишние %s шт на складе не закрывают докупку других "
    "карточек группы. Неттинг по группе снизил бы «ещё докупить» на %s…%s з ПДВ (замечание проверки Z6, итог не меняли)."
    % (fm(GRP_SUR, 0), fm(GROUP_NET["reduction_uah_last_vat_min"]), fm(GROUP_NET["reduction_uah_last_vat_max"])),
    "ТОВ «ІНДЕПТ СТІЛ»: своих ПТУ на склады IRC нет; МШП приходит переміщенням с «ІС Основной склад» (%s без ПДВ) и передаётся в эксплуатацию (ПМА) "
    "на дома: 15 м — 14 998,77, 30 м — 8 812,35. Чей это бюджет — вопрос к пользователю (как в Z1 §11.5)." % fm(IS_IN),
    "Прочие расходы %s без ПДВ: возврат МШП на ІБ_МШП, списание топлива авансового отчёта со склада проекта (Списання товарів на 231, 9 543,92), "
    "переміщення на «МД ООН Админ №1». Возвратов поставщику нет." % fm(OUT_S),
    "Выборка 30 м — два дома (№1, №2): нестабильные карточки (CV > 0,3) и разовые ошибки (вторая дверь на №2, двойное днище на №1) переносятся в "
    "потребность №3–№7 так же, как в Z1 (§6.1–6.2).",
    "15 м по цене последней закупки даёт %s без ПДВ против 2 945 371,50 у Z3 (средний дом × 6): разные правила цены, количества те же." % fm(need15_last_n),
    "Комплект септика (напівфабрикат 00-00023184): после распроведения ІБ00-000174/175 на №1 он списан (1 компл.) и одновременно лежит −1; "
    "складові лежат на складе №1. В потребность он не входит (счёт по складовим), в сверке замкнутости даёт 0.",
]
RES["caveats"] = [x for x in RES["caveats"] if x]
RES["unresolved_nodes"] = [{"code": c, "name": t} for c, t in UNRESOLVED]
# независимая проверка (proverka/q_t4_proverka.py → proverka/q_t4_out.json): переносится в JSON, если есть
_VER = os.path.join(HERE, "proverka", "q_t4_out.json")
if os.path.exists(_VER):
    RES["verification"] = json.load(open(_VER, encoding="utf-8"))
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
with open(OUT_TXT, "w", encoding="utf-8") as f:
    f.write("\n".join(LOG))
log("\nok ->", OUT_JSON)
