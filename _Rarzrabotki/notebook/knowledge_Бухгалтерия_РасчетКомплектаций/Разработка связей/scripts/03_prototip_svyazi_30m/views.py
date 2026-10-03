# -*- coding: utf-8 -*-
"""Два вида по домам 30 м №1–№7: «как сейчас» (Шахматка, по ОН) и «с соответствием» (по строке СС через связи).
Все суммы — с НДС: норма/цена СС — Комплектующие.Сумма (с НДС); факт/склад — нетто BuhBud × (1 + НДС партии) (как бюджет v3)."""
from mapping import *

HOUSES = H30
TODO = ["30-3", "30-4", "30-5", "30-6", "30-7"]       # незавершённые (FIFO склада проекта по номеру, как бюджет v3)
DONE = ["30-1", "30-2"]

# сборка комплекта на складе дома: компонент -> (кол, брутто), ушедшие в карточку-напівфабрикат, по дому
ASM = collections.defaultdict(lambda: [Z, Z])
for r in MV:
    if is_mat(r["СчДт"]) and is_mat(r["СчКт"]) and r["КодДт"] == KIT and r["КодКт"] != KIT:
        c_ = wcls(r["СклКт"])
        if c_.startswith("H:"):
            s = Dd(r["Сумма"])
            ASM[(r["КодКт"], c_[2:])][0] += Dd(r["КолКт"] or 0)
            ASM[(r["КодКт"], c_[2:])][1] += s * (1 + VAT.ratio(r["ПартКт"], r["КодКт"], r["СклКт"]))


def qg(c, cls):
    """(кол, брутто) карточки: cls = 'F:30-1' (списано на дом), 'S:30-1' (лежит на складе дома), 'P' (склад проекта)."""
    if cls == "P":
        a = st[c]["P"]
        return a[0], a[2]
    t, h = cls.split(":")
    a = cons[c][h] if t == "F" else st[c]["H:" + h]
    return a[0], a[2]


def last_price(c):
    x = LAST.get(c)
    return x["price_g"] if x else None


# ═════════════════════════════ ВИД «КАК СЕЙЧАС» (Шахматка: связка «дом + Загальна назва»)
ON_NORM = collections.defaultdict(lambda: [Z, Z])      # ОН -> [норма кол, норма грн] на дом
for l in SS:
    ON_NORM[l["k"]][0] += l["q"]; ON_NORM[l["k"]][1] += l["sum"]
ON_NAME = {}
for l in SS:
    ON_NAME.setdefault(l["k"], l["on"])
for c in U:
    ON_NAME.setdefault(K(CARD[c]["on"]) or ("ном|" + K(CARD[c]["name"])), CARD[c]["on"] or CARD[c]["name"])
ON_CARDS = collections.defaultdict(list)
for c in U:
    ON_CARDS[K(CARD[c]["on"]) or ("ном|" + K(CARD[c]["name"]))].append(c)
ON_KEYS = sorted(set(ON_NORM) | set(ON_CARDS))


def greedy(rows, N):
    """Жадное распределение нормы как РассчитатьАнализ: строки (кол, брутто) по убыванию кол.
    Возврат: (в норме кол, понад кол, понад грн)."""
    rem, inn, over_q, over_g = N, Z, Z, Z
    for row in sorted(rows, key=lambda x: -x[0]):
        q, g = row[0], row[1]
        if q <= 0:
            continue
        take = min(q, max(Z, rem)); rem -= take; inn += take
        over_q += q - take
        og = g - (g / q * take)
        over_g += og
        if len(row) > 2 and og:
            BY_CARD_OVER[row[2]] = BY_CARD_OVER.get(row[2], Z) + og
    return inn, over_q, over_g


BY_CARD_OVER = {}


def now_house(h):
    """По ОН: N, F (списано), Sh (лежит), понад (кол/грн), економія (кол/грн, только №1–№2), треба (D, только №3–№7)."""
    out = {}
    for k in ON_KEYS:
        N, Ns = ON_NORM.get(k, [Z, Z])
        rows, F, Fg, Sh, Shg = [], Z, Z, Z, Z
        for c in ON_CARDS.get(k, []):
            fq, fg = qg(c, "F:" + h); sq, sg = qg(c, "S:" + h)
            F += fq; Fg += fg; Sh += sq; Shg += sg
            if fq + sq > 0:
                rows.append((fq + sq, fg + sg, ("now", h, c)))
        inn, oq, og = greedy(rows, N)
        rec = {"N": N, "Ns": Ns, "F": F, "Fg": Fg, "Sh": Sh, "Shg": Shg, "over_q": oq, "over_g": og,
               "price_ss": (Ns / N) if N else Z, "econ_q": Z, "econ_g": Z, "D": Z}
        gap = max(Z, N - F - Sh)
        if h in DONE:
            rec["econ_q"] = gap; rec["econ_g"] = gap * rec["price_ss"]
        else:
            rec["D"] = gap
        if N or F or Sh:
            out[k] = rec
    return out


NOW = {h: now_house(h) for h in HOUSES}
NOW_SP = {k: (sum(qg(c, "P")[0] for c in ON_CARDS.get(k, [])), sum(qg(c, "P")[1] for c in ON_CARDS.get(k, []))) for k in ON_KEYS}


def now_last_price(k):
    """Цена последней закупки «по ОН»: карточка ОН с самой поздней закупкой (как видит снабженец по Шахматке)."""
    best = None
    for c in ON_CARDS.get(k, []):
        x = LAST.get(c)
        if x and (best is None or x["date"] > best["date"]):
            best = x
    return best["price_g"] if best else None


def now_buy():
    """Шахматка: Перемістити = мін(ΣD, Sp), Докупити = макс(0, ΣD − Sp); по домам — FIFO №3→№7."""
    res = {}
    for k in ON_KEYS:
        N, Ns = ON_NORM.get(k, [Z, Z])
        if N <= 0:
            continue
        sp = max(Z, NOW_SP[k][0])
        rem = sp
        byh = {}
        for h in TODO:
            d = NOW[h][k]["D"] if k in NOW[h] else N
            take = min(d, rem); rem -= take
            byh[h] = {"D": d, "move": take, "buy": d - take}
        sd = sum(v["D"] for v in byh.values())
        lp = now_last_price(k)
        res[k] = {"N": N, "price_ss": Ns / N, "price_last": lp, "Sp": sp, "sumD": sd, "move": min(sd, sp), "buy": max(Z, sd - sp), "byh": byh}
    return res


NOW_BUY = now_buy()

# ═════════════════════════════ ВИД «С СООТВЕТСТВИЕМ» (строка СС через связи)
LINES_N = [l["n"] for l in SS]
GROUP_LINES = {l["n"] for l in SS if l["k"] in GROUP_ON}
GROUP_KEY = "G:гвинтові кріплення"


def contributions(cls):
    """Вклад карточек в строки СС для класса cls: строка -> список (кол в ед. строки, брутто); + поза СС / МШП / группа.
    Многострочные связи: заполняем строки по приоритету до нормы, остаток — на первую строку."""
    lines = collections.defaultdict(list)
    extra = {"poza": collections.defaultdict(lambda: [Z, Z]), "mshp": collections.defaultdict(lambda: [Z, Z]),
             "unmapped": collections.defaultdict(lambda: [Z, Z])}
    group = []
    multi = []
    for c in U:
        m = MAP[c]
        q, g = qg(c, cls)
        if c == KIT:
            continue
        if not q and not g:
            continue
        if m["method"] == "комплект":
            continue
        if m["kind"] == "Групова":
            group.append((q, g)); continue
        if m["kind"] == "поза СС":
            a = extra["poza"][c]; a[0] += q; a[1] += g; continue
        if m["kind"] == "МШП":
            a = extra["mshp"][c]; a[0] += q; a[1] += g; continue
        if not m["links"]:
            a = extra["unmapped"][c]; a[0] += q; a[1] += g; continue
        if len(m["links"]) == 1:
            l = m["links"][0]
            lines[l["line"]].append((q * l["coef"], g, c))
        else:
            multi.append((c, q, g))
    for c, q, g in multi:                            # многострочные: по приоритету до нормы строки
        links = sorted(MAP[c]["links"], key=lambda l: l["prio"])
        rem_q = q
        for i, l in enumerate(links):
            have = sum(x[0] for x in lines[l["line"]])
            room = max(Z, LINE[l["line"]]["q"] - have)
            take = min(rem_q * l["coef"], room) / l["coef"] if l["coef"] else Z
            if take > 0:
                lines[l["line"]].append((take * l["coef"], g * take / q, c)); rem_q -= take
        if rem_q > 0:
            l = links[0]
            lines[l["line"]].append((rem_q * l["coef"], g * rem_q / q, c))
    return lines, extra, group


def kit_state(cls):
    """Комплект септика: факт по складовим (прямо + собранные в напівфабрикат на этом доме); комплектов = мин(кол_i / на_компл_i)."""
    comp = collections.defaultdict(lambda: [Z, Z])
    for c, grp in KIT_CARDS.items():
        if c not in st and c not in cons:
            continue
        q, g = qg(c, cls)
        comp[grp][0] += q; comp[grp][1] += g
        if cls.startswith("F:"):
            a = ASM[(c, cls[2:])]
            comp[grp][0] += a[0]; comp[grp][1] += a[1]
    kits = min((comp[gp][0] / KIT_COMP[gp]) for gp in KIT_COMP)
    return comp, kits


def mapped_house(h):
    out = {}
    linesF, exF, grpF = contributions("F:" + h)
    linesS, exS, grpS = contributions("S:" + h)
    for n in LINES_N:
        if n in GROUP_LINES or n == SEPTIC_LINE:
            continue
        l = LINE[n]
        rows = [(q, g, ("map", h, c)) for q, g, c in linesF.get(n, [])] + [(q, g, ("map", h, c)) for q, g, c in linesS.get(n, [])]
        F = sum(q for q, _, _ in linesF.get(n, [])); Fg = sum(g for _, g, _ in linesF.get(n, []))
        Sh = sum(q for q, _, _ in linesS.get(n, [])); Shg = sum(g for _, g, _ in linesS.get(n, []))
        inn, oq, og = greedy(rows, l["q"])
        rec = {"N": l["q"], "Ns": l["sum"], "F": F, "Fg": Fg, "Sh": Sh, "Shg": Shg, "over_q": oq, "over_g": og,
               "price_ss": l["price"], "econ_q": Z, "econ_g": Z, "D": Z}
        gap = max(Z, l["q"] - F - Sh)
        if h in DONE:
            rec["econ_q"] = gap; rec["econ_g"] = gap * l["price"]
        else:
            rec["D"] = gap
        out[n] = rec
    # групова норма — пул по ОН (как сейчас)
    N = sum(LINE[n]["q"] for n in GROUP_LINES); Ns = sum(LINE[n]["sum"] for n in GROUP_LINES)
    rows = grpF + grpS
    F = sum(q for q, _ in grpF); Sh = sum(q for q, _ in grpS)
    inn, oq, og = greedy(rows, N)
    gap = max(Z, N - F - Sh)
    out[GROUP_KEY] = {"N": N, "Ns": Ns, "F": F, "Fg": sum(g for _, g in grpF), "Sh": Sh, "Shg": sum(g for _, g in grpS),
                      "over_q": oq, "over_g": og, "price_ss": Ns / N, "econ_q": gap if h in DONE else Z,
                      "econ_g": gap * Ns / N if h in DONE else Z, "D": gap if h not in DONE else Z}
    # комплект септика
    l = LINE[SEPTIC_LINE]
    compF, kF = kit_state("F:" + h)
    compS, kS = kit_state("S:" + h)
    kits_have = Decimal(int(kF + kS))
    over_g, over_q, extra_parts = Z, Z, {}
    for gp, per in KIT_COMP.items():
        tot = compF[gp][0] + compS[gp][0]
        allowed = l["q"] * per
        ex = max(Z, tot - allowed)
        if ex > 0 and tot > 0:
            g = (compF[gp][1] + compS[gp][1]) / tot * ex
            over_g += g; extra_parts[gp] = float(ex)
            BY_CARD_OVER[("map", h, "kit:" + gp)] = BY_CARD_OVER.get(("map", h, "kit:" + gp), Z) + g
            over_q += ex / per
    gap = max(Z, l["q"] - kits_have)
    out[SEPTIC_LINE] = {"N": l["q"], "Ns": l["sum"], "F": kF, "Fg": sum(v[1] for v in compF.values()), "Sh": kS,
                        "Shg": sum(v[1] for v in compS.values()), "over_q": over_q, "over_g": over_g, "price_ss": l["price"],
                        "econ_q": gap if h in DONE else Z, "econ_g": gap * l["price"] if h in DONE else Z,
                        "D": gap if h not in DONE else Z, "comp": {gp: float(compF[gp][0] + compS[gp][0]) for gp in KIT_COMP},
                        "extra_parts": extra_parts}
    ex = {kk: {c: v for c, v in d.items()} for kk, d in exF.items()}
    for kk, d in exS.items():
        for c, v in d.items():
            a = ex[kk].setdefault(c, [Z, Z]); a[0] += v[0]; a[1] += v[1]
    return out, ex


MAPPED = {}
MAPPED_EX = {}
for h in HOUSES:
    MAPPED[h], MAPPED_EX[h] = mapped_house(h)

# склад проекта в разрезе строк (для «докупити» с соответствием)
P_lines, P_ex, P_grp = contributions("P")
compP, kitsP = kit_state("P")


def line_last_price(n):
    """Цена последней закупки на ед. строки СС: основная карточка (приоритет 1, затем аналоги) / коэффициент."""
    cands = [x for x in BY_LINE.get(n, []) if x[0] in LAST and x[2]["coef"]]
    cands.sort(key=lambda x: LAST[x[0]]["date"], reverse=True)
    cands.sort(key=lambda x: x[2]["prio"])
    for c, _, l in cands:
        if c in LAST and l["coef"]:
            return LAST[c]["price_g"] / l["coef"], c
    return None, None


def mapped_buy():
    res = {}
    for n in LINES_N:
        if n in GROUP_LINES or n == SEPTIC_LINE:
            continue
        l = LINE[n]
        sp = max(Z, sum(q for q, _, _ in P_lines.get(n, [])))
        rem = sp
        byh = {}
        for h in TODO:
            d = MAPPED[h][n]["D"]
            take = min(d, rem); rem -= take
            byh[h] = {"D": d, "move": take, "buy": d - take}
        sd = sum(v["D"] for v in byh.values())
        lp, lc = line_last_price(n)
        res[n] = {"N": l["q"], "price_ss": l["price"], "price_last": lp, "price_last_card": lc, "Sp": sp, "sumD": sd,
                  "move": min(sd, sp), "buy": max(Z, sd - sp), "byh": byh}
    # групова
    N = sum(LINE[n]["q"] for n in GROUP_LINES); Ns = sum(LINE[n]["sum"] for n in GROUP_LINES)
    sp = max(Z, sum(q for q, _ in P_grp)); rem = sp; byh = {}
    for h in TODO:
        d = MAPPED[h][GROUP_KEY]["D"]; take = min(d, rem); rem -= take
        byh[h] = {"D": d, "move": take, "buy": d - take}
    sd = sum(v["D"] for v in byh.values())
    lp = now_last_price("гвинтові кріплення")
    res[GROUP_KEY] = {"N": N, "price_ss": Ns / N, "price_last": lp, "Sp": sp, "sumD": sd, "move": min(sd, sp), "buy": max(Z, sd - sp), "byh": byh}
    # комплект септика: неттинг по складовим (запасные детали склада проекта используются)
    l = LINE[SEPTIC_LINE]
    spare = {gp: max(Z, compP[gp][0]) for gp in KIT_COMP}
    lp_comp = {}
    for gp in KIT_COMP:
        cs = [c for c, g in KIT_CARDS.items() if g == gp and c in LAST]
        cs.sort(key=lambda c: LAST[c]["date"], reverse=True)
        lp_comp[gp] = LAST[cs[0]]["price_g"] if cs else None
    kit_lp = sum(lp_comp[gp] * per for gp, per in KIT_COMP.items())
    byh = {}
    for h in TODO:
        d = MAPPED[h][SEPTIC_LINE]["D"]
        buy_parts = {}
        for gp, per in KIT_COMP.items():
            need = d * per
            take = min(need, spare[gp]); spare[gp] -= take
            buy_parts[gp] = need - take
        buy_g = sum(buy_parts[gp] * lp_comp[gp] for gp in KIT_COMP)
        buy_ss = buy_g / kit_lp * l["price"]
        byh[h] = {"D": d, "move": d - (buy_g / kit_lp if kit_lp else d), "buy": buy_g / kit_lp if kit_lp else d,
                  "buy_parts": {gp: float(v) for gp, v in buy_parts.items()}, "buy_g_last": buy_g, "buy_g_ss": buy_ss}
    sd = sum(v["D"] for v in byh.values())
    res[SEPTIC_LINE] = {"N": l["q"], "price_ss": l["price"], "price_last": kit_lp, "Sp": Decimal(int(kitsP)), "sumD": sd,
                        "move": sum(v["move"] for v in byh.values()), "buy": sum(v["buy"] for v in byh.values()), "byh": byh,
                        "spare_P": {gp: float(compP[gp][0]) for gp in KIT_COMP}, "lp_comp": {gp: float(v) for gp, v in lp_comp.items()}}
    return res


MAP_BUY = mapped_buy()
