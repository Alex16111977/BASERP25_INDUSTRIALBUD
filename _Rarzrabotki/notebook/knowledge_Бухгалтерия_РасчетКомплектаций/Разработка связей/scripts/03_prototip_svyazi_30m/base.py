# -*- coding: utf-8 -*-
"""Общая база прототипа: загрузка raw.json, классы складов, НДС партий, свёртки по карточкам. Только чтение файлов."""
import os, sys, re, json, collections
from decimal import Decimal, ROUND_HALF_UP
if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = json.load(open(os.path.join(HERE, "raw.json"), encoding="utf-8"))
Z = Decimal(0)
PROJECT_W = {"ІБ _МД IRC 2026", "ІС _МД IRC 2026", "МД IRС 2026", "Цех МД IRS 2026", "МД IRS 2026"}
H30 = [f"30-{i}" for i in range(1, 8)]


def Dd(x):
    return Decimal(str(x)) if x not in (None, "") else Z


def r2(x):
    return float(Dd(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def r3(x):
    return float(Dd(x).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))


K = lambda s: " ".join(str(s or "").replace("\xa0", " ").split()).casefold()


def house(name):
    m = re.search(r"(15|30) м №\s*(\d+)", name or "")
    return f"{m.group(1)}-{m.group(2)}" if (m and "IR" in (name or "")) else None


def wcls(w):
    if not w:
        return "—"
    if w in PROJECT_W:
        return "P"
    h = house(w)
    if h:
        return "H:" + h
    if "МШП" in w:
        return "MSP"
    if "IR" in w:
        return "IRC?"
    return "X"


is_mat = lambda a: (a or "")[:2] in ("20", "22", "28")


class Vat:
    """НДС партии (перенос moy_byudzhet_11.Vat / f40): приход -> НДС строки; перемещение -> рекурсивно по источнику."""
    def __init__(self, receipts, moves):
        self.rc, self.rcdoc, self.rate_nom = collections.defaultdict(lambda: [Z, Z]), collections.defaultdict(lambda: [Z, Z]), {}
        for r in receipts:
            net = Dd(r["Сумма"]) - (Dd(r["НДС"]) if r["ВклНДС"] else Z)
            self.rc[(r["Док"], r["Код"])][0] += net; self.rc[(r["Док"], r["Код"])][1] += Dd(r["НДС"])
            self.rcdoc[r["Док"]][0] += net; self.rcdoc[r["Док"]][1] += Dd(r["НДС"])
            self.rate_nom[r["Код"]] = r["СтавкаНом"]
        self.src = collections.defaultdict(list)
        for r in moves:
            if is_mat(r["СчДт"]) and is_mat(r["СчКт"]):
                self.src[(r["ПартДт"], r["КодДт"], r["СклДт"])].append((r["ПартКт"], r["КодКт"], r["СклКт"], Dd(r["Сумма"])))
        self.memo, self.fallback = {}, set()

    def ratio(self, part, code, skl, depth=0):
        key = (part, code, skl)
        if key in self.memo:
            return self.memo[key]
        res = None
        if (part, code) in self.rc and self.rc[(part, code)][0] != 0:
            res = self.rc[(part, code)][1] / self.rc[(part, code)][0]
        elif part in self.rcdoc and self.rcdoc[part][0] != 0:
            res = self.rcdoc[part][1] / self.rcdoc[part][0]
        elif part.startswith("Авансовий звіт"):
            res = Z
        elif key in self.src and depth < 12:
            tot = sum(s for *_, s in self.src[key])
            if tot:
                res = sum(s * self.ratio(p, c, w, depth + 1) for p, c, w, s in self.src[key]) / tot
        if res is None:
            rt = self.rate_nom.get(code)
            res = Decimal("0.2") if rt in (None, "20%", "") else (Decimal("0.07") if rt == "7%" else Z)
            self.fallback.add(key)
        self.memo[key] = res
        return res


MV = RAW["moves"]
VAT = Vat(RAW["receipts"], MV)

# карточки: код -> {name, on, unit}
CARD = {}
for r in MV:
    for s in ("Дт", "Кт"):
        if r["Код" + s]:
            CARD.setdefault(r["Код" + s], {"name": r["Ном" + s], "on": r["ОН" + s], "unit": r["Ед" + s]})
for r in RAW["receipts"]:
    CARD.setdefault(r["Код"], {"name": r["Ном"], "on": r["ОН"], "unit": ""})

# свёртки по карточке: st[код][класс] = [кол, нетто, брутто]; cons[код][дом] = [кол, нетто, брутто] (Кт мат -> Дт не мат со склада дома)
st = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z]))
cons = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z]))
asm = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z, Z]))   # сборка комплекта: Кт мат -> Дт мат другой карточки
for r in MV:
    dt_, kt_ = r["СчДт"] or "", r["СчКт"] or ""
    s = Dd(r["Сумма"])
    if is_mat(dt_):
        c_ = wcls(r["СклДт"])
        g = s * (1 + VAT.ratio(r["ПартДт"], r["КодДт"], r["СклДт"]))
        a = st[r["КодДт"]][c_]; a[0] += Dd(r["КолДт"] or 0); a[1] += s; a[2] += g
    if is_mat(kt_):
        c_ = wcls(r["СклКт"])
        g = s * (1 + VAT.ratio(r["ПартКт"], r["КодКт"], r["СклКт"]))
        a = st[r["КодКт"]][c_]; a[0] -= Dd(r["КолКт"] or 0); a[1] -= s; a[2] -= g
        if not is_mat(dt_) and c_.startswith("H:"):
            b = cons[r["КодКт"]][c_[2:]]; b[0] += Dd(r["КолКт"] or 0); b[1] += s; b[2] += g
        if is_mat(dt_) and r["КодДт"] != r["КодКт"] and wcls(r["СклДт"]) == c_:
            b = asm[r["КодКт"]][r["КодДт"]]; b[0] += Dd(r["КолКт"] or 0); b[1] += s; b[2] += g

# последняя закупка (BuhBud ПТУ, проведённые): код -> {date, price_g, price_n, kontr}
def _d(s):
    a, b = (s + " 0:0:0").split(" ")[:2]
    d, m, y = a.split(".")
    hh = b.split(":")
    return f"{y}-{m}-{d} {int(hh[0]):02d}:{int(hh[1]):02d}:{int(hh[2]):02d}"


LAST = {}
for r in RAW["receipts"]:
    if not r["Проведен"] or not Dd(r["Кол"]):
        continue
    net = Dd(r["Сумма"]) - (Dd(r["НДС"]) if r["ВклНДС"] else Z)
    dts = _d(r["Дата"])
    cur = LAST.get(r["Код"])
    if cur is None or dts > cur["date"]:
        LAST[r["Код"]] = {"date": dts, "price_n": net / Dd(r["Кол"]), "price_g": (net + Dd(r["НДС"])) / Dd(r["Кол"]),
                          "kontr": r["Контрагент"], "doc": r["Док"]}
