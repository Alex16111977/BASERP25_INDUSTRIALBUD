# -*- coding: utf-8 -*-
"""Ф1–Ф2. Ядро: по каждой «Загальній назві» IRC — норма СС, закупки ERP, движение в BuhBud (склад проекта → склады
домов → списание → остаток), расход по домам (с реальным НДС партии и «як план» по ставке номенклатуры),
данные отчёта по нормам. Только расчёт по выгрузкам f10/f20/f31/f33/f34 (базы не трогает).
Выход: out/core.json + печать контрольных сумм."""
import os, sys, json, re
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from collections import defaultdict

L = lambda n: json.load(open(os.path.join(OUT, n), encoding="utf-8"))
SS = L("ss_erp.json"); PU = L("purch_erp.json"); MV = L("buh_moves.json"); RC = L("buh_receipts.json")
NR = [r for r in L("norm_report_rows.json")["rows"] if "IR" in (r["Подразделение"] or "")]
CUT30 = "2026-09-30 23:59:59"

def K(s):
    return " ".join(str(s or "").split()).casefold()

def house(name):
    m = re.search(r"(15|30) м №\s*(\d+)", name or "")
    return f"{m.group(1)}-{m.group(2)}" if (m and "IR" in (name or "")) else None

SCOPE = [f"15-{i}" for i in range(1, 7)] + ["30-1", "30-2"]   # границы заказчика 02.10.2026
OUTSC = [f"30-{i}" for i in range(4, 8)]
DONE15 = [f"15-{i}" for i in range(1, 7)]
PROJECT = {"ІБ _МД IRC 2026", "ІС _МД IRC 2026", "МД IRС 2026", "Цех МД IRS 2026", "МД IRS 2026"}

def cls(w):
    if not w: return "—"
    if w in PROJECT: return "P"
    h = house(w)
    if h: return "H:" + h
    if "МШП" in w: return "MSP"
    if "IR" in w: return "IRC?"
    return "X"

is_mat = lambda a: (a or "")[:2] in ("20", "22", "28")

ON = defaultdict(lambda: defaultdict(Decimal))   # ON[key][metric]
NAME = {}                                         # key -> отображаемое имя
UNIT = defaultdict(set)

def put(key, metric, val, name=None):
    ON[key][metric] += D(val)
    if name and key not in NAME: NAME[key] = " ".join(str(name).split())

# --- A. норма СС (ТЧ-ОН, как в отчёте директору); на дом по типу и итоги
for r in SS["kom"]:
    t = "15" if "15 м" in r["СС"] else "30"
    k = K(r["ОН"]) or ("номсс|" + K(r["НомСС"]))
    put(k, f"n{t}_q", r["Кол"], r["ОН"] or r["НомСС"]); put(k, f"n{t}_s", r["Сумма"])
    UNIT[k].add("СС:" + str(r["ЕдОН"]))
for k in ON:
    ON[k]["n9_q"] = ON[k]["n15_q"] * 6 + ON[k]["n30_q"] * 2; ON[k]["n9_s"] = ON[k]["n15_s"] * 6 + ON[k]["n30_s"] * 2   # 8 домов в границах
    ON[k]["n13_q"] = ON[k]["n15_q"] * 6 + ON[k]["n30_q"] * 7; ON[k]["n13_s"] = ON[k]["n15_s"] * 6 + ON[k]["n30_s"] * 7

# --- B. закупки ERP (карточное ОН), цены по строкам
pr = defaultdict(list)
for r in PU:
    k = K(r["ОН"]) or ("ном|" + K(r["Ном"]))
    put(k, "pu_q", r["Кол"], r["ОН"] or r["Ном"]); put(k, "pu_g", r["Стоимость"]); put(k, "pu_n", r["СтоимостьБезНДС"])
    UNIT[k].add("ERP:" + str(r["Ед"]))
    if D(r["Кол"]) > 0: pr[k].append(D(r["Стоимость"]) / D(r["Кол"]))
for k, v in pr.items():
    ON[k]["pu_pmin"] = min(v); ON[k]["pu_pmax"] = max(v)

# --- C. НДС партий BuhBud: приход (док, код) -> (нетто, НДС); цепочка перемещений
rc = defaultdict(lambda: [D(0), D(0)]); rcdoc = defaultdict(lambda: [D(0), D(0)]); rate_nom = {}
for r in RC:
    rc[(r["Док"], r["Код"])][0] += D(r["Нетто"]); rc[(r["Док"], r["Код"])][1] += D(r["НДС"])
    rcdoc[r["Док"]][0] += D(r["Нетто"]); rcdoc[r["Док"]][1] += D(r["НДС"])
    rate_nom[r["Код"]] = r["СтавкаНом"]
src = defaultdict(list)   # (партия-новая, код, склад) -> [(партия-источник, код, склад, сумма)]
for r in MV:
    if is_mat(r["СчДт"]) and is_mat(r["СчКт"]):
        src[(r["ПартДт"], r["КодДт"], r["СклДт"])].append((r["ПартКт"], r["КодКт"], r["СклКт"], D(r["Сумма"])))
memo = {}
unresolved = defaultdict(Decimal)

def vat_ratio(part, code, skl, depth=0):
    key = (part, code, skl)
    if key in memo: return memo[key]
    res = None
    if (part, code) in rc and rc[(part, code)][0] != 0:
        res = rc[(part, code)][1] / rc[(part, code)][0]
    elif part in rcdoc and rcdoc[part][0] != 0:
        res = rcdoc[part][1] / rcdoc[part][0]
    elif part.startswith("Авансовий звіт"):
        res = D(0)
    elif key in src and depth < 12:
        tot = sum(s for *_, s in src[key])
        if tot:
            res = sum(s * vat_ratio(p, c, w, depth + 1) for p, c, w, s in src[key]) / tot
    if res is None:
        rt = rate_nom.get(code)
        res = D("0.2") if rt in (None, "20%") else (D("0.07") if rt == "7%" else D(0))
        unresolved[part.split(" ")[0]] += 1
    memo[key] = res
    return res

def rep_rate(code):
    rt = rate_nom.get(code)
    return D("1.2") if rt == "20%" else (D("1.07") if rt == "7%" else (D("1.14") if rt == "14%" else (D(1) if rt else None)))

# --- D. движения BuhBud по ОН (карточное ОН), количества и нетто
flows = defaultdict(lambda: defaultdict(Decimal))    # flows[k][(вид, откуда, куда, ≤30.09?)] -> (q, n)
cons_rows = []
for r in MV:
    dt, kt = r["СчДт"] or "", r["СчКт"] or ""
    s = D(r["Сумма"]); le = r["Период"] <= CUT30
    if is_mat(dt) and is_mat(kt):
        k = K(r["ОНКт"]) or ("ном|" + K(r["НомКт"]))
        a, b = cls(r["СклКт"]), cls(r["СклДт"])
        if a != b or a.startswith("H"):
            put(k, f"mv|{a}>{b}|q", r["КолКт"], r["ОНКт"] or r["НомКт"]); put(k, f"mv|{a}>{b}|n", s)
    elif is_mat(dt):
        k = K(r["ОНДт"]) or ("ном|" + K(r["НомДт"]))
        b = cls(r["СклДт"])
        put(k, f"in|{kt[:3]}>{b}|q", r["КолДт"], r["ОНДт"] or r["НомДт"]); put(k, f"in|{kt[:3]}>{b}|n", s)
    elif is_mat(kt):
        k = K(r["ОНКт"]) or ("ном|" + K(r["НомКт"]))
        a = cls(r["СклКт"]); t = r["ТипРег"]
        typ = "компл" if t.startswith("Комплектац") else ("пма" if t.startswith("Передача мало") else "прочее")
        put(k, f"out|{a}|{typ}|q", r["КолКт"], r["ОНКт"] or r["НомКт"]); put(k, f"out|{a}|{typ}|n", s)
        if a.startswith("H:"):
            ratio = vat_ratio(r["ПартКт"], r["КодКт"], r["СклКт"])
            rr = rep_rate(r["КодКт"])
            put(k, f"c|{a[2:]}|q", r["КолКт"]); put(k, f"c|{a[2:]}|n", s)
            put(k, f"c|{a[2:]}|gt", s * (1 + ratio))
            put(k, f"c|{a[2:]}|gr", s * (rr if rr is not None else D("1.2")))
            UNIT[k].add("BUH:" + str(r["ЕдКт"]))
            if not le: put(k, f"c_after30|{a[2:]}|n", s)

# --- D2. остатки по партиям (нетто и реальный брутто) по классам складов + НДС-артефакт расхода
stk = defaultdict(Decimal)   # (класс, партия, код, склад) -> нетто
stkq = defaultdict(Decimal)
for r in MV:
    dt, kt = r["СчДт"] or "", r["СчКт"] or ""
    s = D(r["Сумма"])
    if is_mat(dt):
        kk = (cls(r["СклДт"]), r["ПартДт"], r["КодДт"], r["СклДт"], K(r["ОНДт"]) or ("ном|" + K(r["НомДт"])))
        stk[kk] += s; stkq[kk] += D(r["КолДт"] or 0)
    if is_mat(kt):
        kk = (cls(r["СклКт"]), r["ПартКт"], r["КодКт"], r["СклКт"], K(r["ОНКт"]) or ("ном|" + K(r["НомКт"])))
        stk[kk] -= s; stkq[kk] -= D(r["КолКт"] or 0)
STK = defaultdict(lambda: [Decimal(0), Decimal(0)])
for (c, part, code, skl, k), n in stk.items():
    if n == 0 and stkq[(c, part, code, skl, k)] == 0: continue
    g = n * (1 + vat_ratio(part, code, skl))
    grp = "P" if c == "P" else (("H_scope" if c[2:] in SCOPE else "H_out") if c.startswith("H:") else c)
    STK[grp][0] += n; STK[grp][1] += g
    put(k, f"stk|{grp}|n", n); put(k, f"stk|{grp}|gt", g); put(k, f"stk|{grp}|q", stkq[(c, part, code, skl, k)])
    if c.startswith("H:"):
        put(k, f"stkh|{c[2:]}|n", n); put(k, f"stkh|{c[2:]}|gt", g); put(k, f"stkh|{c[2:]}|q", stkq[(c, part, code, skl, k)])
art = defaultdict(Decimal)
for r in MV:
    dt, kt = r["СчДт"] or "", r["СчКт"] or ""
    if is_mat(kt) and not is_mat(dt) and cls(r["СклКт"]).startswith("H:") and cls(r["СклКт"])[2:] in SCOPE:
        s = D(r["Сумма"]); true = s * (1 + vat_ratio(r["ПартКт"], r["КодКт"], r["СклКт"])); rr = rep_rate(r["КодКт"]) or D("1.2")
        d = s * rr - true
        art["завышение (ставка ном. > реальн.)" if d > 0 else "занижение (ставка ном. < реальн.)"] += d
        art["15м" if cls(r["СклКт"])[2:].startswith("15") else "30м"] += d

# --- E. отчёт по нормам: план/факт/экономия/остатки по (домик, ОН) ровно как в отчёте
nr_tot = defaultdict(lambda: defaultdict(Decimal))
for r in NR:
    t = "15" if "15 м" in r["Подразделение"] else "30"
    hn = (r["Домик"] or "").replace("№", "").strip()
    h = f"{t}-{hn}"
    k = K(r["ОбщееНазвание"]) or ("анал|" + K(r["Аналитика"]))
    vd = r["ВидДокумента"]
    if vd.startswith("6."):
        if h == "15-1":    # строки «6.» тиражированы по домикам и драйверам — берём одну копию
            put(k, "nr_projstock_q", r["ОстатокЗакупкиКол"]); put(k, "nr_projstock_g", r["ОстатокЗакупкиГрн"])
        continue
    for c in ("ПланКол", "ПланГрн", "ФактВНормеКол", "ФактВНормеГрн", "ФактПонадКол", "ФактПонадГрн",
              "ОстатокОбъектаКол", "ОстатокОбъектаГрн", "ЭкономияКол", "ЭкономияГрн"):
        v = D(r[c] or 0)
        if v:
            put(k, f"nr|{h}|{c}", v, r["ОбщееНазвание"] or r["Аналитика"])
            nr_tot[h][c] += v

core = {"names": {k: NAME.get(k, k) for k in ON}, "units": {k: sorted(v) for k, v in UNIT.items()},
        "on": {k: {m: str(v) for m, v in d.items()} for k, d in ON.items()}}
dump_json(core, "core.json")

# --- контрольные суммы
T = defaultdict(Decimal)
for k, d in ON.items():
    for m, v in d.items():
        T[m] += v
print("Норма 15м/дом", fmt(T["n15_s"]), " 30м/дом", fmt(T["n30_s"]), " 8 домов", fmt(T["n9_s"]), " 13 домов", fmt(T["n13_s"]))
print("Закупки ERP: брутто", fmt(T["pu_g"]), "нетто", fmt(T["pu_n"]))
for h in SCOPE + ["30-4"]:
    print(f"Дом {h}: расход нетто {fmt(T[f'c|{h}|n'])}  брутто реальн. {fmt(T[f'c|{h}|gt'])}  брутто «як план» {fmt(T[f'c|{h}|gr'])}"
          f" | отчёт: факт {fmt(nr_tot[h]['ФактВНормеГрн'] + nr_tot[h]['ФактПонадГрн'])} норма {fmt(nr_tot[h]['ПланГрн'])} экономия {fmt(nr_tot[h]['ЭкономияГрн'])}")
print("Остаток склада проекта по отчёту (одна копия):", fmt(T["nr_projstock_g"]))
print("Остатки по партиям (нетто / реальный брутто):", {k: (fmt(v[0]), fmt(v[1])) for k, v in STK.items()})
print("НДС-артефакт отчёта по нормам (як план − реальный брутто), 8 домов:", {k: fmt(v) for k, v in art.items()})
print("Нерешённые партии (взята ставка номенклатуры):", dict(unresolved))
mvk = defaultdict(Decimal)
for m, v in T.items():
    if (m.startswith("mv|") or m.startswith("in|") or m.startswith("out|")) and m.endswith("|n"):
        mvk[m] += v
for m in sorted(mvk): print("  ", m, fmt(mvk[m]))
