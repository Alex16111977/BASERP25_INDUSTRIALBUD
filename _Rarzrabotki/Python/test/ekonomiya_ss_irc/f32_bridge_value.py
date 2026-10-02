# -*- coding: utf-8 -*-
"""Ф2. Мост в гривнах (BuhBud, без НДС) по классам складов IRC: закупка → склад проекта → склады домов →
списание (комплектация / малоценка / прочее) → остаток. Источник: out/buh_moves.json (f31). Расчёт."""
import os, sys, json, re
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from collections import defaultdict

M = json.load(open(os.path.join(OUT, "buh_moves.json"), encoding="utf-8"))
CUT = sys.argv[1] if len(sys.argv) > 1 else "2026-12-31 23:59:59"   # граница дат включительно

PROJECT = {"ІБ _МД IRC 2026", "ІС _МД IRC 2026", "МД IRС 2026", "Цех МД IRS 2026", "МД IRS 2026"}
MSP = {"ІБ_МШП", "ІС МШП", "ІБ_МШП виробничий"}


def house(name):
    m = re.search(r"(15|30) м №\s*(\d+)", name or "")
    return f"{m.group(1)}м№{m.group(2)}" if m else None


def cls(name):
    if not name:
        return "—"
    if name in PROJECT:
        return "P:проект"
    if name in MSP:
        return "MSP:общий"
    h = house(name)
    if h and "IR" in name:
        low = name.lower()
        kind = "вироб" if "виробнич" in low else ("відвант" if "відвантаж" in low else "монтаж")
        return f"H:{h}:{kind}"
    if "IR" in name:
        return "IRC:прочий:" + name
    return "X:чужой"


def is_mat(acc):
    return acc[:2] in ("20", "22", "28")


flows = defaultdict(Decimal)
for r in M:
    if r["Период"] > CUT:
        continue
    dt, kt = r["СчДт"] or "", r["СчКт"] or ""
    s = D(r["Сумма"])
    t = r["ТипРег"]
    if is_mat(dt) and is_mat(kt):
        flows[("перемещение", cls(r["СклКт"]), cls(r["СклДт"]), t)] += s
    elif is_mat(dt):
        flows[("поступление", "Кт" + kt[:3], cls(r["СклДт"]), t)] += s
    elif is_mat(kt):
        flows[("расход", cls(r["СклКт"]), "Дт" + dt[:3], t)] += s
    else:
        flows[("26/прочее", (kt + ":" + cls(r["СклКт"])), (dt + ":" + cls(r["СклДт"])), t)] += s


def grp(c):
    # свёртка домов до уровня "H:15м" / "H:30м №1-3" / "H:30м №4-7"
    if c.startswith("H:"):
        h = c.split(":")[1]
        n = int(h.split("№")[1])
        if h.startswith("30") and n >= 4:
            return "H30_4-7"
        return "H15" if h.startswith("15") else "H30_1-3"
    return c

agg = defaultdict(Decimal)
for (kind, a, b, t), s in flows.items():
    agg[(kind, grp(a), grp(b), t)] += s
lines = [f"ГРАНИЦА {CUT}"]
for k in sorted(agg, key=lambda x: (x[0], x[1], x[2])):
    lines.append(f"{k[0]:<12} {k[1]:<28} -> {k[2]:<28} {k[3]:<45} {fmt(agg[k]):>15}")
# баланс классов материалов
bal = defaultdict(Decimal)
for (kind, a, b, t), s in flows.items():
    if kind == "перемещение":
        bal[grp(a)] -= s; bal[grp(b)] += s
    elif kind == "поступление":
        bal[grp(b)] += s
    elif kind == "расход":
        bal[grp(a)] -= s
lines.append("ОСТАТОК МАТЕРИАЛОВ ПО КЛАССАМ (приход − расход, расчёт):")
for k, v in sorted(bal.items()):
    lines.append(f"   {k:<35} {fmt(v):>15}")
p = os.path.join(OUT, f"bridge_value_{CUT[:10]}.txt")
open(p, "w", encoding="utf-8").write("\n".join(lines))
print("\n".join(lines))
dump_json({"|".join(k): str(v) for k, v in flows.items()}, f"flows_{CUT[:10]}.json")
