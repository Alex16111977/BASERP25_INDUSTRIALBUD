# -*- coding: utf-8 -*-
"""Markdown-таблицы для ANALIZ.md из out/*.json (без ручного переноса чисел). Выход: out/md_tables.md."""
import os, sys, json
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from collections import defaultdict

L = lambda n: json.load(open(os.path.join(OUT, n), encoding="utf-8"))
A = L("analysis.json")["rows"]; F = L("final.json"); C = L("core.json"); MVP = L("money_vs_purch.json")
ON = {k: defaultdict(Decimal, {m: D(v) for m, v in d.items()}) for k, d in C["on"].items()}
EV = L("economist_vs_ss.json")
econ = defaultdict(lambda: defaultdict(Decimal))
for sh, t in (("IRS 15", "15"), ("IRS 30", "30")):
    for x in EV[sh]["rows"]:
        k = " ".join(str(x["ОН"] or "").split()).casefold()
        econ[k][f"q{t}"] += D(x["q"]); econ[k][f"s{t}"] += D(x["s"])
dv = lambda a, b: a / b if b else None
f2 = lambda x: fmt(x) if x is not None else "—"
f1 = lambda x: fmt(x, 1) if x is not None else "—"
f3 = lambda x: (fmt(x, 3).rstrip("0").rstrip(",") if x is not None else "—")
out = []
w = out.append

def unit(r):
    u = [x.split(":", 1)[1].strip() for x in r["units"] if x.startswith("СС:")]
    return u[0] if u else ""

def tab15(rows, g, nh, plan_h, title):
    w(f"\n#### {title}\n")
    w("| Загальна назва | од. | Норма/дім к-сть | Ціна СС | Норма/дім грн | Списано/дім сер (мін–макс) | Ціна факт (брутто) | Ціна закупки ERP | Ціна економіста | Δ к-сть % | Δ ціна % | Ціновий ефект | Кількісний ефект | Δ грн | Категорія |")
    w("|---|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|---|")
    for r in rows:
        o = ON[r["k"]]
        nq = o[f"nr|{plan_h}|ПланКол"]; ns = o[f"nr|{plan_h}|ПланГрн"]
        cq = D(r[f"{g}_cq"]); cg = D(r[f"{g}_cg"])
        e = econ.get(r["k"], {})
        ep = dv(e.get("s" + plan_h[:2], Decimal(0)), e.get("q" + plan_h[:2], Decimal(0)))
        w(f"| {r['name']} | {unit(r)} | {f3(nq)} | {f2(dv(ns, nq))} | {f2(ns)} | {f3(cq / nh)} ({f3(D(r[g + '_min']))}–{f3(D(r[g + '_max']))}) | "
          f"{f2(dv(cg, cq))} | {f2(dv(D(r['pu_g']), D(r['pu_q'])))} | {f2(ep)} | {f1(D(r[g + '_dq_pct']) if r[g + '_dq_pct'] not in (None, 'None') else None)} | "
          f"{f1(D(r[g + '_dp_pct']) if r[g + '_dp_pct'] not in (None, 'None') else None)} | {f2(D(r[g + '_pe']))} | {f2(D(r[g + '_qe']))} | {f2(D(r[g + '_delta']))} | {classify_row(r, g)} |")

top15 = sorted(A.values(), key=lambda r: -abs(D(r["d15_delta"])))[:25]
tab15(top15, "d15", 6, "15-1", "Ф1-А. 15 м №1–№6: 25 назв с наибольшим |Δ| (расход реальный брутто против нормы ×6)")
top30 = sorted(A.values(), key=lambda r: -abs(D(r["d30_delta"])))[:15]
tab15(top30, "d30", 2, "30-1", "Ф1-Б. 30 м №1–№2: 15 назв с наибольшим |Δ| (расход против нормы ×2)")

w("\n#### Ф1-В. Отчёт директору (расчёт по его формулам, 13 домов): 12 крупнейших экономий и перерасходов\n")
w("| Загальна назва | План 13, к-сть | Ціна план | Закуплено к-сть | Ціна факт | Ціновий чинник | Кількісний чинник | ETC | Відхилення прогнозу |")
w("|---|--:|--:|--:|--:|--:|--:|--:|--:|")
for r in sorted(A.values(), key=lambda r: D(r["dev"]))[:12] + sorted(A.values(), key=lambda r: -D(r["dev"]))[:12]:
    w(f"| {r['name']} | {f3(D(r['n13_q']))} | {f2(dv(D(r['n13_s']), D(r['n13_q'])))} | {f3(D(r['pu_q']))} | {f2(dv(D(r['pu_g']), D(r['pu_q'])))} | "
      f"{f2(D(r['dir_pf']))} | {f2(D(r['dir_qf']))} | {f2(D(r['etc']))} | {f2(D(r['dev']))} |")

w("\n#### ТОП-80 % экономии по расходу 15 м №1–№6\n")
w("| № | Загальна назва | Δ, грн | Накопл. % |")
w("|--:|---|--:|--:|")
for i, (n, d, p) in enumerate(F["top15"], 1):
    w(f"| {i} | {n} | {fmt(d)} | {fmt(p, 1)} |")
w("\n#### ТОП «План на факт отчёта − дефицит по фактическим остаткам»\n")
w("| № | Загальна назва | ETC − дефицит, грн | Накопл. % |")
w("|--:|---|--:|--:|")
for i, (n, d, p) in enumerate(F["topetc"], 1):
    w(f"| {i} | {n} | {fmt(d)} | {fmt(p, 1)} |")
w("\n#### Поставщики: закупки IRC против оплат из казны (10 крупнейших долгов и все авансы > 10 000)\n")
w("| Контрагент | Закуплено (ERP, з ПДВ) | Оплачено ≤30.09 (казна) | Сальдо: + долг / − аванс |")
w("|---|--:|--:|--:|")
ls = MVP["lines"]
for x in ls[:10] + [x for x in ls if D(x["Сальдо(+долг/−аванс)"]) < -10000]:
    w(f"| {x['Контрагент']} | {fmt(x['Закуплено'])} | {fmt(x['Оплачено≤30.09'])} | {fmt(x['Сальдо(+долг/−аванс)'])} |")
p = os.path.join(OUT, "md_tables.md")
open(p, "w", encoding="utf-8").write("\n".join(out))
print(p, len(out))
