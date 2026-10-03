# -*- coding: utf-8 -*-
"""Z6: таблицы Markdown из data/proekt_irc2026_zakupki.json → z6_tables.md (только чтение JSON)."""
import os, sys, json
if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, "..", "..", "..", "data", "proekt_irc2026_zakupki.json"), encoding="utf-8"))
T = D["totals"]
out = []


def fm(v, d=2):
    if v is None:
        return "—"
    s = f"{abs(float(v)):,.{d}f}".replace(",", " ").replace(".", ",")
    return ("−" if float(v) < 0 else "") + s


def fq(v):
    if v is None:
        return "—"
    v = float(v)
    if abs(v - round(v)) < 1e-9:
        return fm(v, 0)
    return fm(v, 3).rstrip("0").rstrip(",")


def P(*a):
    out.append(" ".join(str(x) for x in a))


P("# Z6 — таблицы расчёта закупок проекта IRC 2026")
P("")
P("Снимок:", D["snapshot"]["datetime"], "— движений", D["snapshot"]["moves"], ", последнее", D["snapshot"]["moves_max_period"], ".")
P("")
P("## T1. По узлам строителя (без ПДВ; «докупить» — з ПДВ и без ПДВ, по цене последней закупки)")
P("")
P("| Узел | Карт. | Потребн. 15 м | Потребн. 30 м | Закуплено | Списано 15 м | Списано 30 м | Лежит дома | Лежит склад пр. | Докупить з ПДВ | Докупить без ПДВ | Лишнее |")
P("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
S = {k: 0.0 for k in ("cards", "lines_no_card", "need_15_uah_net", "need_30_uah_net", "purchased_uah_net", "written_off_15_uah_net", "written_off_30_uah_net",
                      "lies_houses_uah_net", "lies_project_uah_net", "to_buy_30_uah_last_vat", "to_buy_30_uah_last_net", "surplus_uah_net")}
for nd, a in D["by_node"].items():
    cards = f"{a['cards']}" + (f" + {a['lines_no_card']} стр." if a["lines_no_card"] else "")
    P(f"| {nd} | {cards} | {fm(a['need_15_uah_net'])} | {fm(a['need_30_uah_net'])} | {fm(a['purchased_uah_net'])} | {fm(a['written_off_15_uah_net'])} | "
      f"{fm(a['written_off_30_uah_net'])} | {fm(a['lies_houses_uah_net'])} | {fm(a['lies_project_uah_net'])} | {fm(a['to_buy_30_uah_last_vat'])} | "
      f"{fm(a['to_buy_30_uah_last_net'])} | {fm(a['surplus_uah_net'])} |")
    for k in S:
        S[k] += a.get(k, 0) or 0
P(f"| **Итого** | **{int(S['cards'])} + {int(S['lines_no_card'])} стр.** | **{fm(S['need_15_uah_net'])}** | **{fm(S['need_30_uah_net'])}** | **{fm(S['purchased_uah_net'])}** | "
  f"**{fm(S['written_off_15_uah_net'])}** | **{fm(S['written_off_30_uah_net'])}** | **{fm(S['lies_houses_uah_net'])}** | **{fm(S['lies_project_uah_net'])}** | "
  f"**{fm(S['to_buy_30_uah_last_vat'])}** | **{fm(S['to_buy_30_uah_last_net'])}** | **{fm(S['surplus_uah_net'])}** |")
P("")
P("## T2. Топ-20 «ещё докупить для 30 м» (з ПДВ)")
P("")
P("| № | Карточка | Ед. | Узел | Остаток потребн. №3–№7 | Склад проекта | Купить | Посл., з ПДВ | СС, з ПДВ | Примечание |")
P("|---:|---|---|---|---:|---:|---:|---:|---:|---|")
for i, x in enumerate(T["top20_to_buy"], 1):
    note = "строка эталона без выдачи (жёлтая)" if x.get("from_etalon_lines_not_issued") else ("нестабильно (CV > 0,3)" if x.get("unstable") else "")
    P(f"| {i} | {x['card']} | {x['unit']} | {x['node']} | {fq(x.get('remaining_need_3_7'))} | {fq(x.get('stock_project'))} | {fq(x['to_buy'])} | "
      f"{fm(x['uah_last_vat'])} | {fm(x['uah_ss_vat'])} | {note} |")
P("")
P("## T3. Топ-20 «останется лишним» (без ПДВ, по себестоимости склада проекта)")
P("")
P("| № | Карточка | Ед. | Узел | Склад проекта | Остаток потребн. 30 м | Лишнее, кол. | Лишнее, грн | Потребн. 15 м (факт) |")
P("|---:|---|---|---|---:|---:|---:|---:|---:|")
for i, x in enumerate(T["top20_surplus"], 1):
    P(f"| {i} | {x['card']} | {x['unit']} | {x['node']} | {fq(x['stock_project'])} | {fq(x['remaining_need_3_7'])} | {fq(x['surplus'])} | {fm(x['uah_net'])} | {fq(x['need_15'])} |")
P("")
P("## T4. По домам: списано и лежит (без ПДВ)")
P("")
P("| Дом | Списано | Лежит на складах дома | Карточек списано |")
P("|---|---:|---:|---:|")
for h, v in T["by_house"].items():
    if v["written_off_uah_net"] or v["lies_uah_net"]:
        P(f"| {h} | {fm(v['written_off_uah_net'])} | {fm(v['lies_uah_net'])} | {v['cards_written_off']} |")
P("")
P("## T5. Закуплено, прочие приходы и расходы (без ПДВ)")
P("")
P("| Вид | Карточек | Документов | Без ПДВ | З ПДВ |")
P("|---|---:|---:|---:|---:|")
for k, v in T["purchased_by_kind"].items():
    P(f"| Закуплено: {k} | {v['cards']} | {v['docs']} | {fm(v['uah_net'])} | {fm(v['uah_vat'])} |")
for k, v in sorted(T["other_in_by_kind"].items(), key=lambda x: -x[1]["uah_net"]):
    P(f"| Прочий приход: {k} | {v['cards']} | {v['docs']} | {fm(v['uah_net'])} | |")
for k, v in sorted(T["other_out_by_kind"].items(), key=lambda x: -x[1]["uah_net"]):
    P(f"| Прочий расход: {k} | {v['cards']} | {v['docs']} | {fm(v['uah_net'])} | |")
P("")
P("## T6. 25 строк эталона 30 м без выдачи на №1–№2 (решение 7) и опыт 15 м")
P("")
P("| Строка СС | Норма на дом | Купить на №3–№7 | Z1, посл. | Карточка 30 м | Карточка-кандидат по 15 м | Списано на 15 м |")
P("|---|---:|---:|---:|---|---|---:|")
for x in D["etalon_lines_not_issued"]:
    c15 = x.get("candidate_card_by_15m_experience") or {}
    P(f"| {x['line']} {x['text']} | {fq(x['norm_per_house'])} {x['unit']} | {fq(x['z1_buy_line_units'])} | {fm(x['z1_uah_last'])} | {x['card_name'] or '—'} | "
      f"{c15.get('card', '—')} | {fq(c15.get('need_15_fact')) if c15 else '—'} |")
P("")
P("## T7. Строки по карточкам с невязкой сверки замкнутости")
P("")
bad = [r for r in D["rows"] if abs(r["balance_diff"]["qty"] or 0) > 0.0005 or abs(r["balance_diff"]["uah_net"] or 0) > 0.05]
P("Невязок нет." if not bad else "\n".join(f"- {r['code']} {r['card']}: {r['balance_diff']}" for r in bad))
P("")
P("## T8. Узлы для карточек и строк, которых нет в закупочной 15 м (решения прораба Z6)")
P("")
P("| Код | Карточка / строка | Группа | Строка СС 30 м | Узел | Как определён | Потребн. 30 м | Склад проекта |")
P("|---|---|---|---|---|---|---:|---:|")
for r in D["rows"]:
    if r["node_how"].startswith("як у закупівельній 15 м"):
        continue
    P(f"| {r['code']} | {r['card']} | {r['group']} | {r.get('ss_line_30') or '—'} | {r['node']} | {r['node_how']} | {fq(r['need_30'])} | {fq(r['stock']['project'])} |")
P("")
open(os.path.join(HERE, "z6_tables.md"), "w", encoding="utf-8").write("\n".join(out))
print("\n".join(out[:60]))
