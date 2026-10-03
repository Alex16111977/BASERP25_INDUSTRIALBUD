# -*- coding: utf-8 -*-
"""Z1. Таблицы markdown для заметки Z1 из data/zakupochnaya_spec_30m.json (только чтение JSON). Выход: z1_tables.md рядом."""
import os, sys, json
if (sys.stdout.encoding or "").lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
J = json.load(open(os.path.join(HERE, "..", "..", "..", "data", "zakupochnaya_spec_30m.json"), encoding="utf-8"))
OUT = []


def F(x, d=2):
    if x is None:
        return "—"
    s = f"{x:,.{d}f}".replace(",", " ").replace(".", ",")
    return s


def Q(x):
    if x is None:
        return "—"
    return (f"{x:,.3f}".rstrip("0").rstrip(".") if x != int(x) else f"{int(x):,}").replace(",", " ").replace(".", ",")


def p(s=""):
    OUT.append(s)


H = ["№3", "№4", "№5", "№6", "№7"]
T = J["totals"]; C = J["compare"]

p("## T1. Итоги по группам")
p("| Группа | Карточек | К докупке | (б) по последней закупке | (б) по цене СС | (а) по последней закупке | (а) по цене СС | Среднее на дом (б), посл. закупка |")
p("|---|---|---|---|---|---|---|---|")
for g in ["еталон", "поза еталоном", "МШП", "всього"]:
    b = T["variant_b_main"][g]; a = T["variant_a_spysano_only"][g]
    p(f"| {g} | {b['cards']} | {b['cards_to_buy']} | {F(b['uah_last'])} | {F(b['uah_ss'])} | {F(a['uah_last'])} | {F(a['uah_ss'])} | {F(b['avg_house_uah'])} |")

p("\n## T2. По домам (FIFO склада проекта №3→№7)")
bh = C["by_house"]; hy = C["hybrid_zakupochnaya_plus_etalon_lines_not_issued"]
p("| Дом | План | Закупочная, посл. | Закупочная, СС | Гибрид, посл. | Гибрид, СС | Шахматка, СС | Шахматка, посл. | Эталон со связью, СС | Эталон со связью, посл. | Бюджет v3 |")
p("|---|---|---|---|---|---|---|---|---|---|---|")
tot = {k: 0 for k in range(10)}
for h in H:
    vals = [bh["zakupochnaya_uah_last"][h], bh["zakupochnaya_uah_ss"][h], hy["by_house_uah_last"][h], hy["by_house_uah_ss"][h],
            bh["shahmatka_uah_ss"][h], bh["shahmatka_uah_last"][h], bh["etalon_mapped_uah_ss"][h], bh["etalon_mapped_uah_last"][h], bh["budget_v3_uah"][h]]
    for i, v in enumerate(vals):
        tot[i] += v
    p(f"| {h} | {T['plan_month'][h]} | " + " | ".join(F(v) for v in vals) + " |")
d = C["dokupyty_3_7"]
p(f"| **Итого** | | **{F(d['zakupochnaya_uah_last'])}** | **{F(d['zakupochnaya_uah_ss'])}** | **{F(hy['uah_last'])}** | **{F(hy['uah_ss'])}** | {F(d['shahmatka_uah_ss'])} | {F(d['shahmatka_uah_last'])} | {F(d['etalon_mapped_uah_ss'])} | {F(d['etalon_mapped_uah_last'])} | {F(d['budget_v3_uah'])} |")

p("\n## T3. Разница с эталоном «со связью» по группам")
for base, nm in (("basis_last", "по последней закупке"), ("basis_ss", "по цене СС")):
    g = C["by_group"][base]
    p(f"- {nm}: эталон по выданным на №1–№2 строкам {F(g['et_issued_lines'])} против группы «эталон» закупочной {F(g['zak']['еталон'])} "
      f"(Δ {F(g['delta_etalon_group_vs_issued_lines'])}); строки без выдачи на №1–№2 {F(g['et_lines_not_issued_1_2'])} → в закупочной 0; "
      f"«поза еталоном» +{F(g['zak']['поза еталоном'])}; МШП +{F(g['zak']['МШП'])}; итого Δ {F(g['delta_total'])}.")


def top_table(lst, title):
    p(f"\n## {title}")
    p("| Позиция (строка СС / карточка) | Группа | Карточка | Ед. | Факт №1 | Факт №2 | Среднее | Норма эталона | Склад проекта | Докупка | Закупочная, грн | Эталон, грн | Δ, грн |")
    p("|---|---|---|---|---|---|---|---|---|---|---|---|---|")
    for x in lst:
        cards = x["cards"] or [{"card": "—", "unit": "", "fact_1": None, "fact_2": None, "avg": None, "etalon_norm": None, "stock_project": None, "buy_qty": None}]
        for i, c in enumerate(cards):
            un = " ⚠" if c.get("unstable") else ""
            first = i == 0
            p(f"| {x['name'] if first else ''} | {x['group'] if first else ''} | {c['card']}{un} | {c['unit']} | {Q(c['fact_1'])} | {Q(c['fact_2'])} | {Q(c['avg'])} | "
              f"{Q(c['etalon_norm'])} | {Q(c['stock_project'])} | {Q(c['buy_qty'])} | {F(x['zak_uah_last']) if first else ''} | {F(x['et_uah_last']) if first else ''} | {('**' + F(x['delta_uah_last']) + '**') if first else ''} |")


top_table(C["top15_more_than_etalon"], "T4. Топ-15: закупочная требует БОЛЬШЕ эталона (по последней закупке)")
top_table(C["top15_less_than_etalon"], "T5. Топ-15: закупочная требует МЕНЬШЕ эталона (строки, выданные на №1–№2; по последней закупке)")

p("\n## T6. Строки эталона без выдачи на №1–№2 (в закупочной 0)")
p("| Строка СС | Докупка эталона №3–№7 | Ед. | По последней закупке | По цене СС |")
p("|---|---|---|---|---|")
s1 = s2 = 0
for x in sorted(C["etalon_lines_not_issued_on_1_2"], key=lambda x: -x["et_uah_ss"]):
    s1 += x["et_uah_last"]; s2 += x["et_uah_ss"]
    p(f"| {x['name']} | {Q(x.get('et_buy_line_units'))} | {x.get('et_unit', '')} | {F(x['et_uah_last'])} | {F(x['et_uah_ss'])} |")
p(f"| **Итого {len(C['etalon_lines_not_issued_on_1_2'])} строк** | | | **{F(s1)}** | **{F(s2)}** |")

p("\n## T7. Септик: детали по факту против состава эталона")
p("| Карточка | Деталь | Эталон на дом | №1: списано / собрано в п/ф / лежит | №2: списано / собрано / лежит | №3 | Склад проекта | Среднее | Докупка | Посл. закупка, грн | СС, грн |")
p("|---|---|---|---|---|---|---|---|---|---|---|")
for x in J["compare"]["septic"]["parts"]:
    p(f"| {x['card']} | {x['part']} | {x['etalon_per_house']} | {Q(x['cons_№1'])} / {Q(x['asm_№1'])} / {Q(x['stock_№1'])} | {Q(x['cons_№2'])} / {Q(x['asm_№2'])} / {Q(x['stock_№2'])} | "
      f"{Q(x['cons_№3'] + x['asm_№3'] + x['stock_№3'])} | {Q(x['stock_project'])} | {Q(x.get('avg'))} | {Q(x.get('buy'))} | {F(x.get('uah_last'))} | {F(x.get('uah_ss'))} |")
se = J["compare"]["septic"]
p(f"\nЗакупочная по септику: {F(se['zakupochnaya_uah_last'])} (СС {F(se['zakupochnaya_uah_ss'])}); эталон со связью: {Q(se['etalon_kit']['etalon_buy_kits'])} компл. = {F(se['etalon_kit']['etalon_uah_last'])} (СС {F(se['etalon_kit']['etalon_uah_ss'])}).")

p("\n## T8. Нестабильные карточки (CV > 0,3), топ по докупке")
u = C["unstable"]
p(f"Всего {u['count']} из {u['of_cards']}; докупка {F(u['buy_uah_last'])}; их среднее на дом {F(u['avg_house_uah_last'])}.\n")
p("| Карточка | Группа | Ед. | №1 | №2 | CV | Докупка | Грн |")
p("|---|---|---|---|---|---|---|---|")
for x in u["top"][:20]:
    p(f"| {x['card']} | {x['group']} | {x['unit']} | {Q(x['fact_1'])} | {Q(x['fact_2'])} | {F(x['cv'])} | {Q(x['buy_qty'])} | {F(x['uah_last'])} |")
s = C["sensitivity_two_houses"]
p(f"\nЧувствительность: минимум №1/№2 — {F(s['min_f1_f2_uah_last'])} (СС {F(s['min_f1_f2_uah_ss'])}); среднее — {F(s['avg_uah_last'])} (СС {F(s['avg_uah_ss'])}); максимум — {F(s['max_f1_f2_uah_last'])} (СС {F(s['max_f1_f2_uah_ss'])}).")

p("\n## T9. Строки закупочной: эталон, топ-30 по докупке")
rows = J["rows"]


def rows_table(rs):
    p("| Карточка | Ед. | Строка СС / способ | №1 | №2 | Среднее | Норма эт. | №3 сп.+леж. | Потребн. №3–№7 | Склад пр. | Докупка | Цена посл. | Грн посл. | Грн СС | №3 | №4 | №5 | №6 | №7 |")
    p("|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|")
    for r in rs:
        un = " ⚠" if r["unstable"] else ""
        bl = "" if r["price_basis_last"] == "остання закупівля" else " *"
        p(f"| {r['card']}{un} | {r['unit']} | {(r['ss_line'] or r['kind'])} / {r['method']} | {Q(r['fact_1'])} | {Q(r['fact_2'])} | {Q(r['avg'])} | {Q(r['etalon_norm'])} | "
          f"{Q(r['fact_3_spysano'] + r['lies_house_3'])} | {Q(r['need_3_7'])} | {Q(r['stock_project'])} | **{Q(r['buy_qty'])}** | {F(r['price_last'])}{bl} | {F(r['buy_uah_last'])} | {F(r['buy_uah_ss'])} | "
          + " | ".join(Q(r["by_house"][h]["buy"]) for h in H) + " |")


rows_table([r for r in rows if r["group"] == "еталон" and r["buy_qty"] > 0][:30])
p("\n## T10. Строки закупочной: «поза еталоном» (все с докупкой)")
rows_table([r for r in rows if r["group"] == "поза еталоном" and r["buy_qty"] > 0])
p("\nБез докупки (склад проекта закрывает): " + "; ".join(f"{r['card']} ({Q(r['avg'])} {r['unit']} на дом, склад {Q(r['stock_project'])})" for r in rows if r["group"] == "поза еталоном" and r["buy_qty"] == 0))
p("\n## T11. Строки закупочной: МШП, топ-12")
rows_table([r for r in rows if r["group"] == "МШП" and r["buy_qty"] > 0][:12])

p("\n## T12. Округление")
rd = T["rounding"]
p(f"Один раз на итог карточки: {F(rd['once_on_total_uah_last'])}; построчно по каждому дому: {F(rd['per_house_uah_last'])} (+{F(rd['delta_uah_last'])}, +{Q(rd['delta_pieces'])} шт.); без округления: {F(rd['no_rounding_uah_last'])}.\n")
p("| Карточка | Среднее на дом | Докупка (один раз) | Докупка (по домам) | +грн |")
p("|---|---|---|---|---|")
for x in rd["top"]:
    p(f"| {x['card']} | {Q(x['avg'])} | {Q(x['buy_once'])} | {Q(x['buy_per_house'])} | {F(x['delta_uah'])} |")

p("\n## T13. Основа цены")
for k, v in T["price_basis_last"].items():
    p(f"- {k}: {v['cards']} карт., {F(v['uah'])}")
fb = T["ss_fallback_vs_fact_cost"]
p(f"- без закупок по цене СС: {fb['cards']} карт. = {F(fb['uah_at_ss'])}; по себестоимости факта было бы {F(fb['uah_at_fact_cost'])}")

p("\n## T14. ЦСП и ОСБ на №2")
co = C["csp_osb_on_2"]
p("| Карточка | Лежит на №2 | Факт №1 | Списано №2 | Среднее (б) | Среднее (а) | Склад проекта | Докупка (б) | Грн (б) | Докупка (а), остаются на №2 | Грн | Докупка (а) + вернуть | Грн |")
p("|---|---|---|---|---|---|---|---|---|---|---|---|---|")
for x in co["cards"]:
    p(f"| {x['card']} | {Q(x['lies_on_2'])} | {Q(x['fact_1'])} | {Q(x['fact_2_written'])} | {Q(x['avg_b'])} | {Q(x['avg_a'])} | {Q(x['stock_project'])} | {Q(x['buy_b'])} | {F(x['uah_b'])} | "
      f"{Q(x['buy_avg_a_no_return'])} | {F(x['uah_avg_a_no_return'])} | {Q(x['buy_avg_a_return'])} | {F(x['uah_avg_a_return'])} |")
p(f"\nИтого докупка: (б) {F(co['total_b_main_uah_last'])}; для ЦСП/ОСБ среднее по (а), остаются на №2 — {F(co['total_avg_a_for_them_stay_on_2_uah_last'])}; по (а) и вернуть на склад проекта — {F(co['total_avg_a_for_them_return_to_project_uah_last'])}.")

p("\n## T15. №2 больше №1 / только на одном доме")
l2 = C["house_2_more_than_1"]
p(f"Карточек, где №2 больше №1 более чем на 500 грн: {l2['count']}, разница {F(l2['uah_last'])}. Только на №1: {l2['only_on_1_count']}; только на №2: {l2['only_on_2_count']}.\n")
p("| Карточка | Группа | Ед. | №1 | №2 | Разница, грн | Влияние на докупку (4 дома), грн |")
p("|---|---|---|---|---|---|---|")
for x in l2["top"][:13]:
    p(f"| {x['card']} | {x['group']} | {x['unit']} | {Q(x['fact_1'])} | {Q(x['fact_2'])} | {F(x['delta_uah_last'])} | {F(x['buy_effect_uah_last'])} |")
p("\nТолько на №1 (топ): " + "; ".join(f"{x['card']} {Q(x['fact_1'])} {x['unit']}" for x in l2["only_on_1"][:8]))
p("\nТолько на №2 (топ): " + "; ".join(f"{x['card']} {Q(x['fact_2'])} {x['unit']}" for x in l2["only_on_2"][:8]))

p("\n## T16. Сверка факта (б) с полем «Остаток» РК")
for h, v in J["houses_base"]["rk_reconciliation"].items():
    p(f"- {h}: карточек {v['cards']}, расхождений {v['diff_count']}" + (": " + "; ".join(f"{x['card']} {Q(x['mine'])} против {Q(x['rk_ostatok'])}" for x in v["diff"]) if v["diff"] else ""))

open(os.path.join(HERE, "z1_tables.md"), "w", encoding="utf-8").write("\n".join(OUT))
print("\n".join(OUT))
