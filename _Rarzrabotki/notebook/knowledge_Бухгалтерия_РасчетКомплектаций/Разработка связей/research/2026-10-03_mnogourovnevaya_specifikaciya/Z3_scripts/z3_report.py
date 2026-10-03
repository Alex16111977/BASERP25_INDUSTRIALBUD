# -*- coding: utf-8 -*-
"""Z3, часть 3: JSON data/demo_specifikaciya_15m.json и таблицы z3_tables.md для заметки Z3. Только чтение файлов."""
from z3_fact import *

OUT_MD = os.environ.get("Z3_MD", os.path.join(HERE, "z3_tables.md"))
F = lambda x: f"{x:,.2f}".replace(",", " ").replace(".", ",") if isinstance(x, (int, float)) and x is not None else ("—" if x is None else str(x))
Q = lambda x: (f"{x:,.3f}".rstrip("0").rstrip(",.") if isinstance(x, float) else str(x)).replace(",", " ").replace(".", ",") if x is not None else "—"
md = []


def T(title, head, rows):
    md.append(f"\n### {title}\n")
    md.append("| " + " | ".join(head) + " |")
    md.append("|" + "|".join("---" for _ in head) + "|")
    for r in rows:
        md.append("| " + " | ".join(str(v) for v in r) + " |")


def tree_lines(sid, depth=0, mult=Decimal(1), out=None):
    out = [] if out is None else out
    S = SPECS[sid]
    share = mult / S["out_q"]
    for ln in S["lines"]:
        qq = ln["q"] * share
        if ln["kind"] == "Вузол":
            V = sum((lf["sum_n"] or Z) for lf in ETALON_LEAVES if (ln["node"] in lf["path"].split("/")))
            out.append(("  " * depth) + f"├─ [Вузол] {SPECS[ln['node']]['name']} × {Q(r6(ln['q']))} {ln['unit']}"
                       + (f" (на будинок {Q(r6(qq))} {ln['unit']})" if share != 1 else "") + f" — норма {F(r2(V))} з ПДВ"
                       + (f"  ← поз. {ln['n']} «{ln['pos']}»" if ln["n"] else ""))
            tree_lines(ln["node"], depth + 1, qq, out)
        else:
            kind = "[Група] " if ln["kind"].startswith("Група") else ""
            c = f"{card_name(ln['card'])} ({ln['card']})" if ln["card"] else ("ОН «" + ln["on"] + "»" if kind else "— картки немає")
            kk = "" if ln["k"] == 1 else f", 1 {CARD.get(ln['card'], {}).get('unit', '')} = {Q(float(ln['k']))} {ln['unit']}"
            qtxt = (f"{r6(ln['q']):.6f}".replace(".", ",") + f" {ln['unit']} на 1 м²") if S["out_u"] == "м²" else f"{Q(r6(ln['q']))} {ln['unit']}"
            out.append(("  " * depth) + f"│  {kind}{c}: {qtxt}{kk}"
                       + (f"  ← поз. {ln['n']} «{ln['pos']}»" if ln["n"] else "  ← складова вузла"))
    return out


# ───────────────────────── эталон: строки и статистика
ROWS = []
for lf in ETALON_LEAVES:
    ln = lf["line"]
    ROWS.append({"node_path": " / ".join(SPECS[x]["name"] for x in lf["path"].split("/")) if lf["path"] else "", "kind": lf["kind"],
                 "pos_n": lf["n"] or SEPTIC_N, "pos": lf["pos"], "etap": lf["etap"], "pos_key": lf["pos_key"],
                 "card_code": lf["card"], "card": card_name(lf["card"]) if lf["card"] else None, "on": lf["on"],
                 "source": ln["src"], "auto_category": AUTO.get(lf["n"] or SEPTIC_N, {}).get("cat"),
                 "checked": ln["checked"], "q_line_per_out": r6(ln["q"]), "unit_line": lf["unit_line"], "q_line_house": r6(lf["q_line"]),
                 "k": r6(ln["k"]), "k_source": ln["ksrc"], "card_unit": CARD.get(lf["card"], {}).get("unit") if lf["card"] else None,
                 "q_card_house": r6(lf["q"]), "price_norm_vat": r6(ln["price"]) if ln["price"] is not None else None,
                 "sum_norm_vat": r2(lf["sum_n"]) if lf["sum_n"] is not None else None, "note": ln["note"]})
cat = collections.Counter(a["cat"] for a in AUTO.values())
cat_sum = collections.defaultdict(lambda: Z)
for n, a in AUTO.items():
    cat_sum[a["cat"]] += LINE[n]["sum"]
yellow_rej = [n for n, a in AUTO.items() if a.get("yellow_rejected")]
manual = [n for n, a in AUTO.items() if a["manual"]]
nocard_rows = sorted({lf["n"] for lf in ETALON_LEAVES if lf["kind"] == "Деталь" and not lf["card"]})
k_rows = [(lf["n"], lf["pos"], lf["unit_line"], card_name(lf["card"]), CARD.get(lf["card"], {}).get("unit"), float(lf["line"]["k"]), lf["line"]["ksrc"])
          for lf in ETALON_LEAVES if lf["card"] and (lf["line"]["k"] != 1 or not lf["line"]["ksrc"].startswith("Єдиниц"))]
area_err = []
for lf in ETALON_LEAVES:
    if lf["path"].endswith(("N02", "N03a")) and lf["n"]:
        dq = lf["q_line"] - LINE[lf["n"]]["q"] * (Decimal("0.5") if lf["n"] == 42 else 1)
        area_err.append({"pos": f"{lf['n']} {lf['pos']}", "q_ss": r6(LINE[lf['n']]["q"]), "q_per_m2": r6(lf["line"]["q"]),
                         "q_house": r6(lf["q_line"]), "dq": r6(dq)})
STATS = {
    "rows_ss": len(SS_ALL), "rows_nonzero": len(SS), "rows_zero_skipped": len(ZERO),
    "auto_fill": {k: {"rows": v, "sum_vat": r2(cat_sum[k])} for k, v in cat.items()},
    "green": cat["зелена: єдина картка ОН"], "yellow": cat["жовта: єдина списана на будинки типу"],
    "choose": sum(v for k, v in cat.items() if k.startswith("вибрати")), "group_rows": cat["Група"], "node_rows": cat["Вузол (позиція-комплект)"],
    "yellow_rejected_by_accountant": yellow_rej, "manual_choices_rows": len(manual), "rows_without_card_after_choice": nocard_rows,
    "lines_need_k": [{"pos_n": a, "pos": b, "unit_line": c, "card": d, "card_unit": e, "k": f, "k_source": g} for a, b, c, d, e, f, g in k_rows],
    "nodes_total": len(NODES), "nodes_top": sum(1 for v in NODES.values() if not v[1]), "nodes_nested": sum(1 for v in NODES.values() if v[1]),
    "area_nodes": [k for k, v in NODES.items() if v[3] == "м²"], "groups": 1, "group_lines": cat["Група"], "leaves": len(ETALON_LEAVES),
    "sum_norm_vat": r2(ETALON_SUM), "sum_ss_vat": r2(sum(l["sum"] for l in SS)), "I4_diff": r2(ETALON_SUM - sum(l["sum"] for l in SS)),
    "sum_norm_net": r2(ETALON_SUM / VAT), "price_protocol": PRICE_PROTOCOL, "area_rounding": area_err,
}
NODES_OUT = [{"id": it["node"], "name": it["name"], "path": it["path"], "K_on_house": r6(it["K"]), "unit": it["unit"], "level": it["level"],
              "pos": it["pos"], "V_norm_vat": r2(it["V"]), "lines": len(SPECS[it["node"]]["lines"]),
              "rows_ss": [r for r in ROWS_OF_NODE.get(it["node"], [])]} for it in ETALON_INST]

# ───────────────────────── варианты
VAR_OUT = {}
for h in H15:
    v = VARIANTS.get(h)
    if not v:
        continue
    lv, _ = explode(v["spec"])
    VAR_OUT[HNAME[h]] = {"spec": SPECS[v["spec"]]["name"], "edits": v["edits"], "copied_nodes": v["copied_nodes"],
                         "norm_vat": r2(sum((lf["sum_n"] or Z) for lf in lv)), "norm_net": r2(sum((lf["sum_n"] or Z) for lf in lv) / VAT),
                         "delta_vs_etalon_vat": r2(sum((lf["sum_n"] or Z) for lf in lv) - ETALON_SUM)}
rules_houses = {"V1": sorted({h for h, v in VARIANTS.items() if any(e["rule"].startswith("V1") for e in v["edits"])}),
                "V2": sorted({h for h, v in VARIANTS.items() if any(e["rule"].startswith("V2") for e in v["edits"])}),
                "V3": sorted({h for h, v in VARIANTS.items() if any(e["rule"].startswith("V3") for e in v["edits"])})}
VAR_SUMMARY = {"rules": {"V1": "конструктивний варіант вузла (септик) — задано пользователем/R2",
                         "V2": "заміна аналогом 1:1 — картка норми на будинок не списувалась, списана інша картка тієї ж ОН (не з еталону), прораб підтвердив",
                         "V3": "«різна комплектація будинків» за ПричиниЦеху — позиція є не на всіх будинках або кількість різна"},
               "houses_by_rule": {k: [HNAME[h] for h in v] for k, v in rules_houses.items()},
               "houses_without_variant": {"тільки V1": [HNAME[h] for h in H15 if h not in rules_houses["V1"]],
                                          "V1+V2": [HNAME[h] for h in H15 if h not in set(rules_houses["V1"]) | set(rules_houses["V2"])],
                                          "V1+V2+V3": [HNAME[h] for h in H15 if h not in set(rules_houses["V1"]) | set(rules_houses["V2"]) | set(rules_houses["V3"])]},
               "v2_rejected_by_builder": [{"house": HNAME[h], "card": card_name(c), "proposed": card_name(b), "why": w} for h, c, b, w in V2_REJECTED]}

# ───────────────────────── норма против факта (по карточкам — вариант)
NF_OUT = {}
for h in H15:
    cv = card_view(h, "variant")
    cards = []
    for key, d in cv.items():
        nm = card_name(key) if not key.startswith(("ОН:", "без")) else key
        cards.append({"key": key, "name": nm, **{k: (r6(v) if k.endswith("_q") else r2(v)) for k, v in d.items()}})
    cards.sort(key=lambda x: -(x.get("over_s", 0) + x.get("poza_s", 0)))
    NF_OUT[HNAME[h]] = {"spec": SPECS[house_spec(h)]["name"], "now_rk": NOW[h], "etalon": NF[h]["etalon"], "variant": NF[h]["variant"],
                        "variant_pma_in_norm_fix": NF_PMA_FIX[h],
                        "variant_effect": {k: r2(Dd(NF[h]["variant"].get(k, 0)) - Dd(NF[h]["etalon"].get(k, 0))) for k in ("in", "over", "poza", "econ_np", "norm")},
                        "cards": cards}

# ───────────────────────── JSON
CAVEATS = [
    "Режим только чтения: BuhBud читалась COM-запросами (ВЫБРАТЬ), в 1С ничего не записано и не проведено. Спецификации, варианты, РК-снимки и регистры существуют только в памяти Python.",
    f"Снимок {RAW['stamp']}: 4 800 движений 20/22/26/28 по складам и партиям IRC (2026), 503 строки ПТУ проекта, 11 205 строк ПТУ 2025–2026 по карточкам. Склад проекта «ІБ _МД IRC 2026» — {F(r2(sum(st[c]['P'][1] for c in st)))} без ПДВ.",
    "Служебные КН сборки септика 30 м №1 ІБ00-000174 и ІБ00-000175 на момент снимка НЕ проведены (утром 03.10 были проведены): состав септика взят из ТЧ документа ІБ00-000174, а не из проводок. Для 15 м это не влияет — полуфабрикатов на 15 м нет.",
    "Выбор карточек для строк «вибрати» и отклонение жёлтых предложений — СИМУЛЯЦИЯ бухгалтера (прораб подсказывает назначение позиции). Это не решение бухгалтера; в механизме каждую строку надо проверить («Перевірено»).",
    "Факт без ПДВ (Хозрасчетный), норма СС з ПДВ; в сравнениях норма ÷1,2. «Вид сейчас» — поля РК: ТабличнаяЧастьОстатков (з ПДВ → ÷1,2) и СписаниеСверхНормы (без ПДВ).",
    "Генератор 15 м №4 (30 768,30 без ПДВ) лежит на складе «МД IRС 2026 15 м №4 Відвантажено» (перемещение ІБ00-002273 от 16.09.2026), не списан; РК №6 «Цикл закрито» его не видит. В норме против факта он учтён как «лежит на складе дома», в регистрах (только списания) — нет.",
    "Строки «Підлога (на 1 м²)» и «Внутрішня обшивка стін (на 1 м²)» хранят количество на 1 м² с 6 знаками: норма дома отличается от СС на миллионные доли (ЦСП 5,000004 шт, плінтус 14,000004 пог.м). Деньги нормы не страдают (цена норми = сумма позиции ÷ количество листа).",
    "Закупочная спецификация: среднее по 6 домам (дом без карточки даёт 0), факт = списано + лежит на складе дома. «Докупка на 1/6 домов» — иллюстрация: склад проекта общий с 30 м и уже частично нужен домам 30 м №3–№7 (Z1).",
    "Цена последней закупки — ПТУ проекта, иначе любая ПТУ 2025–2026, иначе себестоимость факта; цена ≠ себестоимости ×5 — берётся предыдущая закупка (защита от ошибки единицы, Z1 §1.4).",
    "Варианты домов: все 6 домов получили вариант по правилам V1–V3; если не применять V3 (подтверждение цеха «різна комплектація»), по эталону идут №3, №4, №6 при V1 и ни один — при V1+V2.",
]
OUT = {
    "snapshot": {"stamp": RAW["stamp"], "base": 'Srvr="localhost";Ref="bas_industrialbud" (только чтение)', "moves": len(MV),
                 "moves_max_period": max(r["dt"] for r in MV).strftime("%d.%m.%Y %H:%M:%S"), "receipts_project_lines": len(RAW["receipts"]),
                 "receipts_any_lines": len(RAW["receipts_any"]), "project_stock_net": r2(sum(st[c]["P"][1] for c in st)),
                 "project_stock_cards": sum(1 for c in st if st[c]["P"][0] > 0),
                 "fact_15m_kn": r2(sum(f["s"] for f in FACT if f["type"].startswith("Комплектація"))),
                 "fact_15m_kn_docs": len({f["doc"] for f in FACT if f["type"].startswith("Комплектація")}),
                 "fact_15m_pma": r2(sum(f["s"] for f in FACT if f["type"].startswith("Передача"))),
                 "fact_15m_pma_docs": len({f["doc"] for f in FACT if f["type"].startswith("Передача")}),
                 "kn_septic_30m": RAW["kn_septic"], "scripts": "Z3_scripts/z3_collect.py, base15.py, z3_spec.py, z3_fact.py, z3_report.py"},
    "card": {**RAW["ss_head"][0], "rows": len(SS_ALL), "rows_nonzero": len(SS), "rows_zero": len(ZERO),
             "sum_vat": r2(sum(l["sum"] for l in SS)), "sum_net": r2(sum(l["sum"] for l in SS) / VAT),
             "zero_rows": [{"n": l["n"], "etap": l["etap"], "pos": l["pos"], "on": l["on"],
                            "fact_avg": r3(sum(sum(FQ[c][h] for h in H15) for c in ONC.get(l["on_code"], []) if c not in ETALON_CARDS) / 6),
                            "houses_with_fact": sum(1 for h in H15 if any(FQ[c][h] for c in ONC.get(l["on_code"], []) if c not in ETALON_CARDS)),
                            "cards_fact": sorted({card_name(c) for c in ONC.get(l["on_code"], []) if c not in ETALON_CARDS and any(FQ[c][h] for h in H15)}),
                            "fact_s_avg": r2(sum(sum(FS[c][h] for h in H15) for c in ONC.get(l["on_code"], []) if c not in ETALON_CARDS) / 6)} for l in ZERO]},
    "houses": [{"house": HNAME[h], "code": next((x["Код"] for x in RAW["houses"] if house_of(x["Имя"]) == h), None),
                "warehouses": [{"name": w["Имя"], "kind": w["Вид"]} for w in RAW["warehouses"] if house_of((w["Подр"] or "").strip()) == h],
                "rk": RK_OF.get(h), "kn": [r for r in RAW["rk_kn"] if house_of(r["Дом"]) == h],
                "cycle": [r for r in RAW["rk_cycle"] if house_of(r["Дом"]) == h],
                "fact_kn": r2(sum(f["s"] for f in FACT if f["house"] == h and f["type"].startswith("Комплектація"))),
                "fact_pma": r2(sum(f["s"] for f in FACT if f["house"] == h and f["type"].startswith("Передача"))),
                "lying": r2(sum(LS[c][h] for c in LS)), "spec": SPECS[house_spec(h)]["name"]} for h in H15],
    "etalon": {"rows": ROWS, "nodes": NODES_OUT, "stats": STATS, "disputed": [{"row": r, "node": SPECS[n]["name"], "why": w} for r, n, w in DISPUTED],
               "tree_text": tree_lines("E")},
    "variants": {"summary": VAR_SUMMARY, "by_house": VAR_OUT},
    "norma_fakt_by_house": NF_OUT,
    "registers_sim": {"nodes": [SPECS[n]["name"] for n in REG_NODES], "by_house": {HNAME[h]: REG[h] for h in H15},
                      "invariants": {**{HNAME[h]: INV[h] for h in H15}, "I4_etalon_vs_ss": r2(ETALON_SUM - sum(l["sum"] for l in SS))}},
    "zakupochnaya": {"kind": "Закупівельна (третій вид А_СпецификацииДомов), плоска по картках, статус «Діє»", "rows": ZAK,
                     "totals_by_group": {g: {k: r2(v) for k, v in d.items()} for g, d in TOT.items()}, "totals": TOT_ALL,
                     "etalon_norm_net_normprice": r2(ET_TOT_NP), "etalon_norm_net_lastprice": r2(ET_TOT_LAST),
                     "etalon_lines_never_issued": NEVER, "etalon_lines_never_issued_np": r2(NEVER_NP),
                     "by_node": {k: {"avg_last": r2(v[0]), "buy6_last": r2(v[1]), "cards": v[2]} for k, v in BY_NODE.items()},
                     "top_over": TOP_OVER, "top_under": TOP_UNDER, "compare_all": cmp_list},
    "compare_now_vs_spec": {"now_rk": {HNAME[h]: NOW[h] for h in H15}, "transitions": {HNAME[h]: TRANS[h] for h in H15},
                            "septic": {HNAME[h]: SEPTIC[h] for h in H15}},
    "caveats": CAVEATS,
}


def _j(o):
    if isinstance(o, Decimal):
        return float(o)
    if isinstance(o, datetime.datetime):
        return o.strftime("%d.%m.%Y %H:%M:%S")
    return str(o)


with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(OUT, f, ensure_ascii=False, indent=1, default=_j)

# ───────────────────────── таблицы
md.append(f"# Z3 таблицы (снимок {RAW['stamp']})")
md.append("\n## Дерево эталона\n```")
md.extend(tree_lines("E"))
md.append("```")
T("T1. «Заповнити з СС» — что предлагает база", ["Категория", "Позиций", "Норма з ПДВ"],
  [[k, v, F(r2(cat_sum[k]))] for k, v in sorted(cat.items(), key=lambda kv: -kv[1])] + [["**Итого**", sum(cat.values()), F(r2(sum(cat_sum.values())))]])
T("T2. Узлы эталона", ["Узел (путь)", "Кол-во на дом", "Строк", "Норма з ПДВ"],
  [[" / ".join(SPECS[x]["name"] for x in it["path"].split("/")), f"{Q(r6(it['K']))} {it['unit']}", len(SPECS[it['node']]['lines']), F(r2(it["V"]))] for it in ETALON_INST])
T("T3. Строки с коэффициентом / ручной единицей", ["Поз.", "Позиция", "Ед. строки", "Карточка", "Ед. карточки", "k", "Источник k"],
  [[a, b, c, d, e, Q(f), g] for a, b, c, d, e, f, g in k_rows])
Q6 = lambda x: f"{x:.6f}".replace(".", ",")
T("T4. Площадные узлы: округление 6 знаков", ["Позиция", "СС", "на 1 м²", "на дом", "Δ"],
  [[x["pos"], Q(x["q_ss"]), Q6(x["q_per_m2"]), Q6(x["q_house"]), Q6(x["dq"])] for x in area_err])
rows = []
for h in H15:
    v = VARIANTS.get(h)
    for e in v["edits"]:
        rows.append([HNAME[h], e["rule"].split(" ")[0], e["where"], e["pos"], e["was"], e["now"], e["why"]])
T("T5. Варианты домов: правки", ["Дом", "Правило", "Где", "Позиция СС", "Было", "Стало", "Почему"], rows)
T("T5б. Скопированные узлы", ["Дом", "Копии узлов (цепочка родителей)"],
  [[HNAME[h], ", ".join(c["name"] for c in VARIANTS[h]["copied_nodes"])] for h in H15])
T("T6. Норма против факта по домам (без ПДВ)", ["Дом", "РК (статус)", "Факт: списано+лежит", "Сейчас РК: понад", "Сейчас РК: економія",
                                              "Спец.: в нормі", "понад", "поза спец.", "МШП", "економія (ц. норми)", "в т.ч. норма без картки"],
  [[HNAME[h], f"№{int(NOW[h]['rk'])} ({NOW[h]['rk_status']})", F(NF[h]["variant"]["fact"]), F(NOW[h]["ponad"]), F(NOW[h]["ekonomiya"]),
    F(NF[h]["variant"]["in"]), F(NF[h]["variant"].get("over", 0)), F(NF[h]["variant"].get("poza", 0)), F(NF[h]["variant"].get("mshp", 0)),
    F(NF[h]["variant"]["econ_np"]), F(NF[h]["variant"]["econ_nocard"])] for h in H15])
T("T6б. Эффект вариантов (вариант − эталон), без ПДВ", ["Дом", "Норма", "в нормі", "понад", "поза спец.", "економія"],
  [[HNAME[h]] + [F(r2(Dd(NF[h]["variant"].get(k, 0)) - Dd(NF[h]["etalon"].get(k, 0)))) for k in ("norm", "in", "over", "poza", "econ_np")] for h in H15])
T("T7. Сейчас → по спецификации: переходы «понад» (без ПДВ)", ["Дом", "Хибне понад зникло", "Понад зменшилось", "Понад збільшилось", "Справжнє з'явилось"],
  [[HNAME[h], F(TRANS[h]["sum_gone"]), F(TRANS[h]["sum_decreased"]), F(TRANS[h]["sum_increased"]), F(TRANS[h]["sum_appeared"])] for h in H15])
for h in ("15-1",):
    T(f"T7а. {HNAME[h]}: хибне «понад» зникло (топ-12)", ["Карточка", "РК понад, кол", "РК понад, грн", "Спец., кол"],
      [[x["name"][:50], Q(x["now_ponad_q"]), F(x["now_ponad_s"]), Q(x["spec_ponad_poza_q"])] for x in TRANS[h]["false_ponad_gone"][:12]])
    T(f"T7б. {HNAME[h]}: справжнє «понад/поза» з'явилось (топ-12)", ["Карточка", "Спец., кол", "Спец., грн", "Вид"],
      [[x["name"][:50], Q(x["spec_ponad_poza_q"]), F(x["spec_ponad_poza_s"]), x["spec_kind"]] for x in TRANS[h]["real_appeared"][:12]])
T("T8. Септик: сейчас (РК) / эталон / вариант (без ПДВ)", ["Дом", "РК: понад", "Эталон: в нормі", "Эталон: поза", "Эталон: економія", "Вариант: в нормі", "Вариант: поза", "Вариант: економія"],
  [[HNAME[h], F(SEPTIC[h]["now"]["ponad_s"]), F(SEPTIC[h]["etalon"]["in_s"]), F(SEPTIC[h]["etalon"]["poza_s"]), F(SEPTIC[h]["etalon"]["econ_s"]),
    F(SEPTIC[h]["variant"]["in_s"]), F(SEPTIC[h]["variant"]["poza_s"]), F(SEPTIC[h]["variant"]["econ_s"])] for h in H15])
for nid in REG_NODES:
    rows = []
    for h in H15:
        d = REG[h]["nodes"][nid]
        if "uzly" not in d:
            rows.append([HNAME[h], "—", "—", "—", "—", d["note"]]); continue
        nf = [r for r in d["sostav"] if r["vid"] == "Норма"]
        fv = collections.defaultdict(float)
        for r in d["sostav"]:
            if r["vid"] != "Норма":
                fv[r["vid"].split(" (")[0]] += r["sum"]
        rows.append([HNAME[h], F(sum(r["sum"] for r in nf)), "; ".join(f"{k} {F(round(v, 2))}" for k, v in fv.items()),
                     f"{F(d['sobrano_pct'])} %", "; ".join(f"{u['op']} {F(u['pct'])} % ({u['stage']})" for u in d["uzly"] if u["rec"] == "Расход"),
                     f"{Q(d['zalyshok_total'])}"])
    T(f"T9. Регистры: {SPECS[nid]['name']}", ["Дом", "Норма, без ПДВ", "Факт «з чого»", "Зібрано", "Вибуло", "Залишок"], rows)
T("T10. Инварианты §6.11", ["Дом", "И2 регистры = Хозрасчетный", "И3", "И5", "И8", "И9 Септик", "И9 Водостік"],
  [[HNAME[h], f"{F(INV[h]['I2_fact_registers'])} = {F(INV[h]['I2_fact_hozraschetny'])} {'✓' if INV[h]['I2_ok'] else '✗'}",
    "✓" if INV[h]["I3_ok"] else "✗", "✓", "✓" if INV[h]["I8_ok"] else "✗", "✓" if INV[h]["I9_N12a"] else "✗", "✓" if INV[h]["I9_N05a"] else "✗"] for h in H15])
T("T11. Закупочная: итоги по группам (на 1 средний дом и докупка; без ПДВ)", ["Группа", "Карточек", "Средний дом, посл. закупка", "Средний дом, ц. норми*",
                                                                           "Докупка на 1 дом", "Докупка на 6 домов"],
  [[g, int(d["cards"]), F(r2(d["avg_uah_last"])), F(r2(d["avg_uah_norm"])), F(r2(d["buy_1_house_uah"])), F(r2(d["buy_6_houses_uah"]))] for g, d in TOT.items()]
  + [["**Всего**", int(TOT_ALL["cards"]), F(TOT_ALL["avg_uah_last"]), F(TOT_ALL["avg_uah_norm"]), F(TOT_ALL["buy_1_house_uah"]), F(TOT_ALL["buy_6_houses_uah"])]])
T("T12. Топ-15: тип систематически тратит БОЛЬШЕ эталона (без ПДВ)", ["Карточка", "Группа", "Ед.", "Среднее", "Норма эт.", "Δ посл. закупка", "Δ ц. норми"],
  [[x["name"][:45], x["group"], x["unit"], Q(x["avg_q"]), Q(x["norm_q"]), F(x["d_last"]), F(x["d_np"])] for x in TOP_OVER])
T("T13. Топ-15: тип тратит МЕНЬШЕ эталона (без ПДВ)", ["Карточка", "Ед.", "Среднее", "Норма эт.", "Δ посл. закупка", "Δ ц. норми"],
  [[x["name"][:45], x["unit"], Q(x["avg_q"]), Q(x["norm_q"]), F(x["d_last"]), F(x["d_np"])] for x in TOP_UNDER])
T("T14. Строки эталона без единой выдачи на 6 домах", ["Позиция", "Карточка", "Норма", "Ед.", "Норма, ц. норми без ПДВ", "Узел"],
  [[x["pos"], x["card"], Q(x["norm_q"]), x["unit"], F(x["norm_np"]), x["node"]] for x in NEVER])
T("T15. Закупочная по узлам прораба (средний дом, посл. закупка, без ПДВ)", ["Узел", "Карточек", "Средний дом", "Докупка на 6 домов"],
  [[k, v[2], F(r2(v[0])), F(r2(v[1]))] for k, v in sorted(BY_NODE.items(), key=lambda kv: -kv[1][0])])
T("T16. Закупочная: топ-25 по докупке на 6 домов", ["Карточка", "Группа", "Узел", "Ед.", "Среднее", "Склад пр.", "Докупка 6", "Цена посл.", "Грн"],
  [[x["card"][:40], x["group"], x["node"][:40], x["unit"], Q(x["avg"]), Q(x["stock_project_q"]), Q(x["buy_6_houses_q"]), F(x["price_last"]), F(x["buy_6_houses_uah"])]
   for x in sorted(ZAK, key=lambda x: -x["buy_6_houses_uah"])[:25]])
T("T17. Нулевые строки СС: факт карточек их ОН, которых нет в эталоне", ["Поз.", "Позиция", "ОН", "Домов с фактом", "Среднее кол-во", "Средн. грн/дом", "Карточки"],
  [[z["n"], z["pos"], z["on"], z["houses_with_fact"], Q(z["fact_avg"]), F(z["fact_s_avg"]), "; ".join(z["cards_fact"])[:90]] for z in OUT["card"]["zero_rows"]])
T("T18. Сводка по домам: сейчас (РК) против спецификации (вариант), без ПДВ",
  ["Дом", "Факт", "РК понад", "Спец. понад", "Спец. поза спец.", "Спец. МШП", "РК економія", "Спец. економія (рядки з карткою)", "Норма без картки", "Хибне понад зникло: → в нормі / → МШП"],
  [[HNAME[h], F(NF[h]["variant"]["fact"]), F(NOW[h]["ponad"]), F(NF[h]["variant"].get("over", 0)), F(NF[h]["variant"].get("poza", 0)), F(NF[h]["variant"].get("mshp", 0)),
    F(NOW[h]["ekonomiya"]), F(r2(Dd(NF[h]["variant"]["econ_np"]) - Dd(NF[h]["variant"]["econ_nocard"]))), F(NF[h]["variant"]["econ_nocard"]),
    f"{F(TRANS[h]['sum_gone_to_norm'])} / {F(TRANS[h]['sum_gone_to_mshp'])}"] for h in H15])
T("T19. Правка «ПМА картки з норми — не МШП»: вариант, без ПДВ", ["Дом", "в нормі M5 → правка", "МШП M5 → правка", "економія M5 → правка"],
  [[HNAME[h], f"{F(NF[h]['variant']['in'])} → {F(NF_PMA_FIX[h]['in'])}", f"{F(NF[h]['variant']['mshp'])} → {F(NF_PMA_FIX[h]['mshp'])}",
    f"{F(NF[h]['variant']['econ_np'])} → {F(NF_PMA_FIX[h]['econ_np'])}"] for h in H15])
open(OUT_MD, "w", encoding="utf-8").write("\n".join(md))
print("JSON ->", OUT_JSON, "| MD ->", OUT_MD)
print("stats:", {k: v for k, v in STATS.items() if k not in ("lines_need_k", "area_rounding", "auto_fill")})
