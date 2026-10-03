# -*- coding: utf-8 -*-
"""Z3, часть 4: независимая проверка прототипа 15 м (раздел «Проверка» заметки Z3). ТОЛЬКО ЧТЕНИЕ BuhBud.
1) Свои COM-запросы (ВЫБРАТЬ): СС 000000008; факт по складам домов 15 м (Кт 20/22 со склада дома -> Дт не 20/22) и остатки на них;
   ПДВ закупки по карточкам; РК по СС (ТабличнаяЧастьОстатков, СписаниеСверхНормы); склад проекта.
2) Инварианты по модели строителя (z3_fact, в памяти) и по JSON.
Пишет в data/demo_specifikaciya_15m.json ключи «proverka» и zakupochnaya.vs_etalon_decomposition. Запускать ПОСЛЕ z3_report.py.
Каждый запрос предварительно прогнан тестом (Rule #-1); псевдонимы — не ключевые слова."""
import os, sys, json, collections
from z3_fact import *          # модель строителя (только память, ничего не пишет)
import win32com.client

J = json.load(open(OUT_JSON, encoding="utf-8"))
B = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
HOUSE_WH = 'С.Подразделение В ИЕРАРХИИ (ВЫБРАТЬ С2.ПодразделениеОбъекта ИЗ Справочник.СтруктураСебестоимости КАК С2 ГДЕ С2.Код = "000000008")'


def qq(text, **params):
    z = B.NewObject("Запрос"); z.Text = text
    for k, v in params.items():
        z.SetParameter(k, v)
    try:
        vt = z.Execute().Выгрузить()
    except Exception as e:
        raise RuntimeError(f"ЗАПРОС УПАЛ: {e.excepinfo[2] if getattr(e, 'excepinfo', None) else e}")
    cols = [vt.Колонки.Получить(i).Имя for i in range(vt.Колонки.Количество())]
    out = []
    for i in range(vt.Количество()):
        r = vt.Получить(i); d = {}
        for j, c in enumerate(cols):
            v = r.Получить(j)
            d[c] = v if (v is None or isinstance(v, (str, int, float, bool))) else B.String(v)
        out.append(d)
    return out


hn = lambda s: "15-" + str(s).strip()[-1]
P = {}

# ── 1. СС 000000008
ss = qq("""ВЫБРАТЬ КОЛИЧЕСТВО(К.НомерСтроки) КАК Строк, СУММА(К.Сумма) КАК Сум, СУММА(ВЫБОР КОГДА К.Количество <> 0 ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Ненулевых,
	СУММА(ВЫБОР КОГДА К.Номенклатура = ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК БезКарточки
ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К ГДЕ К.Ссылка.Код = "000000008" """)[0]
P["ss_008"] = {"rows": ss["Строк"], "nonzero": ss["Ненулевых"], "sum_vat": r2(ss["Сум"]), "rows_without_card": ss["БезКарточки"],
               "ok": ss["Строк"] == 190 and ss["Ненулевых"] == 168 and abs(Dd(ss["Сум"]) - Dd("704612.01")) < Decimal("0.005")}

# ── 2. Факт по складам домов (2025–2026): Кт 20/22 склада дома -> Дт не 20/22; остатки складов домов
TT = f"""ВЫБРАТЬ С.Ссылка КАК Склад ПОМЕСТИТЬ ВТ_Скл ИЗ Справочник.Склады КАК С ГДЕ {HOUSE_WH}
;
ВЫБРАТЬ П.Ссылка КАК Счет ПОМЕСТИТЬ ВТ_Сч ИЗ ПланСчетов.Хозрасчетный КАК П ГДЕ П.Код В ("20", "22")
;
"""
fact = qq(TT + """ВЫБРАТЬ ВЫРАЗИТЬ(Дв.СубконтоКт3 КАК Справочник.Склады).Подразделение.Наименование КАК Дом,
	ВЫРАЗИТЬ(Дв.СубконтоКт1 КАК Справочник.Номенклатура).Код КАК Код, ПРЕДСТАВЛЕНИЕ(ВЫРАЗИТЬ(Дв.СубконтоКт1 КАК Справочник.Номенклатура).СтавкаНДС) КАК Ставка,
	Дв.СчетКт.Код КАК СчКт, ТИПЗНАЧЕНИЯ(Дв.Регистратор) КАК ТипРег, Дв.Регистратор КАК Рег, СУММА(Дв.Сумма) КАК Сум, СУММА(Дв.КоличествоКт) КАК Кол
ИЗ РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(ДАТАВРЕМЯ(2025, 1, 1), ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59),
		СчетКт В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Сч.Счет ИЗ ВТ_Сч КАК ВТ_Сч), , ) КАК Дв
ГДЕ Дв.СубконтоКт3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл) И НЕ Дв.СчетДт В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Сч.Счет ИЗ ВТ_Сч КАК ВТ_Сч)
СГРУППИРОВАТЬ ПО ВЫРАЗИТЬ(Дв.СубконтоКт3 КАК Справочник.Склады).Подразделение.Наименование, ВЫРАЗИТЬ(Дв.СубконтоКт1 КАК Справочник.Номенклатура).Код,
	ПРЕДСТАВЛЕНИЕ(ВЫРАЗИТЬ(Дв.СубконтоКт1 КАК Справочник.Номенклатура).СтавкаНДС), Дв.СчетКт.Код, ТИПЗНАЧЕНИЯ(Дв.Регистратор), Дв.Регистратор""")
lying = qq(TT + """ВЫБРАТЬ ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Подразделение.Наименование КАК Дом, ВЫРАЗИТЬ(О.Субконто1 КАК Справочник.Номенклатура).Код КАК Код,
	ПРЕДСТАВЛЕНИЕ(О.Субконто3) КАК Склад, О.КоличествоОстатокДт КАК Кол, О.СуммаОстатокДт КАК Сум
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Сч.Счет ИЗ ВТ_Сч КАК ВТ_Сч), , ) КАК О
ГДЕ О.Субконто3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл)""")
codes = sorted({r["Код"] for r in fact} | {r["Код"] for r in lying})
arr = B.NewObject("Массив")
for c in codes:
    arr.Добавить(c)
ptu = qq("""ВЫБРАТЬ Т.Номенклатура.Код КАК Код, Т.Ссылка.СуммаВключаетНДС КАК Вкл, СУММА(Т.Сумма) КАК Сум, СУММА(Т.СуммаНДС) КАК НДС
ИЗ Документ.ПоступлениеТоваровУслуг.Товары КАК Т
ГДЕ Т.Ссылка.Проведен И Т.Ссылка.Дата МЕЖДУ ДАТАВРЕМЯ(2025, 1, 1) И ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59) И Т.Номенклатура.Код В (&Коды) И Т.Количество > 0
СГРУППИРОВАТЬ ПО Т.Номенклатура.Код, Т.Ссылка.СуммаВключаетНДС""", Коды=arr)
vat = collections.defaultdict(lambda: [Z, Z])
for r in ptu:
    s, n = Dd(r["Сум"]), Dd(r["НДС"]); a = vat[r["Код"]]; a[0] += s - (n if r["Вкл"] else Z); a[1] += n
rate = {c: (v[1] / v[0] if v[0] else None) for c, v in vat.items()}
CQ = collections.defaultdict(lambda: Z); CS = collections.defaultdict(lambda: Z)
H = {h: collections.defaultdict(lambda: Z) for h in H15}
docs = collections.defaultdict(set)
for r in fact:
    h = hn(r["Дом"]); s = Dd(r["Сум"]); a = H[h]
    a["net"] += s; a["kn" if r["СчКт"].startswith("20") else "pma"] += s
    rb = Decimal("0.2") if "20" in (r["Ставка"] or "") else Z
    a["gross_x12"] += s * VAT; a["gross_rate_card"] += s * (1 + rb)
    a["gross_rate_ptu"] += s * (1 + (rate.get(r["Код"]) if rate.get(r["Код"]) is not None else rb))
    CQ[(r["Код"], h)] += Dd(r["Кол"]); CS[(r["Код"], h)] += s
    docs[r["ТипРег"]].add(r["Рег"])
LYQ = collections.defaultdict(lambda: Z); LYS = collections.defaultdict(lambda: Z)
for r in lying:
    if Dd(r["Кол"]) or Dd(r["Сум"]):
        h = hn(r["Дом"]); LYQ[(r["Код"], h)] += Dd(r["Кол"]); LYS[(r["Код"], h)] += Dd(r["Сум"]); H[h]["lying"] += Dd(r["Сум"])
fh = {}
for h in H15:
    a = H[h]; jj = J["norma_fakt_by_house"][HNAME[h]]["variant"]
    fh[HNAME[h]] = {"spysano_net": r2(a["net"]), "kn_net": r2(a["kn"]), "pma_net": r2(a["pma"]), "lying_net": r2(a["lying"]),
                    "fakt_net": r2(a["net"] + a["lying"]), "fakt_x1_2": r2((a["net"] + a["lying"]) * VAT),
                    "spysano_z_pdv_stavka_kartky": r2(a["gross_rate_card"]), "spysano_z_pdv_stavka_zakupivli": r2(a["gross_rate_ptu"]),
                    "json_fact": jj["fact"], "ok": abs(a["net"] + a["lying"] - Dd(jj["fact"])) < Decimal("0.005")}
P["fact_by_house"] = fh
P["fact_totals"] = {"spysano_net": r2(sum(H[h]["net"] for h in H15)), "kn_net": r2(sum(H[h]["kn"] for h in H15)), "pma_net": r2(sum(H[h]["pma"] for h in H15)),
                    "lying_net": r2(sum(H[h]["lying"] for h in H15)), "kn_docs": len(docs.get("Комплектація номенклатури", ())),
                    "pma_docs": len(docs.get("Передача малоцінних активів в експлуатацію", ())),
                    "spysano_x1_2": r2(sum(H[h]["net"] for h in H15) * VAT), "spysano_z_pdv_stavka_zakupivli": r2(sum(H[h]["gross_rate_ptu"] for h in H15)),
                    "note": "з ПДВ ×1,2 — как РК (СуммаОстатка = нетто ×1,2); «ставка закупівлі» — ΣПДВ/Σнетто ПТУ 2025–2026 карточки (оценка: часть поставщиков без ПДВ)"}

# ── 3. Септик по домам (факт COM) и сверка с JSON
SEPT = ["00-00010526", "00000017651", "00-00005944", "00-00005945", "00-00006019", "00-00012879", "00-00017426"]
sep = {}
for h in H15:
    rows = {card_name(c): [r3(CQ[(c, h)]), r2(CS[(c, h)])] for c in SEPT if CQ[(c, h)]}
    js = sum(Dd(x["fact_s"]) for x in J["compare_now_vs_spec"]["septic"][HNAME[h]]["rows"])
    sep[HNAME[h]] = {"rows": rows, "sum_net": r2(sum(CS[(c, h)] for c in SEPT)), "json_sum": r2(js),
                     "ok": abs(sum(CS[(c, h)] for c in SEPT) - js) < Decimal("0.005")}
P["septic_by_house"] = sep

# ── 4. Закупочная: каждая карточка × дом = COM (списано + лежит); 10 крупнейших
ZR = {z["code"]: z for z in J["zakupochnaya"]["rows"]}
allc = {c for c, _ in CQ} | {c for c, _ in LYQ}
mism = []
for c in allc | set(ZR):
    z = ZR.get(c)
    for h in H15:
        f = CQ[(c, h)] + LYQ[(c, h)]
        if z is None or abs(f - Dd(z["by_house"][HNAME[h]])) > Decimal("0.0006"):
            mism.append([c, HNAME[h], r3(f), z["by_house"][HNAME[h]] if z else None])
top = sorted(allc, key=lambda c: -sum(CS[(c, h)] + LYS[(c, h)] for h in H15))[:10]
P["zakupochnaya_vs_com"] = {"cards_com": len(allc), "cards_json": len(ZR), "mismatch_card_house": mism[:20], "ok": not mism,
                            "top10": [{"code": c, "card": card_name(c), "by_house_com": [r3(CQ[(c, h)] + LYQ[(c, h)]) for h in H15],
                                       "fakt_net_com": r2(sum(CS[(c, h)] + LYS[(c, h)] for h in H15)), "json_fact_s_total": ZR[c]["fact_s_total"],
                                       "json_avg": ZR[c]["avg"], "price_last": ZR[c]["price_last"], "price_basis": ZR[c]["price_basis"],
                                       "node": ZR[c]["node"]} for c in top]}

# ── 5. РК по СС 000000008: суммы и построчно против факта
rk = qq("""ВЫБРАТЬ Т.Ссылка.Номер КАК РК, Т.Ссылка.Подразделение.Наименование КАК Дом, ПРЕДСТАВЛЕНИЕ(Т.Ссылка.Статус) КАК Статус,
	Т.Номенклатура.Код КАК Код, Т.Счет.Код КАК Счет, Т.Остаток КАК Ост, Т.СуммаОстатка КАК СумОст, Т.ВНормеСумма КАК ВНормеСум,
	Т.ПонадНормуСумма КАК ПонадСум, Т.ЕкономіяСума КАК ЭкСум
ИЗ Документ.РасчетКомплектаций.ТабличнаяЧастьОстатков КАК Т
ГДЕ Т.Ссылка.Спецификация.Код = "000000008" И НЕ Т.Ссылка.ПометкаУдаления""")
sv = qq("""ВЫБРАТЬ Т.Ссылка.Номер КАК РК, СУММА(Т.Сумма) КАК Сум ИЗ Документ.РасчетКомплектаций.СписаниеСверхНормы КАК Т
ГДЕ Т.Ссылка.Спецификация.Код = "000000008" И НЕ Т.Ссылка.ПометкаУдаления СГРУППИРОВАТЬ ПО Т.Ссылка.Номер""")
SVS = {r["РК"]: Dd(r["Сум"]) for r in sv}
R = collections.defaultdict(lambda: collections.defaultdict(lambda: Z)); RQ = collections.defaultdict(lambda: [Z, Z]); RKN = {}
for r in rk:
    h = hn(r["Дом"]); RKN[h] = (r["РК"], r["Статус"]); a = R[h]
    for k, f in (("ost", "СумОст"), ("vn", "ВНормеСум"), ("ponad", "ПонадСум"), ("ek", "ЭкСум")):
        a[k] += Dd(r[f])
    if (r["Счет"] or "").startswith("22"):
        a["ponad_22"] += Dd(r["ПонадСум"]); a["vn_22"] += Dd(r["ВНормеСум"])
    x = RQ[(r["Код"], h)]; x[0] += Dd(r["Ост"]); x[1] += Dd(r["СумОст"]) / VAT
rkout = {}
for h in H15:
    a = R[h]; jn = J["compare_now_vs_spec"]["now_rk"][HNAME[h]]
    diffs = []
    for (c, hh) in set(RQ) | {k for k in CQ}:
        if hh != h:
            continue
        f = (CQ[(c, h)], CS[(c, h)])
        if abs(f[0] - RQ[(c, h)][0]) > Decimal("0.0005") or abs(f[1] - RQ[(c, h)][1]) > Decimal("0.05"):
            diffs.append({"code": c, "card": card_name(c), "fakt_q": r3(f[0]), "fakt_s": r2(f[1]), "rk_q": r3(RQ[(c, h)][0]), "rk_s_net": r2(RQ[(c, h)][1])})
    rkout[HNAME[h]] = {"rk": RKN[h][0], "status": RKN[h][1], "ostatok_net": r2(a["ost"] / VAT), "v_normi_net": r2(a["vn"] / VAT),
                       "ponad_net": r2(a["ponad"] / VAT), "ekonomiya_net": r2(a["ek"] / VAT), "sverkh_normy": r2(SVS.get(RKN[h][0], Z)),
                       "ponad_mshp_22_net": r2(a["ponad_22"] / VAT), "v_normi_mshp_22_net": r2(a["vn_22"] / VAT),
                       "json_equal": all(abs(Dd(jn[k]) - Dd(r2(v))) < Decimal("0.005") for k, v in
                                         (("ostatok", a["ost"] / VAT), ("v_normi", a["vn"] / VAT), ("ponad", a["ponad"] / VAT), ("ekonomiya", a["ek"] / VAT))),
                       "rk_vs_spysano_by_card": diffs}
P["rk_now"] = rkout

# ── 6. Склад проекта (Подразделение = ПодразделениеЗакупки СС)
st_p = qq("""ВЫБРАТЬ С.Ссылка КАК Склад ПОМЕСТИТЬ ВТ_Скл ИЗ Справочник.Склады КАК С
ГДЕ С.Подразделение В (ВЫБРАТЬ СС.ПодразделениеЗакупки ИЗ Справочник.СтруктураСебестоимости КАК СС ГДЕ СС.Код = "000000008")
;
ВЫБРАТЬ П.Ссылка КАК Счет ПОМЕСТИТЬ ВТ_Сч ИЗ ПланСчетов.Хозрасчетный КАК П ГДЕ П.Код В ("20", "22")
;
ВЫБРАТЬ ВЫРАЗИТЬ(О.Субконто1 КАК Справочник.Номенклатура).Код КАК Код, О.КоличествоОстатокДт КАК Кол, О.СуммаОстатокДт КАК Сум
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Сч.Счет ИЗ ВТ_Сч КАК ВТ_Сч), , ) КАК О
ГДЕ О.Субконто3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл)""")
SP = collections.defaultdict(lambda: [Z, Z])
for r in st_p:
    SP[r["Код"]][0] += Dd(r["Кол"]); SP[r["Код"]][1] += Dd(r["Сум"])
smis = [c for c in ZR if abs(SP[c][0] - Dd(ZR[c]["stock_project_q"])) > Decimal("0.0005") or abs(SP[c][1] - Dd(ZR[c]["stock_project_s"])) > Decimal("0.011")]
P["project_stock"] = {"sum_net": r2(sum(v[1] for v in SP.values())), "cards": sum(1 for v in SP.values() if v[0] or v[1]),
                      "zakupochnaya_rows_mismatch": smis, "ok": not smis}

# ── 7. Инварианты по модели строителя
inv = collections.defaultdict(list)
for h in H15:
    for mode in ("etalon", "variant"):
        leaves, inst, recs, rem, norm = DETAIL[(h, mode)]
        fq = collections.defaultdict(lambda: Z); fs = collections.defaultdict(lambda: Z)
        for f in fact_rows(h, True):
            fq[f["code"]] += f["q"]; fs[f["code"]] += f["s"]
        rq = collections.defaultdict(lambda: Z); rs = collections.defaultdict(lambda: Z); inl = collections.defaultdict(lambda: Z)
        for r in recs:
            rq[r["card"]] += r["q"]; rs[r["card"]] += r["s"]
            if r["kind"] == "В нормі":
                inl[r["leaf"]] += r["q"]
        inv["I1_card_q_s"] += [(h, mode, c) for c in set(fq) | set(rq) if abs(fq[c] - rq[c]) > Decimal("1e-9") or abs(fs[c] - rs[c]) > Decimal("0.005")]
        inv["I8_in_le_norm"] += [(h, mode, k) for k, v in inl.items() if v - norm[k] > Decimal("1e-9")]
        j = J["norma_fakt_by_house"][HNAME[h]][mode]
        if abs(Dd(j["in"]) + Dd(j["over"]) + Dd(j["poza"]) + Dd(j["mshp"]) - Dd(j["fact"])) > Decimal("0.02"):
            inv["I1_json"].append((h, mode))
el, _ = explode("E")
byn = collections.defaultdict(lambda: Z)
for lf in el:
    byn[lf["n"] or SEPTIC_N] += lf["sum_n"] or Z
inv["I4_pos"] = [l["n"] for l in SS if abs(byn[l["n"]] - l["sum"]) > Decimal("0.001")]


def cmap(lv):
    m = collections.defaultdict(lambda: Z)
    for lf in lv:
        m[lf["card"] or ("ОН:" + K(lf["on"]) if lf["kind"].startswith("Група") else "без:" + str(lf["n"]))] += lf["q"]
    return m


em = cmap(el); var_chk = {}
for h in H15:
    vm = cmap(explode(house_spec(h))[0])
    changed = sorted(k for k in set(em) | set(vm) if abs(em[k] - vm[k]) > Decimal("1e-6"))
    var_chk[HNAME[h]] = {"edits": len(J["variants"]["by_house"][HNAME[h]]["edits"]), "cards_changed": len(changed),
                         "norm_vat": J["variants"]["by_house"][HNAME[h]]["norm_vat"]}
reg_chk = {}
for h in H15:
    for nid, node in J["registers_sim"]["by_house"][HNAME[h]]["nodes"].items():
        if "sostav" not in node:
            continue
        cards_u = {lf["card"] for lf in explode(house_spec(h))[0] if lf["card"] and (lf["path"].split("/")[-1].split("@")[0] == nid)}
        z_ch = sum(Dd(x["sum"]) for x in node["sostav"] if x["vid"] in ("В нормі", "Понад нормою"))
        f_u = sum(f["s"] for f in FACT if f["house"] == h and f["code"] in cards_u)
        reg_chk[f"{HNAME[h]} / {node['node']}"] = {"z_choho": r2(z_ch), "fakt_detaley": r2(f_u), "ok": abs(z_ch - f_u) < Decimal("0.03")}
zk_avg = [z["code"] for z in J["zakupochnaya"]["rows"] if abs(Dd(z["avg"]) * 6 - sum(Dd(v) for v in z["by_house"].values())) > Decimal("0.003")]
P["invariants"] = {
    "I1_fakt_eq_vnormi_ponad_poza_mshp": not (inv["I1_card_q_s"] or inv["I1_json"]),
    "I2_each_card_once": len({z["code"] for z in J["zakupochnaya"]["rows"]}) == len(J["zakupochnaya"]["rows"])
    and abs(sum(Dd(z["fact_s_total"]) for z in J["zakupochnaya"]["rows"]) - sum(Dd(J["norma_fakt_by_house"][HNAME[h]]["variant"]["fact"]) for h in H15)) < Decimal("0.01"),
    "I4_etalon_eq_ss": r2(sum(byn.values())) == r2(sum(l["sum"] for l in SS_ALL)) and not inv["I4_pos"],
    "I8_v_normi_le_norma": not inv["I8_in_le_norm"],
    "variant_eq_etalon_plus_edits": var_chk,
    "registers_z_choho_eq_fakt_detaley": reg_chk,
    "zakupochnaya_avg_x6_eq_sum": not zk_avg,
}

# ── 8. Разложение «средний дом − норма эталона» (по последней закупке, одинаковые неокруглённые цены)
never_last = Z
for x in NEVER:
    k = x["key"]
    if k.startswith("без"):
        never_last += ET_NORM_S[k]
    else:
        pl, _, _ = price_last(k); never_last += ET_NORM[k] * (pl if pl is not None else ETALON_PRICE_BY_CARD[k] / VAT)
et_avg = sum(Dd(z["avg_uah_last"]) for z in J["zakupochnaya"]["rows"] if z["group"].startswith("еталон"))
tg = J["zakupochnaya"]["totals_by_group"]
dec = {"etalon_norm_last": r2(ET_TOT_LAST), "never_issued_13_rows": r2(-never_last),
       "issued_etalon_rows": r2(et_avg - (ET_TOT_LAST - never_last)), "poza_etalonom": tg["поза еталоном"]["avg_uah_last"], "mshp": tg["МШП"]["avg_uah_last"],
       "avg_house_last": J["zakupochnaya"]["totals"]["avg_uah_last"],
       "note": "норма эталона по посл. закупке включает 10 строк без карточки по цене норми (19 931,67)"}
dec["delta"] = r2(Dd(dec["avg_house_last"]) - Dd(dec["etalon_norm_last"]))
J["zakupochnaya"]["vs_etalon_decomposition"] = dec
P["vs_etalon_decomposition"] = dec
P["stamp"] = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
J["proverka"] = P
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(J, f, ensure_ascii=False, indent=1, default=lambda o: float(o) if isinstance(o, Decimal) else str(o))
print(json.dumps({k: (v if k in ("ss_008", "fact_totals", "vs_etalon_decomposition", "project_stock") else "…") for k, v in P.items()},
                 ensure_ascii=False, indent=1, default=str))
print("fact ok:", all(v["ok"] for v in fh.values()), "| septic ok:", all(v["ok"] for v in sep.values()), "| zak ok:", P["zakupochnaya_vs_com"]["ok"],
      "| rk json:", all(v["json_equal"] for v in rkout.values()), "| rk≠факт:", {k: v["rk_vs_spysano_by_card"] for k, v in rkout.items() if v["rk_vs_spysano_by_card"]})
print("invariants:", {k: v for k, v in P["invariants"].items() if not isinstance(v, dict)},
      "| registers ok:", all(v["ok"] for v in reg_chk.values()), "| variants:", var_chk)
