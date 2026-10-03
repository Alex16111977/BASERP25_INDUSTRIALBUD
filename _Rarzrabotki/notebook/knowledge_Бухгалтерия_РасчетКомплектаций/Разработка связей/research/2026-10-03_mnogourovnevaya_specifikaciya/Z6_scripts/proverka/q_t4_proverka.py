# -*- coding: utf-8 -*-
"""Z6, независимая проверка расчёта закупок проекта «МД IRC 2026». ТОЛЬКО ЧТЕНИЕ BuhBud (ВЫБРАТЬ), ERP не трогается.

Факт берётся другими виртуальными таблицами, чем в z6_collect.py (там ДвиженияССубконто):
  ОборотыДтКт(…, Регистратор) — закуплено / прочие приходы / списано / прочие расходы по карточке;
  Обороты(…) с КорСчетом      — третий путь для итогов закуплено и списано;
  Остатки()                   — склад проекта и склады домов сейчас;
  ТЧ «Товары» ПТУ              — закупки по документам (и мимо склада проекта);
  ОборотыДтКт по партии ПТУ   — куда ушли закупки мимо склада проекта.
Классы складов определены заново (склад проекта — имена класса P, дом — «… 15|30 м №N …» с «IR»).
Затем: сверка с data/proekt_irc2026_zakupki.json по каждой карточке, инварианты, пересчёт z6_calc.py на raw.json (воспроизводимость).
Выход: q_t4_out.json рядом со скриптом (z6_calc.py переносит его в JSON как блок verification).
Запуск: python q_t4_proverka.py   (Z6_RAW_DIR — папка raw.json расчёта; Z6_VER_DIR — временная папка проверки)."""
import os, sys, re, json, math, datetime, collections, subprocess
from decimal import Decimal as D
if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
Z6S = os.path.dirname(HERE)
sys.path.insert(0, Z6S)
from h import connect, q            # noqa: E402

ROOT = os.path.abspath(os.path.join(Z6S, "..", "..", ".."))
JPATH = os.path.join(ROOT, "data", "proekt_irc2026_zakupki.json")
RAWDIR = os.environ.get("Z6_RAW_DIR", r"C:/Users/SUPPOR~1/AppData/Local/Temp/claude/z6_tmp/calc")
VERDIR = os.environ.get("Z6_VER_DIR", r"C:/Users/SUPPOR~1/AppData/Local/Temp/claude/z6_tmp/verify")
os.makedirs(VERDIR, exist_ok=True)
Z = D(0)
d = lambda x: D(str(x if x not in (None, "") else 0))
r2 = lambda x: float(d(x).quantize(D("0.01")))
P = {"ІБ _МД IRC 2026", "ІС _МД IRC 2026", "МД IRС 2026", "Цех МД IRS 2026", "МД IRS 2026"}
norm = lambda s: " ".join((s or "").replace("\xa0", " ").split())
mat = lambda a: (a or "")[:2] in ("20", "22", "28")


def wc(w):
    w = norm(w)
    if not w:
        return "-"
    if w in P:
        return "P"
    m = re.search(r"(15|30) м №\s*(\d+)", w)
    if m and ("IR" in w or "ІR" in w):
        return f"H{m.group(1)}-{m.group(2)}"
    return "IRC?" if "IR" in w else "X"


insys = lambda c: c == "P" or c.startswith("H")
HN = {f"H{t}-{n}": f"{t} м №{n}" for t in (15, 30) for n in range(1, 8)}

# ═════════════ 1. COM-запросы (только ВЫБРАТЬ)
B = connect("buh")
WH = """
ВЫБРАТЬ С.Ссылка КАК Склад
ПОМЕСТИТЬ ВТ_Скл
ИЗ Справочник.Склады КАК С
ГДЕ С.Наименование ПОДОБНО "%IR%" ИЛИ С.Наименование ПОДОБНО "%ІR%" ИЛИ С.Наименование ПОДОБНО "%RC 2026%" ИЛИ С.Наименование ПОДОБНО "%RС 2026%"
	ИЛИ С.Подразделение.Наименование ПОДОБНО "%IR%" ИЛИ С.Подразделение.Наименование ПОДОБНО "%RС%2026%"
;
ВЫБРАТЬ П.Ссылка КАК Счет ПОМЕСТИТЬ ВТ_Мат ИЗ ПланСчетов.Хозрасчетный КАК П ГДЕ П.Код В ("20", "22", "26", "28")
;
"""
MATQ = "В ИЕРАРХИИ (ВЫБРАТЬ ВТ_Мат.Счет ИЗ ВТ_Мат КАК ВТ_Мат)"
SEL = """
ВЫБРАТЬ
	ПРЕДСТАВЛЕНИЕ(О.Регистратор) КАК Рег, О.СчетДт.Код КАК СчДт, О.СчетКт.Код КАК СчКт,
	ВЫРАЗИТЬ(О.СубконтоДт1 КАК Справочник.Номенклатура).Код КАК КодДт, ВЫРАЗИТЬ(О.СубконтоКт1 КАК Справочник.Номенклатура).Код КАК КодКт,
	ВЫРАЗИТЬ(О.СубконтоДт3 КАК Справочник.Склады).Наименование КАК СклДт, ВЫРАЗИТЬ(О.СубконтоКт3 КАК Справочник.Склады).Наименование КАК СклКт,
	О.Организация.Наименование КАК Орг, О.СуммаОборот КАК Сумма, О.КоличествоОборотДт КАК КолДт, О.КоличествоОборотКт КАК КолКт
ИЗ РегистрБухгалтерии.Хозрасчетный.ОборотыДтКт(ДАТАВРЕМЯ(2026, 1, 1), ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59), Регистратор, %s) КАК О
ГДЕ (О.СубконтоДт3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл) ИЛИ О.СубконтоКт3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл))
	%s
"""
t0 = datetime.datetime.now()
DTKT = q(B, WH + SEL % ("СчетДт " + MATQ + ", , , ", "")) + q(B, WH + SEL % (", , СчетКт " + MATQ + ", ", "И НЕ О.СчетДт " + MATQ))
OST = q(B, WH + """
ВЫБРАТЬ О.Счет.Код КАК Счет, ВЫРАЗИТЬ(О.Субконто1 КАК Справочник.Номенклатура).Код КАК Код,
	ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Наименование КАК Скл, СУММА(О.КоличествоОстаток) КАК Кол, СУММА(О.СуммаОстаток) КАК Сумма
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет """ + MATQ + """, , ) КАК О
ГДЕ О.Субконто3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл)
СГРУППИРОВАТЬ ПО О.Счет.Код, ВЫРАЗИТЬ(О.Субконто1 КАК Справочник.Номенклатура).Код, ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Наименование""")
KOR = q(B, WH + """
ВЫБРАТЬ О.КорСчет.Код КАК Кор, О.Счет.Код КАК Счет, ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Наименование КАК Скл,
	СУММА(О.СуммаОборотДт) КАК Дт, СУММА(О.СуммаОборотКт) КАК Кт
ИЗ РегистрБухгалтерии.Хозрасчетный.Обороты(ДАТАВРЕМЯ(2026, 1, 1), ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59), , Счет """ + MATQ + """, , , , ) КАК О
ГДЕ О.Субконто3 В (ВЫБРАТЬ ВТ_Скл.Склад ИЗ ВТ_Скл КАК ВТ_Скл)
СГРУППИРОВАТЬ ПО О.КорСчет.Код, О.Счет.Код, ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Наименование""")
PTU = q(B, """
ВЫБРАТЬ Т.Ссылка.Номер КАК Номер, Т.Ссылка.Дата КАК Дата, Т.Ссылка.Склад.Наименование КАК Скл, Т.Номенклатура.Код КАК Код, Т.Количество КАК Кол,
	ВЫБОР КОГДА Т.Ссылка.СуммаВключаетНДС ТОГДА Т.Сумма - Т.СуммаНДС ИНАЧЕ Т.Сумма КОНЕЦ КАК Нетто,
	ВЫБОР КОГДА Т.Ссылка.СуммаВключаетНДС ТОГДА Т.Сумма ИНАЧЕ Т.Сумма + Т.СуммаНДС КОНЕЦ КАК Брутто
ИЗ Документ.ПоступлениеТоваровУслуг.Товары КАК Т
ГДЕ Т.Ссылка.Проведен И Т.Ссылка.Дата МЕЖДУ ДАТАВРЕМЯ(2026, 1, 1) И ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59)
	И (Т.Ссылка.Склад.Наименование ПОДОБНО "%IR%" ИЛИ Т.Ссылка.Склад.Наименование ПОДОБНО "%RС 2026%"
		ИЛИ Т.Ссылка.Подразделение.Наименование ПОДОБНО "%IRC 2026%" ИЛИ Т.Ссылка.Подразделение.Родитель.Наименование ПОДОБНО "%IRC 2026%")""")
SIDE_PTU = """
ВЫБРАТЬ П.Ссылка КАК Счет ПОМЕСТИТЬ ВТ_М ИЗ ПланСчетов.Хозрасчетный КАК П ГДЕ П.Код В ("20", "22", "28")
;
ВЫБРАТЬ Д.Ссылка КАК Док ПОМЕСТИТЬ ВТ_ПТУ ИЗ Документ.ПоступлениеТоваровУслуг КАК Д
ГДЕ Д.Проведен И Д.Дата >= ДАТАВРЕМЯ(2026, 1, 1)
	И (Д.Подразделение.Наименование ПОДОБНО "%IRC 2026%" ИЛИ Д.Подразделение.Родитель.Наименование ПОДОБНО "%IRC 2026%")
	И НЕ Д.Склад.Наименование ПОДОБНО "%IR%"
;
"""
SIDE_MOVES = q(B, SIDE_PTU + """
ВЫБРАТЬ ВЫРАЗИТЬ(О.СубконтоКт3 КАК Справочник.Склады).Наименование КАК Откуда, ВЫРАЗИТЬ(О.СубконтоДт3 КАК Справочник.Склады).Наименование КАК Куда,
	О.СчетДт.Код КАК СчДт, СУММА(О.СуммаОборот) КАК Сумма
ИЗ РегистрБухгалтерии.Хозрасчетный.ОборотыДтКт(ДАТАВРЕМЯ(2026, 1, 1), ДАТАВРЕМЯ(2026, 12, 31, 23, 59, 59), , , , СчетКт В ИЕРАРХИИ (ВЫБРАТЬ ВТ_М.Счет ИЗ ВТ_М КАК ВТ_М), , ) КАК О
ГДЕ О.СубконтоКт2 В (ВЫБРАТЬ ВТ_ПТУ.Док ИЗ ВТ_ПТУ КАК ВТ_ПТУ)
СГРУППИРОВАТЬ ПО ВЫРАЗИТЬ(О.СубконтоКт3 КАК Справочник.Склады).Наименование, ВЫРАЗИТЬ(О.СубконтоДт3 КАК Справочник.Склады).Наименование, О.СчетДт.Код""")
SIDE_LEFT = q(B, SIDE_PTU + """
ВЫБРАТЬ ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Наименование КАК Скл, О.Счет.Код КАК Сч, СУММА(О.СуммаОстаток) КАК Сумма
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет В ИЕРАРХИИ (ВЫБРАТЬ ВТ_М.Счет ИЗ ВТ_М КАК ВТ_М), , ) КАК О
ГДЕ О.Субконто2 В (ВЫБРАТЬ ВТ_ПТУ.Док ИЗ ВТ_ПТУ КАК ВТ_ПТУ)
СГРУППИРОВАТЬ ПО ВЫРАЗИТЬ(О.Субконто3 КАК Справочник.Склады).Наименование, О.Счет.Код""")
json.dump({"dtkt": DTKT, "ost": OST, "kor": KOR, "ptu": PTU, "side_moves": SIDE_MOVES, "side_left": SIDE_LEFT},
          open(os.path.join(VERDIR, "q_t4_raw.json"), "w", encoding="utf-8"), ensure_ascii=False, default=str)
print("COM:", len(DTKT), "оборотов,", len(OST), "остатков,", len(KOR), "оборотов×КорСчет,", len(PTU), "строк ПТУ")

# ═════════════ 2. Свёртки по моим правилам
J = json.load(open(JPATH, encoding="utf-8"))
T = J["totals"]
rtype = lambda s: re.match(r"(.+?)( ІБ| ІС| СТ|\s\d)", s or "").group(1) if re.match(r"(.+?)( ІБ| ІС| СТ|\s\d)", s or "") else (s or "")[:40]
PC = collections.defaultdict(lambda: [Z, Z]); OI = collections.defaultdict(lambda: [Z, Z]); OO = collections.defaultdict(lambda: [Z, Z])
WO = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z]))
KIND = collections.defaultdict(lambda: [Z, set(), set()]); OIK = collections.defaultdict(lambda: [Z, set(), set()]); OOK = collections.defaultdict(lambda: [Z, set(), set()])
WOK = collections.defaultdict(D)
for r in DTKT:
    dt, kt = r["СчДт"] or "", r["СчКт"] or ""
    cd, ck = wc(r["СклДт"]), wc(r["СклКт"])
    s = d(r["Сумма"]); typ = rtype(r["Рег"])
    if mat(dt) and insys(cd) and not mat(kt):
        a = PC[r["КодДт"]]; a[0] += d(r["КолДт"]); a[1] += s
        k = KIND[typ]; k[0] += s; k[1].add(r["КодДт"]); k[2].add(r["Рег"])
    if mat(dt) and insys(cd) and mat(kt) and not insys(ck):
        a = OI[r["КодДт"]]; a[0] += d(r["КолДт"]); a[1] += s
        k = OIK[norm(r["СклКт"])]; k[0] += s; k[1].add(r["КодДт"]); k[2].add(r["Рег"])
    if mat(kt) and insys(ck) and not mat(dt):
        if ck.startswith("H"):
            a = WO[r["КодКт"]][ck]; a[0] += d(r["КолКт"]); a[1] += s
            WOK[(ck[1:3] + " м", typ, r["Орг"])] += s
        else:
            a = OO[r["КодКт"]]; a[0] += d(r["КолКт"]); a[1] += s
            k = OOK["списание со склада проекта: " + typ + " (Дт " + dt[:3] + ")"]; k[0] += s; k[1].add(r["КодКт"]); k[2].add(r["Рег"])
    if mat(kt) and insys(ck) and mat(dt) and not insys(cd):
        a = OO[r["КодКт"]]; a[0] += d(r["КолКт"]); a[1] += s
        k = OOK["на «%s»" % norm(r["СклДт"])]; k[0] += s; k[1].add(r["КодКт"]); k[2].add(r["Рег"])
STK = collections.defaultdict(lambda: [Z, Z]); STKP = collections.defaultdict(lambda: [Z, Z]); STW = collections.defaultdict(D); STC = collections.defaultdict(set)
for r in OST:
    if not mat(r["Счет"]):
        continue
    c = wc(r["Скл"])
    a = STK[r["Код"]]; a[0] += d(r["Кол"]); a[1] += d(r["Сумма"])
    if c == "P":
        b = STKP[r["Код"]]; b[0] += d(r["Кол"]); b[1] += d(r["Сумма"])
    STW[norm(r["Скл"])] += d(r["Сумма"])
    if d(r["Сумма"]) or d(r["Кол"]):
        STC[c].add(r["Код"])
OUT = {"stamp": t0.strftime("%d.%m.%Y %H:%M:%S"), "base": 'BuhBud Srvr="localhost";Ref="bas_industrialbud" — только ВЫБРАТЬ; ERP не трогалась',
       "script": "Z6_scripts/proverka/q_t4_proverka.py",
       "paths": "ОборотыДтКт(Регистратор), Обороты(КорСчет), Остатки(), ТЧ ПТУ, ОборотыДтКт по партии ПТУ — другие ВТ, чем ДвиженияССубконто в z6_collect.py"}
F = OUT["facts"] = {}
F["purchased"] = {"by_kind": {k: {"uah_net": r2(v[0]), "cards": len(v[1]), "docs": len(v[2])} for k, v in KIND.items()},
                  "uah_net": r2(sum(v[1] for v in PC.values())), "cards": len([c for c, v in PC.items() if v[0] or v[1]]),
                  "json_uah_net": T["purchased_uah_net"]}
ptu_sys = [r for r in PTU if insys(wc(r["Скл"]))]
ptu_side = [r for r in PTU if not insys(wc(r["Скл"]))]
F["ptu_documents"] = {"docs": len({r["Номер"] + str(r["Дата"]) for r in ptu_sys}), "cards": len({r["Код"] for r in ptu_sys}),
                      "qty": r2(sum(d(r["Кол"]) for r in ptu_sys)), "uah_net": r2(sum(d(r["Нетто"]) for r in ptu_sys)),
                      "uah_vat": r2(sum(d(r["Брутто"]) for r in ptu_sys)), "warehouses": sorted({norm(r["Скл"]) for r in ptu_sys}),
                      "side_uah_net": r2(sum(d(r["Нетто"]) for r in ptu_side)), "side_uah_vat": r2(sum(d(r["Брутто"]) for r in ptu_side)),
                      "side_cards": len({r["Код"] for r in ptu_side})}
reg_ptu = collections.defaultdict(lambda: [Z, Z]); doc_ptu = collections.defaultdict(lambda: [Z, Z])
for r in DTKT:
    if mat(r["СчДт"]) and not mat(r["СчКт"]) and (r["Рег"] or "").startswith("Надходження товарів і послуг") and insys(wc(r["СклДт"])):
        reg_ptu[r["КодДт"]][0] += d(r["КолДт"]); reg_ptu[r["КодДт"]][1] += d(r["Сумма"])
for r in ptu_sys:
    doc_ptu[r["Код"]][0] += d(r["Кол"]); doc_ptu[r["Код"]][1] += d(r["Нетто"])
F["ptu_documents"]["cards_doc_vs_register_mismatch"] = sum(1 for c in set(reg_ptu) | set(doc_ptu)
                                                           if abs(reg_ptu[c][0] - doc_ptu[c][0]) > D("0.001") or abs(reg_ptu[c][1] - doc_ptu[c][1]) > D("0.01"))
rows = {r["code"]: r for r in J["rows"] if not r["code"].startswith("СС30")}
F["top15_purchased"] = [{"code": c, "card": rows[c]["card"] if c in rows else "?", "unit": rows[c]["unit"] if c in rows else "",
                         "qty": float(v[0]), "uah_net": r2(v[1]), "json_qty": rows[c]["purchased"]["qty"] if c in rows else None,
                         "json_uah_net": rows[c]["purchased"]["uah_net"] if c in rows else None}
                        for c, v in sorted(PC.items(), key=lambda kv: -kv[1][1])[:15]]
hs = collections.defaultdict(lambda: [Z, set()])
for c, dd in WO.items():
    for h, v in dd.items():
        hs[h][0] += v[1]; hs[h][1].add(c)
F["written_off_by_house"] = {HN[h]: {"uah_net": r2(v[0]), "cards": len(v[1]), "json_uah_net": T["by_house"][HN[h]]["written_off_uah_net"],
                                     "json_cards": T["by_house"][HN[h]]["cards_written_off"]} for h, v in sorted(hs.items())}
F["written_off_15_uah_net"] = r2(sum(v[0] for h, v in hs.items() if h.startswith("H15")))
F["written_off_30_uah_net"] = r2(sum(v[0] for h, v in hs.items() if h.startswith("H30")))
F["written_off_by_doc_org"] = {" | ".join(k): r2(v) for k, v in sorted(WOK.items())}
F["stock"] = {"project_uah_net": r2(sum(v for w, v in STW.items() if wc(w) == "P")), "project_cards": len(STC["P"]),
              "houses_uah_net": r2(sum(v for w, v in STW.items() if wc(w).startswith("H"))),
              "by_warehouse": {w: r2(v) for w, v in sorted(STW.items()) if v}, "account_26_houses_uah_net_not_materials":
              r2(sum(d(r["Сумма"]) for r in OST if (r["Счет"] or "")[:2] == "26"))}
F["other_in"] = {"uah_net": r2(sum(v[1] for v in OI.values())), "by_source": {k: {"uah_net": r2(v[0]), "cards": len(v[1]), "docs": len(v[2])} for k, v in OIK.items()}}
F["other_out"] = {"uah_net": r2(sum(v[1] for v in OO.values())), "by_kind": {k: {"uah_net": r2(v[0]), "cards": len(v[1]), "docs": len(v[2])} for k, v in OOK.items()}}
kor = collections.defaultdict(D)
for r in KOR:
    if not mat(r["Счет"]):
        continue
    c = wc(r["Скл"]); cls = "P" if c == "P" else (c[:3] if c.startswith("H") else c)
    km = "мат" if mat(r["Кор"]) else "не мат"
    kor[(cls, "Дт", km)] += d(r["Дт"]); kor[(cls, "Кт", km)] += d(r["Кт"])
F["korschet_path"] = {"purchased_uah_net": r2(sum(v for k, v in kor.items() if k[1] == "Дт" and k[2] == "не мат" and insys(k[0]))),
                      "written_off_15_uah_net": r2(kor[("H15", "Кт", "не мат")]), "written_off_30_uah_net": r2(kor[("H30", "Кт", "не мат")]),
                      "project_written_off_not_house_uah_net": r2(kor[("P", "Кт", "не мат")])}
sm = collections.defaultdict(lambda: collections.defaultdict(D))
for r in SIDE_MOVES:
    sm[norm(r["Откуда"])][norm(r["Куда"]) if mat(r["СчДт"]) else "списано (Дт %s)" % r["СчДт"]] += d(r["Сумма"])
F["side_purchases_flow"] = {
    "into_irc_system": r2(sum(v for s in sm.values() for w, v in s.items() if insys(wc(w)))),
    "to_other_objects": {s: {w: r2(v) for w, v in dd.items() if not insys(wc(w)) and "МШП" not in w} for s, dd in sm.items()},
    "to_other_mshp_warehouses": r2(sum(v for s in sm.values() for w, v in s.items() if not insys(wc(w)) and "МШП" in w)),
    "still_on_purchase_warehouse": {"%s / %s" % (norm(r["Скл"]), r["Сч"]): r2(r["Сумма"]) for r in SIDE_LEFT}}
F["side_purchases_flow"]["to_other_objects"] = {s: v for s, v in F["side_purchases_flow"]["to_other_objects"].items() if v}

# ═════════════ 3. По карточке: мои данные против JSON
codes = set(PC) | set(OI) | set(OO) | set(WO) | set(STK)
mism = []
for c in sorted(codes | set(rows)):
    j = rows.get(c)
    wq = sum(v[0] for v in WO[c].values()); ws = sum(v[1] for v in WO[c].values())
    mine = {"purchased.qty": PC[c][0], "purchased.uah_net": PC[c][1], "other_in.qty": OI[c][0], "other_in.uah_net": OI[c][1],
            "other_out.qty": OO[c][0], "other_out.uah_net": OO[c][1], "written_off.total": wq, "written_off.uah_net": ws,
            "stock_ostatki_now.qty": STK[c][0], "stock_ostatki_now.uah_net": STK[c][1], "stock.project": STKP[c][0], "stock.project_uah_net": STKP[c][1]}
    if abs(PC[c][0] + OI[c][0] - wq - OO[c][0] - STK[c][0]) > D("0.0005") or abs(PC[c][1] + OI[c][1] - ws - OO[c][1] - STK[c][1]) > D("0.005"):
        mism.append({"code": c, "what": "замкнутость по моим данным"})
    if j is None:
        mism.append({"code": c, "what": "нет в JSON"}); continue
    for k, v in mine.items():
        a, b = k.split(".")
        jv = d(j[a][b])
        if abs((v if b.endswith("qty") or b in ("total", "project") else d(r2(v))) - jv) > (D("0.0015") if b in ("qty", "total", "project") else D("0.011")):
            mism.append({"code": c, "what": k, "mine": float(v), "json": float(jv)})
    for h, v in WO[c].items():
        jb = j["written_off"]["by_house"].get(HN[h])
        if jb is None or abs(v[0] - d(jb["qty"])) > D("0.0015") or abs(d(r2(v[1])) - d(jb["uah_net"])) > D("0.011"):
            mism.append({"code": c, "what": "списано " + HN[h]})
OUT["per_card_vs_json"] = {"cards_mine": len(codes), "cards_json": len(rows), "mismatches": len(mism), "examples": mism[:20]}

# ═════════════ 4. Инварианты на JSON
E = collections.defaultdict(list)
H15 = [f"15 м №{i}" for i in range(1, 7)]
hq = lambda r, h: d((r["stock"]["houses"].get(h) or {}).get("qty") if isinstance(r["stock"]["houses"].get(h), dict) else r["stock"]["houses"].get(h))
S = collections.defaultdict(D)
for r in J["rows"]:
    tb = r["to_buy_30"]; ps = r["code"].startswith("СС30"); p = r.get("need_30_parts") or {}
    S["buy_g"] += d(tb["uah_last"]); S["buy_n"] += d(tb["uah_last_net"]); S["need30"] += d(r["need_30_uah_net"])
    if d(tb["qty"]) < 0:
        E["докупить < 0"].append(r["code"])
    if ps:
        if abs(d(tb["qty"]) - d(r["need_30"])) > D("0.0015"):
            E["строка без карточки: докупить ≠ потребность"].append(r["code"])
        continue
    wb = r["written_off"]["by_house"]
    f15 = sum(d((wb.get(h) or {}).get("qty")) for h in H15) + sum(hq(r, h) for h in H15)
    if abs(f15 - d(r["need_15"])) > D("0.0015"):
        E["потребность 15 ≠ факт 15"].append(r["code"])
    f1 = d((wb.get("30 м №1") or {}).get("qty")) + hq(r, "30 м №1"); f2 = d((wb.get("30 м №2") or {}).get("qty")) + hq(r, "30 м №2")
    f3 = d((wb.get("30 м №3") or {}).get("qty")) + hq(r, "30 м №3")
    ph = d(p.get("avg_1_2")) + d(p.get("etalon_lines_not_issued_per_house"))
    if abs(f1 - d(p.get("fact_1"))) > D("0.0015") or abs(f2 - d(p.get("fact_2"))) > D("0.0015") or abs(f3 - d(p.get("fact_3_so_far"))) > D("0.0015"):
        E["факт №1–№3 ≠ списано + лежит"].append(r["code"])
    if abs(f1 + f2 + max(ph, f3) + 4 * ph - d(r["need_30"])) > D("0.005"):       # части в JSON округлены до 0,001
        E["потребность 30 ≠ факт №1–№2 + max(ср., факт №3) + ср. × 4 (+ строки эталона)"].append(r["code"])
    qr, qq = d(tb["qty_raw"]), d(tb["qty"])
    exp = D(math.ceil(qr - D("0.0000001"))) if (r["piece"] and qr > 0) else qr
    if abs(qq - exp) > D("0.0015"):
        E["округление не на итог карточки"].append(r["code"])
    if abs(sum(d(v["buy"]) for v in tb["by_house"].values()) - qq) > D("0.0015"):
        E["Σ по домам ≠ докупить"].append(r["code"])
    if r["piece"] and any(d(v["buy"]) != d(v["buy"]).to_integral_value() for v in tb["by_house"].values()):
        E["дробная штука по дому"].append(r["code"])
    sp, sur = max(Z, d(r["stock"]["project"])), d(r["surplus"]["qty"])
    if abs(sp + qr - d(p.get("remaining_3_7")) - sur) > D("0.003"):
        E["склад + докупить ≠ остаток потребности + лишнее"].append(r["code"])
    if qr > D("0.0015") and sur > D("0.0015"):
        E["одновременно докупить и лишнее"].append(r["code"])
    if d(r["surplus"]["uah_net"]) - d(r["stock"]["project_uah_net"]) > D("0.011"):
        E["лишнее дороже склада"].append(r["code"])
    lhs = d(r["purchased"]["qty"]) + d(r["other_in"]["qty"]) - d(r["other_out"]["qty"]) + qq
    rhs = d(r["need_15"]) + d(r["need_30"]) + sur + (qq - qr)
    if abs(lhs - rhs) > D("0.005"):
        E["тождество по количеству"].append(r["code"])
nodes = collections.Counter(r["node"] for r in J["rows"])
BN = J["by_node"]
node_sum = {k: r2(sum(d(v.get(k)) for v in BN.values())) for k in ("need_15_uah_net", "need_30_uah_net", "purchased_uah_net", "written_off_uah_net",
                                                                    "lies_project_uah_net", "lies_houses_uah_net", "to_buy_30_uah_last_vat",
                                                                    "to_buy_30_uah_last_net", "surplus_uah_net")}
OUT["invariants"] = {
    "violations": {k: {"count": len(v), "codes": v[:10]} for k, v in E.items()},
    "rows": len(J["rows"]), "unique_codes": len({r["code"] for r in J["rows"]}), "rows_without_node": sum(1 for r in J["rows"] if not r.get("node")),
    "nodes_in_rows_not_in_by_node": sorted(set(nodes) - set(BN)), "by_node_cards": sum(v["cards"] for v in BN.values()),
    "by_node_lines_no_card": sum(v.get("lines_no_card", 0) for v in BN.values()), "by_node_sums": node_sum,
    "rows_sum_vs_totals": {"to_buy_30_uah_last_vat": [r2(S["buy_g"]), T["to_buy_30_uah_last_vat"]],
                           "to_buy_30_uah_last_net": [r2(S["buy_n"]), T["to_buy_30_uah_last_net"]],
                           "need_30_uah_net": [r2(S["need30"]), T["need_30_uah_net"]]},
    "rounding_note": "итоги считаются из неокруглённых значений; сумма округлённых строк/узлов/домов отличается от итога на 0,01–0,07"}

# ═════════════ 5. Воспроизводимость: z6_calc.py на том же raw.json во временную папку
src = open(os.path.join(Z6S, "z6_calc.py"), encoding="utf-8").read()
src = src.replace('HERE = os.path.dirname(os.path.abspath(__file__))', 'HERE = %r' % Z6S, 1)
src = src.replace('OUT_JSON = os.path.join(DATA, "proekt_irc2026_zakupki.json")', 'OUT_JSON = %r' % os.path.join(VERDIR, "rerun.json"))
src = src.replace('OUT_TXT = os.path.join(HERE, "z6_out.txt")', 'OUT_TXT = %r' % os.path.join(VERDIR, "rerun_out.txt"))
open(os.path.join(VERDIR, "z6_calc_rerun.py"), "w", encoding="utf-8").write(src)
rc = subprocess.run([sys.executable, "-X", "utf8", os.path.join(VERDIR, "z6_calc_rerun.py")], capture_output=True, text=True, encoding="utf-8",
                    env={**os.environ, "Z6_RAW_DIR": RAWDIR}).returncode
diffs = []
if rc == 0:
    A = {k: v for k, v in J.items() if k != "verification"}
    Bj = {k: v for k, v in json.load(open(os.path.join(VERDIR, "rerun.json"), encoding="utf-8")).items() if k != "verification"}

    def walk(a, b, p=""):
        if isinstance(a, dict) and isinstance(b, dict):
            for k in set(a) | set(b):
                if k not in a or k not in b:
                    diffs.append(p + "/" + str(k)); continue
                walk(a[k], b[k], p + "/" + str(k))
        elif isinstance(a, list) and isinstance(b, list):
            if len(a) != len(b):
                diffs.append(p + " len")
            for i, (x, y) in enumerate(zip(a, b)):
                walk(x, y, f"{p}[{i}]")
        elif a != b:
            diffs.append(p)
    walk(A, Bj)
OUT["reproducibility"] = {"rerun_exit_code": rc, "differences_vs_published_json": len(diffs), "examples": diffs[:10]}
OUT["corrections"] = [
    {"what": "закупки на подразделение IRC мимо склада проекта (80 722,90 без ПДВ)",
     "was": "в систему они приходят перемещением и учтены в прочих приходах; около 20 239,47 ушло с ІБ_МШП прямо на другие объекты "
            "(Кривий ріг, ООН Админ №1/№2, Каменец Цемент, Чорноморськ)",
     "now": "в систему IRC пришло 8 546,75; на другие объекты — 20 239,47: с ІБ_МШП 11 559,47 (ООН Админ №1 5 564,93, Каменец Цемент 4 344,34, "
            "ООН Админ №2 1 650,20), с «ІС Основной склад» 8 680,00 (Кривий ріг 7 595,00, Чорноморськ 1 085,00); 175,50 — на «ІБ_МШП виробничий»; "
            "51 761,18 лежит на ІБ_МШП (221 — 44 410,78, паливо 203 — 7 350,40)",
     "where": "заметка §3.3, §8 п. 6, §10 п. 3; JSON caveats, totals.purchased_irc_subdivision_other_warehouse.flow_uah_net"},
    {"what": "сравнение с бюджетом v3",
     "was": "против v3 (1 313 551,98) +44 495,55; почти всё — обогреватели, остальное +1 318,59 — разница Z1 с v3",
     "now": "v3 G_h — норма СС × цена СС, з ПДВ; МШП в v3 — отдельный член МШП_h 12 002,98. +44 495,55 сравнивает разные базы цены, "
            "по одной базе (цена СС) +173 667,80. Z6 включает МШП, поэтому в бюджете заменяет G_h + МШП_h = 1 325 554,96: "
            "+32 492,57 по посл. закупке, +161 664,82 по цене СС",
     "where": "заметка §0, §4.2, §9 п. 1, §10 п. 5 и 10; JSON compare.to_buy_30_vs_budget_v3, caveats"}]
OUT["remarks"] = [
    "Решение 3: «Група» крепежа взаимозаменяема, а Z6 неттингует склад по карточке: лишние 3 375 шт «Саморіз 4,8*19» не закрывают докупку "
    "других карточек группы. Неттинг по группе снизил бы докупку на 890,06…5 463,99 з ПДВ. Итог не меняли (totals.group_fasteners_netting_check).",
    "Округление: итоги считаются из неокруглённых значений. Сумма округлённых домов 1 358 047,54 (итог 1 358 047,53), домов №3–№6 837 349,75 "
    "(октябрь 837 349,74), узлов — на 0,01–0,02 больше итогов, строк JSON — на 0,00–0,05. Это не ошибка.",
    "Вторая дверь на №2: 3 × 22 037,148571… = 66 111,45 (верно); снижение по четырём решениям 336 106,43, докупка 1 021 941,10 — верно."]
OUT["stamp_end"] = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
json.dump(OUT, open(os.path.join(HERE, "q_t4_out.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
print(json.dumps({"per_card": OUT["per_card_vs_json"]["mismatches"], "violations": {k: v["count"] for k, v in OUT["invariants"]["violations"].items()},
                  "repro": OUT["reproducibility"]["differences_vs_published_json"], "purchased": F["purchased"]["uah_net"],
                  "written_off": [F["written_off_15_uah_net"], F["written_off_30_uah_net"]], "stock": [F["stock"]["project_uah_net"], F["stock"]["houses_uah_net"]],
                  "other": [F["other_in"]["uah_net"], F["other_out"]["uah_net"]], "side": F["side_purchases_flow"]}, ensure_ascii=False, indent=1))
