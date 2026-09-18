# -*- coding: utf-8 -*-
"""Анализ дампов книги ДДС и базы (без COM, без записи).

Вход: dds_kniga_dump.json, dds_baza_dump.json (из measure_dds_*).
Выход на экран:
  A. Карта «лист → роль → подразделение»: РС выгрузки (строго/нормализовано), затем
     СтруктураПредприятия (строго, затем мягко), с пометкой «требует уточнения».
  B. Диапазон данных листа (до служебного хвоста), проверка иерархии:
     сумма уровня 1 = сумма его детей уровня 2? листья на уровне 0/1 с деньгами.
  C. Эталон: по листу и месяцу сумма уровня 2 (безнал / нал), число ненулевых ячеек.
  D. Статьи уровня 2: сопоставление с рабочими статьями ДДС по нормализованному имени,
     проверка стороны (ПОСТУПЛЕНИЯ ↔ Приход, СПИСАНИЕ ↔ Расход), ненайденные.
  E. Знак: контрольные строки.
"""
import sys, json, os, re
sys.stdout.reconfigure(encoding='utf-8')
from collections import defaultdict, Counter

D = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
kniga = json.load(open(os.path.join(D, "dds_kniga_dump.json"), encoding="utf-8"))
baza = json.load(open(os.path.join(D, "dds_baza_dump.json"), encoding="utf-8"))


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").replace("\t", " ").replace("\n", " ").replace("\r", " ")
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip().lower()


def soft(s):
    s = norm(s)
    for ch in "ьъ'`’ʼ\".,":
        s = s.replace(ch, "")
    for a, b in (("і", "и"), ("ї", "и"), ("ы", "и"), ("є", "е"), ("ё", "е"), ("ґ", "г")):
        s = s.replace(a, b)
    # латиница/кириллица-двойники в именах объектов (IRC / IRС, EMS / ЕМS)
    for a, b in (("c", "с"), ("e", "е"), ("s", "ѕ"), ("m", "м"), ("i", "і"), ("r", "r")):
        pass
    tr = str.maketrans({"c": "с", "e": "е", "m": "м", "s": "s", "o": "о", "a": "а", "p": "р", "x": "х", "k": "к", "h": "н", "t": "т", "b": "в"})
    return s.translate(tr)


def fmt(x):
    return f"{x:,.2f}".replace(",", " ")


TAIL_RE = re.compile(r"free cf|переходящ|сальдо|cf за|баланс по объекту|fcf")

# ---------- A. карта листов ----------
print("=" * 110)
print("A. КАРТА «ЛИСТ → ПОДРАЗДЕЛЕНИЕ»")
rs_by_name = {}
for r in baza["rs_vygruzki"]:
    k = norm(r["ИмяСтраницы"])
    if k and k not in rs_by_name:
        rs_by_name[k] = r
podr_strict = defaultdict(list)
podr_soft = defaultdict(list)
for r in baza["all_podr"]:
    if r["Пометка"]:
        continue
    podr_strict[norm(r["Наименование"])].append(r)
    podr_soft[soft(r["Наименование"])].append(r)
under_prod = {r["Код"] for r in baza["pod_proizvodstvo"]}

sheet_map = {}
for sh in kniga["sheets"]:
    name = sh["name"]
    k = norm(name)
    rec = {"sheet": name, "state": sh["state"], "blocks": [(b["col"], b["month"]) for b in sh.get("blocks", [])],
           "via": None, "code": None, "podr": None, "flags": [], "candidates": []}
    if name.startswith("Date"):
        rec["via"] = "сырые движения"; rec["flags"].append("служебный")
        sheet_map[name] = rec; continue
    if not sh.get("blocks"):
        rec["via"] = "нет датовых блоков в строке 3"; rec["flags"].append("служебный")
        sheet_map[name] = rec; continue
    # РС выгрузки
    if k in rs_by_name:
        r = rs_by_name[k]
        rec["via"] = "РС выгрузки"; rec["code"] = r["ПодрКод"]; rec["podr"] = r["ПодрИмя"]
        if r["ОтчетПоНаправлению"]:
            rec["flags"].append("СВОД направления (РС)")
        if r["ТолькоЭтоПодразделение"]:
            rec["flags"].append("ЦО направления (РС)")
    else:
        cand = podr_strict.get(k, [])
        if len(cand) == 1:
            rec["via"] = "СтруктураПредприятия строго"; rec["code"] = cand[0]["Код"]; rec["podr"] = cand[0]["Наименование"]
        elif len(cand) > 1:
            rec["via"] = "НЕОДНОЗНАЧНО строго"; rec["candidates"] = [(c["Код"], c["Наименование"], c["РодительИмя"]) for c in cand]
            rec["flags"].append("ТРЕБУЕТ УТОЧНЕНИЯ: несколько строгих совпадений")
        else:
            cand = podr_soft.get(soft(name), [])
            if len(cand) == 1:
                rec["via"] = "СтруктураПредприятия МЯГКО"; rec["code"] = cand[0]["Код"]; rec["podr"] = cand[0]["Наименование"]
                rec["flags"].append("подобрано нестрого — проверьте")
            elif len(cand) > 1:
                rec["via"] = "НЕОДНОЗНАЧНО мягко"; rec["candidates"] = [(c["Код"], c["Наименование"], c["РодительИмя"]) for c in cand]
                rec["flags"].append("ТРЕБУЕТ УТОЧНЕНИЯ: несколько мягких совпадений")
            else:
                rec["via"] = "НЕ НАЙДЕНО"; rec["flags"].append("ТРЕБУЕТ УТОЧНЕНИЯ: подразделение не найдено")
                # подсказки: похожие имена под Производством
                toks = [t for t in soft(name).split() if len(t) > 2]
                hints = []
                for r in baza["pod_proizvodstvo"]:
                    sn = soft(r["Наименование"])
                    if any(t in sn for t in toks):
                        hints.append((r["Код"], r["Наименование"]))
                rec["candidates"] = hints[:8]
    if rec["code"] and rec["code"] not in under_prod:
        rec["flags"].append(f"НЕ под «Производство» (0Ц-000004)")
    if sh["state"] != "visible":
        rec["flags"].append("скрыт в книге")
    sheet_map[name] = rec

for name, rec in sheet_map.items():
    print(f"  {name!r:<34} {rec['state']:<8} блоки={[m for _, m in rec['blocks']]}")
    print(f"      → {rec['via']}; {rec['code']} {rec['podr']}; {'; '.join(rec['flags'])}"
          + (f"; кандидаты: {rec['candidates']}" if rec['candidates'] else ""))

# ---------- B/C/D. по листам с блоками ----------
st_by_norm = defaultdict(list)
for r in baza["statyi"]:
    st_by_norm[norm(r["Наименование"])].append(r)

etalon = {}
article_usage = defaultdict(lambda: {"income": set(), "expense": set(), "sum": 0.0})
print("\n" + "=" * 110)
print("B/C. ДИАПАЗОН ДАННЫХ, ИЕРАРХИЯ, ЭТАЛОН ПО ЛИСТАМ")
for sh in kniga["sheets"]:
    if not sh.get("blocks") or sh["name"].startswith("Date"):
        continue
    rows = sh["rows"]
    months = [b["month"] for b in sh["blocks"]]
    # хвост
    tail_row = None
    for x in rows:
        if x["r"] >= 100 and x["name"] and TAIL_RE.search(norm(x["name"])):
            tail_row = x["r"]; break
    data_rows = [x for x in rows if x["r"] >= 6 and (tail_row is None or x["r"] < tail_row)]
    # секция: ПОСТУПЛЕНИЯ до строки СПИСАНИЕ
    spis_row = None
    for x in data_rows:
        if x["name"] and norm(x["name"]) == "списание":
            spis_row = x["r"]; break
    # иерархия: для каждой строки уровня 1 — сумма детей уровня 2 до следующей строки уровня <=1
    hier_bad = []
    leaf01 = []   # строки уровня 0/1 с деньгами, у которых нет детей уровня 2
    lvl3 = []
    for i, x in enumerate(data_rows):
        if x["level"] == 3:
            lvl3.append((x["r"], x["name"]))
        if x["level"] not in (0, 1):
            continue
        own = {m: (x["vals"][m]["bn"], x["vals"][m]["nal"]) for m in months}
        kids = {m: [0.0, 0.0] for m in months}
        nkids = 0
        for y in data_rows[i + 1:]:
            if y["level"] <= x["level"]:
                break
            if y["level"] == 2 or (x["level"] == 0 and y["level"] == 1):
                # для уровня 0 дети — уровень 1 (а их дети — уровень 2); считаем прямых детей
                pass
            if y["level"] == x["level"] + 1:
                nkids += 1
                for m in months:
                    kids[m][0] += y["vals"][m]["bn"]; kids[m][1] += y["vals"][m]["nal"]
        has_money = any(abs(v[0]) > 0.005 or abs(v[1]) > 0.005 for v in own.values())
        if nkids == 0 and has_money:
            leaf01.append((x["r"], x["level"], x["name"], {m: own[m] for m in months if own[m] != (0.0, 0.0)}))
        elif nkids > 0:
            for m in months:
                if abs(own[m][0] - kids[m][0]) > 0.011 or abs(own[m][1] - kids[m][1]) > 0.011:
                    hier_bad.append((x["r"], x["level"], x["name"], m, "own", (round(own[m][0], 2), round(own[m][1], 2)), "kids", (round(kids[m][0], 2), round(kids[m][1], 2))))
    # эталон уровня 2
    et = {}
    for m in months:
        bn = nal = 0.0; cnt_bn = cnt_nal = 0; rows_used = 0
        for x in data_rows:
            if x["level"] != 2:
                continue
            v = x["vals"][m]
            if v["bn"] != 0:
                bn += v["bn"]; cnt_bn += 1
            if v["nal"] != 0:
                nal += v["nal"]; cnt_nal += 1
            if v["bn"] != 0 or v["nal"] != 0:
                rows_used += 1
        et[m] = {"bn": round(bn, 2), "nal": round(nal, 2), "cells_bn": cnt_bn, "cells_nal": cnt_nal, "rows": rows_used}
    etalon[sh["name"]] = et
    lvl2_named = [x for x in data_rows if x["level"] == 2 and x["name"]]
    lvl2_noname = [x for x in data_rows if x["level"] == 2 and not x["name"] and any(x["vals"][m]["bn"] or x["vals"][m]["nal"] for m in months)]
    print(f"\n  ЛИСТ «{sh['name']}» ({sh['state']}): данные строки 6..{(tail_row - 1) if tail_row else '?'}, хвост с {tail_row}, СПИСАНИЕ на {spis_row}; уровень-2 строк с именем {len(lvl2_named)}, без имени с деньгами {len(lvl2_noname)}; уровень 3: {lvl3[:5]}")
    for m in months:
        e = et[m]
        print(f"      {m}: L2 безнал {fmt(e['bn']):>16} ({e['cells_bn']:>2} яч.)  нал {fmt(e['nal']):>15} ({e['cells_nal']:>2} яч.)  строк {e['rows']:>2}")
    if leaf01:
        print(f"      ЛИСТЬЯ НА УРОВНЕ 0/1 С ДЕНЬГАМИ (уровень-2 их не увидит): {len(leaf01)}")
        for l in leaf01[:12]:
            print(f"         r{l[0]} L{l[1]} «{l[2]}»: {l[3]}")
    if hier_bad:
        print(f"      ИЕРАРХИЯ НЕ СХОДИТСЯ (строка ≠ сумма прямых детей): {len(hier_bad)}")
        for h in hier_bad[:12]:
            print(f"         r{h[0]} L{h[1]} «{h[2]}» {h[3]}: own={h[5]} kids={h[7]}")
    # использование статей
    for x in lvl2_named:
        side = "income" if (spis_row is None or x["r"] < spis_row) else "expense"
        tot = sum(x["vals"][m]["bn"] + x["vals"][m]["nal"] for m in months)
        if tot == 0 and not any(x["vals"][m]["bn"] or x["vals"][m]["nal"] for m in months):
            continue
        u = article_usage[norm(x["name"])]
        u[side].add(sh["name"]); u["sum"] += tot

# ---------- D. статьи ----------
print("\n" + "=" * 110)
print("D. СТАТЬИ УРОВНЯ 2 С ДЕНЬГАМИ (по всем листам с блоками) ↔ рабочие статьи ДДС")
not_found, ambiguous, side_mismatch, ok = [], [], [], []
for nm, u in sorted(article_usage.items(), key=lambda kv: -abs(kv[1]["sum"])):
    cand = st_by_norm.get(nm, [])
    sides = ("П" if u["income"] else "") + ("Р" if u["expense"] else "")
    if not cand:
        not_found.append((nm, sides, round(u["sum"], 2), sorted(u["income"] | u["expense"])[:4]))
    elif len(cand) > 1:
        ambiguous.append((nm, sides, [(c["Код"], c["Тип"], c["Родитель"]) for c in cand]))
    else:
        c = cand[0]
        t = c["Тип"] or ""
        exp = set()
        if u["income"]:
            exp.add("Приход")
        if u["expense"]:
            exp.add("Расход")
        if t and t not in exp or (u["income"] and u["expense"]):
            side_mismatch.append((nm, sides, c["Код"], t, round(u["sum"], 2)))
        ok.append((nm, c["Код"], t))
print(f"  всего имён с деньгами: {len(article_usage)}; найдено однозначно {len(ok)}; НЕ найдено {len(not_found)}; неоднозначно {len(ambiguous)}; сторона/тип не сходится {len(side_mismatch)}")
print("  НЕ НАЙДЕНО (имя, сторона П/Р, сумма по книге, листы):")
for x in not_found:
    print(f"     {x}")
print("  НЕОДНОЗНАЧНО:")
for x in ambiguous:
    print(f"     {x}")
print("  СТОРОНА В КНИГЕ ≠ А_ТипПоказателя статьи (или имя и в приходе, и в расходе):")
for x in side_mismatch:
    print(f"     {x}")
print("  найденные без типа показателя:", [(n, c) for n, c, t in ok if not t])

# ---------- E. знак ----------
print("\n" + "=" * 110)
print("E. ЗНАК")
for sh in kniga["sheets"]:
    if sh["name"] in ("МД IRC 2026", "ЦО_Производство", "МД ООН 2026"):
        for x in sh["rows"]:
            if x["name"] and norm(x["name"]) in ("списание", "поступления", "зарплата производственного персонала"):
                print(f"  «{sh['name']}» r{x['r']} L{x['level']} «{x['name']}»: " + "; ".join(f"{m}: бн {fmt(v['bn'])} нал {fmt(v['nal'])}" for m, v in x["vals"].items()))
neg_l2 = []
for sh in kniga["sheets"]:
    for x in sh.get("rows", []):
        if x["level"] == 2 and x["name"]:
            for m, v in x["vals"].items():
                if v["bn"] < 0 or v["nal"] < 0:
                    neg_l2.append((sh["name"], x["r"], x["name"], m, v["bn"], v["nal"]))
print(f"  отрицательных ячеек уровня 2 по всей книге: {len(neg_l2)}: {neg_l2[:10]}")

json.dump({"sheet_map": sheet_map, "etalon": etalon,
           "articles": {"not_found": not_found, "ambiguous": ambiguous, "side_mismatch": side_mismatch, "ok": ok}},
          open(os.path.join(D, "dds_analysis.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=list)
print("\nСохранено:", os.path.join(D, "dds_analysis.json"))
