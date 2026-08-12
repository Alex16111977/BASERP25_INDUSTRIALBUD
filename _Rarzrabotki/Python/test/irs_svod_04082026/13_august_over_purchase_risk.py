# -*- coding: utf-8 -*-
"""Августовский план закупок (МАТЕРИАЛИ.xlsx) против уже закупленного — по общим названиям.
Цель: не купить повторно то, что уже закуплено общим котлом в июле.

Сравнение ведётся В ДЕНЬГАХ (грн с НДС): единицы измерения в кошторисе Excel и в СС
для части позиций РАЗНЫЕ (Excel — м.п./м2, СС — тонны), поэтому количественное
сопоставление по ним недопустимо. Расхождение единиц выводится отдельным списком.
"""
import json, sys, re
sys.stdout.reconfigure(encoding='utf-8')

pos = json.load(open("data_positions.json", encoding="utf-8"))["plan"]
xl = json.load(open("data_excel_koshtoris.json", encoding="utf-8"))
byON = {d["ОН"]: d for d in json.load(open("data_byon_calc.json", encoding="utf-8"))}


def norm(s):
    s = (s or "").lower().replace("ё", "е").replace("*", "х").replace("×", "х")
    s = s.replace(",", ".").replace("ʼ", "'").replace("’", "'")
    s = re.sub(r"[\s ]+", " ", s)
    return re.sub(r"[^0-9a-zа-яіїєґ'.х/]+", " ", s).strip()


def unorm(u):
    return re.sub(r"[^a-zа-яіїєґ0-9]", "", (u or "").lower())


SS = {"15": "МД IRS 2026 15 м (МД IRS 2026)", "30": "МД IRS 2026 30 м (МД IRS 2026)"}

aug = {}
unmapped = []
unit_bad = []
for tag, ssname in SS.items():
    plan = [p for p in pos if p["СС"] == ssname]
    pmap = {}
    for p in plan:
        pmap.setdefault(norm(p["ТекстСС"]), []).append(p)
    for e in xl[tag]["rows"]:
        if e["СерпеньСума"] <= 0.004:
            continue
        cand = pmap.get(norm(e["Назва"]))
        if not cand:
            unmapped.append((tag, e["Назва"], e["СерпеньСума"]))
            continue
        p = cand[0]
        on = p["ОН"] or "(без ЗН)"
        d = aug.setdefault(on, {"Сума": 0.0, "поз": []})
        d["Сума"] += e["СерпеньСума"]
        d["поз"].append("IRS %s: %s" % (tag, e["Назва"]))
        if unorm(e["Ед"]) != unorm(p["Ед"]):
            unit_bad.append({"Проєкт": "IRS " + tag, "Позиція": e["Назва"], "ЗН": on,
                             "ЕдExcel": e["Ед"], "Ед1С": p["Ед"],
                             "СумаСерпень": round(e["СерпеньСума"], 2)})

graf_tot = sum(d["Сума"] for d in aug.values())
print("Позицій графіка серпня зіставлено із ЗН: %d, сума %s грн" % (len(aug), "{:,.2f}".format(graf_tot)))
if unmapped:
    print("НЕ зіставлено: %d на %s грн" % (len(unmapped), "{:,.2f}".format(sum(u[2] for u in unmapped))))

rows = []
for on, a in aug.items():
    b = byON.get(on)
    plan_grn = b["ПланГрн"] if b else 0.0
    fact_grn = b["ФактГрн"] if b else 0.0
    ost_grn = b["Осталось"] if b else 0.0
    graf = round(a["Сума"], 2)
    over = round(max(graf - ost_grn, 0.0), 2)
    if plan_grn > 0 and fact_grn >= plan_grn:
        st = "ЗАКРИТО — не купувати"
    elif ost_grn <= 0.005:
        st = "ЗАКРИТО — не купувати"
    elif over > 0.005:
        st = "ЗМЕНШИТИ до залишку"
    else:
        st = "купувати за графіком"
    rows.append({"ЗН": on, "ГрафікСерпень": graf, "ПланГрн": plan_grn, "ФактГрн": fact_grn,
                 "ЗалишокГрн": ost_grn, "ЗайвеГрн": over,
                 "Викон%": round(fact_grn / plan_grn * 100, 1) if plan_grn else 0.0,
                 "Безпечно": round(min(graf, ost_grn), 2), "Статус": st,
                 "Позиції": "; ".join(a["поз"][:5])})

rows.sort(key=lambda r: -r["ЗайвеГрн"])
risk = [r for r in rows if r["ЗайвеГрн"] > 0.005]
print("\n" + "=" * 120)
print("ГРАФІК СЕРПНЯ %s грн ПРОТИ ЗАЛИШКУ ДО ЗАКУПІВЛІ %s грн  ->  надлишок %s грн"
      % ("{:,.2f}".format(graf_tot),
         "{:,.2f}".format(sum(r["ЗалишокГрн"] for r in rows)),
         "{:,.2f}".format(sum(r["ЗайвеГрн"] for r in risk))))
print("=" * 120)
h = "{:<34}{:>14}{:>14}{:>14}{:>14}{:>13}  {}"
print(h.format("Загальна назва", "Графік серп.", "План СС", "Вже куплено", "Залишок", "ЗАЙВЕ", "Статус"))
for r in risk:
    print(h.format(r["ЗН"][:32], "{:,.2f}".format(r["ГрафікСерпень"]), "{:,.2f}".format(r["ПланГрн"]),
                   "{:,.2f}".format(r["ФактГрн"]), "{:,.2f}".format(r["ЗалишокГрн"]),
                   "{:,.2f}".format(r["ЗайвеГрн"]), r["Статус"]))

ok = [r for r in rows if r["ЗайвеГрн"] <= 0.005]
print("\nКупувати за графіком без змін: %d ЗН на %s грн" % (len(ok), "{:,.2f}".format(sum(r["ГрафікСерпень"] for r in ok))))
print("Безпечний ліміт закупівлі серпня разом: %s грн" % "{:,.2f}".format(sum(r["Безпечно"] for r in rows)))

print("\n" + "=" * 120)
print("РОЗБІЖНІСТЬ ОДИНИЦЬ Excel vs СС (кількісне порівняння по них неможливе): %d позицій" % len(unit_bad))
print("=" * 120)
seen = set()
for u in sorted(unit_bad, key=lambda x: -x["СумаСерпень"]):
    k = (u["ЗН"], u["ЕдExcel"], u["Ед1С"])
    if k in seen:
        continue
    seen.add(k)
    print("  {:<32}{:<40} Excel: {:<10} 1С: {:<10} серп.: {:>12}".format(
        u["ЗН"][:30], u["Позиція"][:38], u["ЕдExcel"], u["Ед1С"], "{:,.2f}".format(u["СумаСерпень"])))

json.dump({"rows": rows, "unit_bad": unit_bad, "unmapped": unmapped},
          open("data_august_risk.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n-> data_august_risk.json")
