# -*- coding: utf-8 -*-
"""Payload графіка оплат для документа А_ГрафикОплатМатериалов.

З МАТЕРИАЛИ.xlsx (аркуші IRS 15 / IRS 30) робимо рядки
(Підрозділ, СС, НоменклатураСС, ОбщееНазвание, Місяць, Кількість, Сума).

Ключ зіставлення тексту Excel -> рядок СС -> ОН: нормалізоване ім'я.
Той самий алгоритм, що дав 177/177 попадань при звірці графіка серпня.
"""
import json, sys, re
sys.stdout.reconfigure(encoding='utf-8')

PODR = "МД IRS 2026"
SSNAME = {"15": "МД IRS 2026 15 м (МД IRS 2026)", "30": "МД IRS 2026 30 м (МД IRS 2026)"}
MONTHS = {"липень": "2026-07-01", "серпень": "2026-08-01", "вересень": "2026-09-01"}


def norm(s):
    s = (s or "").lower().replace("ё", "е").replace("*", "х").replace("×", "х")
    s = s.replace(",", ".").replace("ʼ", "'").replace("’", "'")
    s = re.sub(r"[\s ]+", " ", s)
    return re.sub(r"[^0-9a-zа-яіїєґ'.х/]+", " ", s).strip()


pos = json.load(open("data_positions.json", encoding="utf-8"))["plan"]
xl = json.load(open("data_excel_koshtoris.json", encoding="utf-8"))

rows, unmapped = [], []
for tag, ss in SSNAME.items():
    plan = [p for p in pos if p["СС"] == ss]
    pmap = {}
    for p in plan:
        pmap.setdefault(norm(p["ТекстСС"]), []).append(p)
    for e in xl[tag]["rows"]:
        cand = pmap.get(norm(e["Назва"]))
        on = cand[0]["ОН"] if cand else ""
        for mname, mdate in MONTHS.items():
            kol = e[mname.capitalize() + "Кол"] if (mname.capitalize() + "Кол") in e else 0.0
            summ = e[mname.capitalize() + "Сума"] if (mname.capitalize() + "Сума") in e else 0.0
            if abs(summ) < 0.005 and abs(kol) < 0.0005:
                continue
            if not cand:
                unmapped.append({"СС": ss, "Назва": e["Назва"], "Місяць": mdate, "Сума": summ})
                continue
            rows.append({"Подразделение": PODR, "СС": ss, "НоменклатураСС": e["Назва"].strip(),
                         "ОбщееНазвание": on, "Месяц": mdate,
                         "Количество": round(kol, 3), "Сумма": round(summ, 2)})

print("рядків графіка: %d" % len(rows))
print("не зіставлено: %d на %.2f грн" % (len(unmapped), sum(u["Сума"] for u in unmapped)))
for u in unmapped[:10]:
    print("   %s | %s | %.2f" % (u["СС"][:22], u["Назва"][:44], u["Сума"]))

by_month = {}
for r in rows:
    by_month[r["Месяц"]] = by_month.get(r["Месяц"], 0.0) + r["Сумма"]
print("\nпо місяцях:")
for m in sorted(by_month):
    print("   %s : %14s" % (m, "{:,.2f}".format(by_month[m])))
print("   %-10s %14s" % ("РАЗОМ", "{:,.2f}".format(sum(by_month.values()))))

no_on = [r for r in rows if not r["ОбщееНазвание"]]
print("\nрядків без ОН: %d на %.2f" % (len(no_on), sum(r["Сумма"] for r in no_on)))

json.dump({"rows": rows, "unmapped": unmapped, "by_month": by_month},
          open("data_graph_payload.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n-> data_graph_payload.json")
