# -*- coding: utf-8 -*-
"""Аудит п.4 и п.10 (книга): план работ на квартал и бюджеты проектов из принятой книги экономиста.

1. «План  робіт-на Квартал_Жовтень  грудень.xlsx» (двойные пробелы в имени!):
   лист «відвантаження Мд План» — дома по проектам: план/факт по неделям внутри месяца, итог месяца
   «Всього (План/Факт)»; лист «Відвантаження Мк » — тонны по конструктивам.
2. «Принятый бюджет/Бюджет_Жовтень26-Грудень26.xlsx» — по листу проекта колонки бюджета проекта:
   «СС» (договор/лимит на весь проект), «Факт» (на 01.10), «Бюджет» (IV кв.) по строкам-статьям.
   Колонки ищутся по заголовку строки 4 (у листов МД ООН/IRC перед «Факт» стоят колонки по типам домов).
Вывод: data/plan_rabot.json, data/kniga_proekty.json.
"""
import re, openpyxl, warnings
from moy_byudzhet_common import *
warnings.filterwarnings("ignore")

PLAN = os.path.join(BASE_DIR, "План  робіт-на Квартал_Жовтень  грудень.xlsx")
KNIGA = os.path.join(BASE_DIR, "Принятый бюджет", "Бюджет_Жовтень26-Грудень26.xlsx")
МЕС = {"червень": "2026-06", "липень": "2026-07", "серпень": "2026-08", "вересень": "2026-09", "жовтень": "2026-10",
       "листопад": "2026-11", "грудень": "2026-12", "травень": "2026-05"}


def блоки_месяцев(ws, row_month=2, row_head=3):
    """[(месяц, колонка «Всього, т (План)», колонка «… (Факт)»)] по шапке листа."""
    out, cur = [], None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row_month, c).value
        if isinstance(v, str) and v.strip().lower() in МЕС:
            cur = МЕС[v.strip().lower()]
        h = ws.cell(row_head, c).value
        if cur and isinstance(h, str) and h.startswith("Всього") and "План" in h:
            out.append((cur, c, c + 1))
    return out


wb = openpyxl.load_workbook(PLAN, data_only=True)
plan = {"md": [], "mk": []}
ws = wb["відвантаження Мд План"]
бл = блоки_месяцев(ws)
for r in range(4, 32):
    имя = ws.cell(r, 2).value
    if not имя:
        continue
    rec = {"row": r, "project": str(имя).strip(), "type": str(ws.cell(r, 3).value or "").strip(),
           "qty": ws.cell(r, 4).value, "months": {}}
    for м, cp, cf in бл:
        p, f = ws.cell(r, cp).value or 0, ws.cell(r, cf).value or 0
        if p or f:
            rec["months"][м] = {"plan": float(p), "fact": float(f)}
    plan["md"].append(rec)
ws = wb["Відвантаження Мк "]
бл = блоки_месяцев(ws)
for r in range(4, 8):
    имя = ws.cell(r, 1).value
    if not имя:
        continue
    rec = {"row": r, "item": str(имя).strip(), "tonnes": ws.cell(r, 2).value, "months": {}}
    for м, cp, cf in бл:
        p, f = ws.cell(r, cp).value, ws.cell(r, cf).value
        p = float(p) if isinstance(p, (int, float)) else 0.0
        f = float(f) if isinstance(f, (int, float)) else 0.0
        if p or f:
            rec["months"][м] = {"plan": p, "fact": f}
        elif isinstance(ws.cell(r, cp).value, str):
            rec["months"][м] = {"plan_text": ws.cell(r, cp).value}
    plan["mk"].append(rec)
save("plan_rabot.json", plan)
print("План МД (дома: план/факт по месяцам):")
for x in plan["md"]:
    print(f"  {x['project']:<16} {x['type']:<14} {x['qty']}: " + ", ".join(f"{m[5:]}: {v['plan']:g}/{v['fact']:g}" for m, v in sorted(x["months"].items())))
print("План МК (тонны):")
for x in plan["mk"]:
    print(f"  {x['item']:<22} {x['tonnes']}: " + ", ".join(f"{m[5:]}: {v}" for m, v in sorted(x["months"].items())))

# --- книга экономиста: бюджеты проектов ---
ЛИСТЫ = {"ЦО_Производство": "Производство", "МД ООН 2026": "МД ООН 2026", "МД ООН АДМИН": "МД ООН Админ",
         "МД IRC 2026": "МД IRC 2026", "МД Вооз ЕМS Мала 2026": "МД Вооз ЕМS Мала 2026", "МД МХП ОРІЛЬ": "МД МХП ОРІЛЬ",
         "МК Каменец Цемент": "МК Каменец Цемент", "МК Кривий ріг": "МК Кривий ріг", "МК Чорноморськ": "МК Чорноморськ",
         "МД Кернел Кухня": "МД Кернел Кухня", "МД ООН 2025": "МД ООН 2025", "МК Глобино": "МК Глобино",
         "МХП санітарний вузол": "МХП санітарний вузол"}
wb = openpyxl.load_workbook(KNIGA, data_only=True)
kniga = {}
for лист, проект in ЛИСТЫ.items():
    ws = wb[лист]
    шапка = {c: str(ws.cell(4, c).value or "").strip() for c in range(14, 30)}
    кф = next((c for c, h in шапка.items() if h.lower().startswith("факт")), None)
    if кф is None:
        kniga[проект] = {"sheet": лист, "rows": [], "note": "нет колонок бюджета проекта"}
        continue
    кс, кб = кф - 1, кф + 1  # СС (итог по проекту) — слева от «Факт», «Бюджет» — справа
    домов = {шапка[c]: ws.cell(3, c).value for c in range(14, кф) if isinstance(ws.cell(3, c).value, (int, float))}
    rows = []
    for r in range(6, 156):
        имя = ws.cell(r, 1).value
        if not имя:
            continue
        def num(c):
            v = ws.cell(r, c).value
            return float(v) if isinstance(v, (int, float)) else 0.0
        ss, f, b = num(кс), num(кф), num(кб)
        if ss or f or b:
            rows.append({"row": r, "name": str(имя).strip(), "level": int(ws.row_dimensions[r].outline_level or 0),
                         "ss": ss, "fact": f, "budget": b,
                         "per_type": {шапка[c]: num(c) for c in range(14, кс) if num(c)}})
    kniga[проект] = {"sheet": лист, "cols": {"ss": кс, "fact": кф, "budget": кб}, "houses": домов, "rows": rows}
save("kniga_proekty.json", kniga)
print("\nКнига экономиста — проект: договор (СС «Поступление от Заказчика») / факт / бюджет IV кв.:")
for п, k in kniga.items():
    for x in k["rows"]:
        if x["name"].lower().startswith("поступление от заказчика"):
            print(f"  {п:<24} {fmt(x['ss']):>16} {fmt(x['fact']):>16} {fmt(x['budget']):>16}  остаток {fmt(x['ss'] - x['fact'])}")
