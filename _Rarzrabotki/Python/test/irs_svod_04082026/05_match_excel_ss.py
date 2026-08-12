# -*- coding: utf-8 -*-
"""Сверка кошториса Excel <-> план СС по ключу НоменклатураСС."""
import sys, json, re
sys.stdout.reconfigure(encoding='utf-8')
pos = json.load(open("data_positions.json",encoding="utf-8"))
xl  = json.load(open("data_excel_koshtoris.json",encoding="utf-8"))

def norm(s):
    s = (s or "").lower().replace("ё","е").replace("*","х").replace("×","х")
    s = s.replace(",",".").replace("ʼ","'").replace("’","'")
    s = re.sub(r"[\s\u00a0]+"," ", s)
    return re.sub(r"[^0-9a-zа-яіїєґ'.х/]+"," ", s).strip()

SS = {"15":"МД IRS 2026 15 м (МД IRS 2026)","30":"МД IRS 2026 30 м (МД IRS 2026)"}
report={}
for tag, ssname in SS.items():
    plan = [p for p in pos["plan"] if p["СС"]==ssname]
    exc  = xl[tag]["rows"]
    pmap = {}
    for p in plan: pmap.setdefault(norm(p["ТекстСС"]), []).append(p)
    used=set(); pairs=[]; only_xl=[]
    for e in exc:
        k = norm(e["Назва"])
        cand = pmap.get(k)
        if cand:
            p = cand.pop(0); used.add(id(p))
            pairs.append((e,p))
        else:
            only_xl.append(e)
    matched_ids = {id(p) for _,p in pairs}
    only_pl = [p for p in plan if id(p) not in matched_ids]
    diffs=[]
    for e,p in pairs:
        dk = round(e["КолДом"]-p["КолДом"],4); dc = round(e["Цена"]-p["Цена"],2); ds = round(e["СумаДом"]-p["СуммаДом"],2)
        if abs(dk)>0.001 or abs(dc)>0.01 or abs(ds)>0.01:
            diffs.append({"Назва":e["Назва"],"Ед":e["Ед"],
                          "КолExcel":e["КолДом"],"Кол1С":p["КолДом"],"ΔКол":dk,
                          "ЦенаExcel":e["Цена"],"Цена1С":p["Цена"],"ΔЦена":dc,
                          "СумаExcel":e["СумаДом"],"Сума1С":p["СуммаДом"],"ΔСума":ds})
    report[tag]={"позExcel":len(exc),"поз1С":len(plan),"сопоставлено":len(pairs),
                 "тільки_в_Excel":only_xl,"тільки_в_1С":only_pl,"розбіжності":diffs,
                 "ΔСумаДом":round(sum(d["ΔСума"] for d in diffs),2),
                 "Домов":xl[tag]["домов"]}
    r=report[tag]
    print(f"=== IRS {tag} === Excel={len(exc)} поз., 1С={len(plan)} поз., сопоставлено={len(pairs)}")
    print(f"    тільки в Excel: {len(only_xl)} | тільки в 1С: {len(only_pl)} | розбіжностей у парах: {len(diffs)}")
    print(f"    Δ суми/дім по парах = {r['ΔСумаДом']:,.2f}  → на {r['Домов']} домів = {r['ΔСумаДом']*r['Домов']:,.2f}")
    for d in sorted(diffs,key=lambda x:-abs(x['ΔСума']))[:8]:
        print(f"      {d['Назва'][:44]:<46} Excel {d['СумаExcel']:>11,.2f} | 1С {d['Сума1С']:>11,.2f} | Δ {d['ΔСума']:>10,.2f}")
    for e in only_xl[:6]: print(f"      [тільки Excel] {e['Назва'][:50]:<52} {e['СумаДом']:>11,.2f}")
    for p in only_pl[:6]: print(f"      [тільки 1С]    {p['ТекстСС'][:50]:<52} {p['СуммаДом']:>11,.2f}")
    print()
json.dump(report, open("data_match.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("-> data_match.json")
