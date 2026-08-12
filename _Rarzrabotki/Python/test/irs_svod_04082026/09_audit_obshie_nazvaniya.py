# -*- coding: utf-8 -*-
"""Аудит ОН: сироты-факт, сироты-план, ценовые разрывы."""
import sys, json
sys.stdout.reconfigure(encoding='utf-8')
mat  = json.load(open("data_materials.json",encoding="utf-8"))
fact = json.load(open("data_fact_nom.json",encoding="utf-8"))
byON = {d["ОН"]: d for d in mat["byON"]}

orph_fact = {k:v for k,v in byON.items() if v["ПланГрн"]==0 and v["ФактГрн"]>0}
orph_plan = {k:v for k,v in byON.items() if v["ПланГрн"]>0 and v["ФактГрн"]==0}
both      = {k:v for k,v in byON.items() if v["ПланГрн"]>0 and v["ФактГрн"]>0}

print(f"ОН всего={len(byON)}  обе стороны={len(both)}  сироты-ФАКТ={len(orph_fact)}  сироты-ПЛАН={len(orph_plan)}")
print(f"Σ сирот-факт = {sum(v['ФактГрн'] for v in orph_fact.values()):,.2f}")
print(f"Σ сирот-план = {sum(v['ПланГрн'] for v in orph_plan.values()):,.2f}\n")

print("="*110); print("СИРОТЫ-ФАКТ (закуплено, но такого ОН в плане СС нет) — карточки внутри:")
nom_by_on={}
for x in fact: nom_by_on.setdefault(x["ОН"],[]).append(x)
for k,v in sorted(orph_fact.items(), key=lambda kv:-kv[1]["ФактГрн"]):
    print(f"\n  ОН «{k}»  факт={v['ФактГрн']:,.2f}  кол={v['ФактКол']:g}  ціна={v['ЦенаФакт']:,.2f}")
    for x in nom_by_on.get(k,[]):
        print(f"      {x['Код']:<14}{x['Имя'][:46]:<48}{x['Кол']:>10g} {x['Ед']:<8}{x['Сумма']:>12,.2f}  ціна={x['Цена']:,.2f}")

print("\n"+"="*110); print("СИРОТЫ-ПЛАН (план есть, закупок нет) ТОП-20 по сумме — кандидаты-приёмники:")
for k,v in sorted(orph_plan.items(), key=lambda kv:-kv[1]["ПланГрн"])[:20]:
    print(f"  {k[:52]:<54} план={v['ПланГрн']:>13,.2f} кол={v['ПланКол']:>9g} ціна={v['ЦенаПлан']:>11,.2f}")

print("\n"+"="*110); print("ЦЕНОВОЙ РАЗРЫВ >×3 внутри ОН (есть и план, и факт):")
gaps=[]
for k,v in both.items():
    cp,cf = v["ЦенаПлан"], v["ЦенаФакт"]
    if cp>0 and cf>0:
        rel = cf/cp if cf>cp else cp/cf
        if rel>3: gaps.append((rel,k,v))
for rel,k,v in sorted(gaps, key=lambda t:-t[0]):
    print(f"  ×{rel:>6.1f}  {k[:44]:<46} цінаПлан={v['ЦенаПлан']:>11,.2f}  цінаФакт={v['ЦенаФакт']:>11,.2f}  факт={v['ФактГрн']:>12,.2f}")
    for x in nom_by_on.get(k,[]):
        print(f"           └ {x['Код']:<14}{x['Имя'][:42]:<44}{x['Цена']:>11,.2f}")

json.dump({"orph_fact":orph_fact,"orph_plan":orph_plan,
           "gaps":[{"rel":round(r,2),"ОН":k,**v} for r,k,v in gaps],
           "nom_by_on":nom_by_on},
          open("data_on_audit.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n-> data_on_audit.json")
