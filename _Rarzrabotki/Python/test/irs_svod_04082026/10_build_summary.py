# -*- coding: utf-8 -*-
import json, sys
sys.stdout.reconfigure(encoding='utf-8')
mat=json.load(open("data_materials.json",encoding="utf-8"))
byON=mat["byON"]; tot=mat["tot"]
for d in byON:
    d["Прогноз"]=round(d["ФактГрн"]+d["Осталось"],2)
    d["Відхилення"]=round(d["ПланГрн"]-d["Прогноз"],2)
eco=[d for d in byON if d["Відхилення"]>0.005 and d["ПланГрн"]>0]
per=[d for d in byON if d["Відхилення"]<-0.005]
print(f"ПЛАН={tot['ПланГрн']:,.2f}  ФАКТ={tot['ФактГрн']:,.2f}  ЗАЛИШОК={tot['Осталось']:,.2f}")
print(f"ПРОГНОЗ={tot['Прогноз']:,.2f}  ВІДХИЛЕННЯ={tot['ОтклонениеПрогноз']:,.2f} ({tot['Прогноз']/tot['ПланГрн']*100:.1f}% плану)")
print(f"\nекономія: {len(eco)} позицій на {sum(d['Відхилення'] for d in eco):,.2f}")
print(f"перевитрата: {len(per)} позицій на {sum(-d['Відхилення'] for d in per):,.2f}")
print("\nТОП-10 ЕКОНОМІЯ:")
for d in sorted(eco,key=lambda x:-x["Відхилення"])[:10]:
    print(f"  {d['ОН'][:40]:<42}план={d['ПланГрн']:>12,.2f} прогноз={d['Прогноз']:>12,.2f} економія={d['Відхилення']:>11,.2f} ({d['Прогноз']/d['ПланГрн']*100:>5.1f}%)")
print("\nТОП-10 ПЕРЕВИТРАТА:")
for d in sorted(per,key=lambda x:x["Відхилення"])[:10]:
    pl=d['ПланГрн'] or 0
    pct = f"{d['Прогноз']/pl*100:>5.1f}%" if pl else "  н/д"
    print(f"  {d['ОН'][:40]:<42}план={pl:>12,.2f} прогноз={d['Прогноз']:>12,.2f} перевитр={-d['Відхилення']:>11,.2f} ({pct})")
print("\nТОП-15 ЩЕ ЗАКУПИТИ (залишок плану):")
for d in sorted(byON,key=lambda x:-x["Осталось"])[:15]:
    print(f"  {d['ОН'][:40]:<42}залишок={d['Осталось']:>12,.2f}  план={d['ПланГрн']:>12,.2f} факт={d['ФактГрн']:>12,.2f} ({d['Процент']:>5.1f}%)")
json.dump(byON, open("data_byon_calc.json","w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n-> data_byon_calc.json")
