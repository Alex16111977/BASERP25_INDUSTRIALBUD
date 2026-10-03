import sys
sys.path.insert(0, r"..\svyaz")
sys.stdout.reconfigure(encoding="utf-8")
from views import *
for r in MV:
    if KIT in (r["КодДт"], r["КодКт"]) or (r["КодКт"] in KIT_CARDS and wcls(r["СклКт"]).startswith("H:30")) or (r["КодДт"] in KIT_CARDS and wcls(r["СклДт"]).startswith("H:30")):
        print(r["Период"], r["Рег"][:45], r["СчДт"], r["СчКт"], r["КодДт"], r["СклДт"][:30], r["КодКт"], r["СклКт"][:30], r["Сумма"], r["КолДт"], r["КолКт"])
print("ASM", {k: [float(x) for x in v] for k, v in ASM.items()})
print("KIT card on/name", CARD[KIT])
for c in KIT_CARDS: print(c, CARD.get(c))
for h in H30[:3]:
    print(h, "MAP septic", {k: (float(v) if not isinstance(v, dict) else v) for k, v in MAPPED[h][SEPTIC_LINE].items()})
    k = LINE[SEPTIC_LINE]["k"]; print(h, "NOW septic ОН", k, {kk: float(v) for kk, v in NOW[h].get(k, {}).items()})
print("NOW_SP septic", NOW_SP.get(LINE[SEPTIC_LINE]["k"]), "ON_CARDS", ON_CARDS.get(LINE[SEPTIC_LINE]["k"]))
b = MAP_BUY[SEPTIC_LINE]; print({k: v for k, v in b.items() if k != "byh"}); 
for h, v in b["byh"].items(): print(h, v)
print("NOW_BUY septic", {k: v for k, v in NOW_BUY[LINE[SEPTIC_LINE]["k"]].items() if k != "byh"})
