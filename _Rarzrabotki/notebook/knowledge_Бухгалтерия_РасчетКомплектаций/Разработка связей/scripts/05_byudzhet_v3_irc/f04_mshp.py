# -*- coding: utf-8 -*-
"""Замечание 1: состав МШП_h (выдачи на 30 м №1/№2/№3, Кт 22* или ОН «мшп») — нетто по позициям, ОН, норма 009, в G_h ли. Кэш = живые движения."""
import sys, json, re, collections
sys.stdout.reconfigure(encoding="utf-8")
OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\ekonomiya_ss_irc\out"
DATA = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\Мой расчет\data"
MV = json.load(open(OUT + r"\buh_moves.json", encoding="utf-8"))
CORE = json.load(open(OUT + r"\core.json", encoding="utf-8"))
F = json.load(open(DATA + r"\irc_fakt.json", encoding="utf-8"))
K = lambda s: " ".join(str(s or "").replace("\xa0", " ").split()).casefold()
is_mat = lambda a: (a or "")[:2] in ("20", "22", "28")
def house(name):
    m = re.search(r"(15|30) м №\s*(\d+)", name or "")
    return f"{m.group(1)}-{m.group(2)}" if (m and "IR" in (name or "")) else None
norm = {k for k, o in CORE["on"].items() if float(o.get("nr|30-1|ПланКол", 0)) > 0 or float(o.get("nr|30-1|ПланГрн", 0)) > 0}
gkeys = {it["key"] for it in F["deficit"]["items"]}
print(MV[0].keys())
agg = collections.defaultdict(lambda: collections.defaultdict(float))
info = {}
for r in MV:
    if is_mat(r["СчКт"]) and not is_mat(r["СчДт"]):
        h = house(r["СклКт"])
        if h in ("30-1", "30-2", "30-3") and ((r["СчКт"] or "").startswith("22") or K(r["ОНКт"]) == "мшп"):
            agg[r["НомКт"]][h] += float(r["Сумма"] or 0)
            info[r["НомКт"]] = (K(r["ОНКт"]), r["СчКт"])
tot = collections.defaultdict(float); grp = collections.defaultdict(lambda: collections.defaultdict(float))
for n, d in sorted(agg.items(), key=lambda x: -sum(x[1].values())):
    on, acc = info[n]
    tag = "НОРМА+G" if on in gkeys else ("НОРМА" if on in norm else "—")
    print(f"{n[:45]:<45} ОН={on[:22]:<22} сч={acc} {tag:<8} " + " ".join(f"{h}:{d.get(h,0):9.2f}" for h in ("30-1", "30-2", "30-3")))
    for h, v in d.items():
        tot[h] += v; grp[("мшп" if on == "мшп" else ("норма" if on in norm else "інше"))][h] += v
print("Σ нетто", dict(tot))
for g, d in grp.items():
    print(g, {h: round(v, 2) for h, v in d.items()})
