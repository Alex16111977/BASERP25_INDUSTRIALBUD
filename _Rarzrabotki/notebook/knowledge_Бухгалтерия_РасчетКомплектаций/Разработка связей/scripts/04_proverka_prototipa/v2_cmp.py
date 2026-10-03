import sys, json, re, collections
sys.path.insert(0, r"..\svyaz")
sys.stdout.reconfigure(encoding="utf-8")
from base import *
V = json.load(open("v1.json", encoding="utf-8"))
def hcls(s):
    if s == "ІБ _МД IRC 2026": return "P"
    m = re.search(r"30 м №\s*(\d+)", s); return "30-" + m.group(1) if m else "?"
MAT = ("20", "22", "28")
F = collections.defaultdict(lambda: [0.0, 0.0]); Ftot = collections.defaultdict(float)
for r in V["obor"]:
    h = hcls(r["Скл"])
    if h == "P" or (r["Кор"] or "")[:2] in MAT: continue
    F[(r["Код"], h)][0] += r["КолКт"]; F[(r["Код"], h)][1] += r["СумКт"]; Ftot[h] += r["СумКт"]
S = collections.defaultdict(lambda: [0.0, 0.0]); Stot = collections.defaultdict(float)
for r in V["ost"]:
    h = hcls(r["Скл"]); S[(r["Код"], h)][0] += r["Кол"]; S[(r["Код"], h)][1] += r["Сум"]; Stot[h] += r["Сум"]
# сравнение итогов: моё (нетто) vs base cons/st (нетто)
for h in ["30-1", "30-2", "30-3", "30-4"]:
    bf = sum(float(cons[c][h][1]) for c in cons); bs = sum(float(st[c]["H:" + h][1]) for c in st)
    bfg = sum(float(cons[c][h][2]) for c in cons); bsg = sum(float(st[c]["H:" + h][2]) for c in st)
    print(h, f"списано нетто: моё {Ftot[h]:.2f} base {bf:.2f} (брутто base {bfg:.2f}) | склад нетто: моё {Stot[h]:.2f} base {bs:.2f} (брутто {bsg:.2f})")
bp = sum(float(st[c]["P"][1]) for c in st); print("P склад нетто: моё", round(Stot["P"], 2), "base", round(bp, 2))
# по карточкам — расхождения
bad = 0
for c in set(k for k, _ in F) | set(cons):
    for h in ["30-1", "30-2", "30-3"]:
        a = F.get((c, h), [0, 0]); b = cons[c][h] if c in cons else [0, 0, 0]
        if abs(a[0] - float(b[0])) > 1e-6 or abs(a[1] - float(b[1])) > 0.01:
            bad += 1; print("F≠", c, h, a, [float(x) for x in b[:2]], CARD.get(c, {}).get("name"))
for c in set(k for k, _ in S) | set(st):
    for h in ["30-1", "30-2", "30-3", "P"]:
        a = S.get((c, h), [0, 0]); key = "P" if h == "P" else "H:" + h
        b = st[c][key] if c in st else [0, 0, 0]
        if abs(a[0] - float(b[0])) > 1e-6 or abs(a[1] - float(b[1])) > 0.01:
            bad += 1; print("S≠", c, h, a, [float(x) for x in b[:2]], CARD.get(c, {}).get("name"))
print("расхождений по карточкам:", bad)
POS = {"00000005296": "гайка М12", "00000011364": "гайка М48", "00-00009980": "плінтус на стільницю", "00-00016723": "S-8",
       "00000007615": "труба 100*100*3", "00-00021697": "держак ринви", "00-00010526": "днище", "00000017651": "кільце", "00-00005944": "кільце2",
       "00-00006019": "люк", "00-00012879": "плита", "00-00017426": "плита2", "00-00023184": "КОМПЛЕКТ септика"}
for c, nm in POS.items():
    print(c, nm, "F:", {h: F.get((c, h), [0, 0]) for h in ["30-1", "30-2", "30-3"]}, "S:", {h: S.get((c, h), [0, 0]) for h in ["30-1", "30-2", "30-3", "P"]})
