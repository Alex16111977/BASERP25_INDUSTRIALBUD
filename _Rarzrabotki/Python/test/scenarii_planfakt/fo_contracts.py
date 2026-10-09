# -*- coding: utf-8 -*-
"""Финотчёт 09.10: по каждому договору — к получению (K) и подписанная часть (без поданных актов H)."""
import sys, os, re, json
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dump import load

P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\09102026\Финансовый_отчет Производства 09-10-2026 с расшифровкой.xlsx"
wf, wv = load(P, False), load(P, True)
ws, wsv = wf.worksheets[0], wv.worksheets[0]


def v(c):
    x = wsv[c].value
    return float(x) if isinstance(x, (int, float)) else 0.0


obj = None
res = {}
for r in range(6, ws.max_row + 1):
    a = ws["A%d" % r].value
    c = ws["C%d" % r].value
    if a and isinstance(c, str) and re.fullmatch(r"=C%d" % (r + 1), c):
        obj = a.strip()
        res[obj] = {"K": v("K%d" % r), "H": v("H%d" % r), "G": v("G%d" % r), "I": v("I%d" % r), "J": v("J%d" % r),
                    "contracts": []}
        continue
    k = ws["K%d" % r].value
    if obj and isinstance(k, str) and "MAX(0" in k.upper():
        j = ws["J%d" % r].value
        j_from_h = isinstance(j, str) and "H" in j.upper()
        G, I, J, H = v("G%d" % r), v("I%d" % r), v("J%d" % r), v("H%d" % r)
        pess = round(max(0.0, G + I + (0.0 if j_from_h else J)), 2)
        res[obj]["contracts"].append({"row": r, "name": a, "K": v("K%d" % r), "G": G, "I": I, "J": J, "H": H,
                                      "Kformula": k, "Jformula": j, "pess": pess})
for o, d in res.items():
    pess = round(sum(c["pess"] for c in d["contracts"]), 2)
    d["pess"] = pess
    print("%-75s K=%14.2f H=%13.2f G=%14.2f I=%13.2f J=%11.2f | песс=%13.2f" % (o[:75], d["K"], d["H"], d["G"], d["I"], d["J"], pess))
    for c in d["contracts"]:
        print("    r%-4d %-55s K=%12.2f G=%13.2f I=%12.2f J=%10.2f H=%12.2f песс=%12.2f  %s" % (
            c["row"], str(c["name"])[:55], c["K"], c["G"], c["I"], c["J"], c["H"], c["pess"], c["Kformula"]))
json.dump(res, open(__file__.replace(".py", ".json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
