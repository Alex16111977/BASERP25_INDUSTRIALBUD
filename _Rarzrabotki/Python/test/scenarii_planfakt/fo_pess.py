# -*- coding: utf-8 -*-
import sys, os, re, json
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dump import load
P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\09102026\Финансовый_отчет Производства 09-10-2026 с расшифровкой.xlsx"
wsv = load(P, True).worksheets[0]
res = json.load(open("fo_contracts.json", encoding="utf-8"))
def ev(f, row, hzero):
    def rep(m):
        col, r = m.group(1), int(m.group(2))
        if col == "H" and hzero: return "0"
        x = wsv["%s%d" % (col, r)].value
        return repr(float(x) if isinstance(x, (int, float)) else 0.0)
    e = re.sub(r"\b([A-Z])(\d+)\b", rep, f[1:])
    e = e.replace("ROUND", "round").replace("MAX", "max").replace("MIN", "min")
    return eval(e)
for o, d in res.items():
    s0 = s1 = 0
    for c in d["contracts"]:
        k_full = ev(c["Kformula"], c["row"], False); k_signed = ev(c["Kformula"], c["row"], True)
        s0 += k_full; s1 += k_signed
        if abs(k_full - c["K"]) > 0.01: print("  !!! пересчёт K не сходится", o, c["row"], k_full, c["K"])
    d["opt"], d["pess"] = round(s0, 2), round(s1, 2)
    print("%-60s опт(K)=%14.2f  песс(без поданных актов)=%14.2f" % (o[:60], s0, s1))
json.dump(res, open("fo_contracts.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
