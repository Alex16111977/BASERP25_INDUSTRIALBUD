import sys, json, collections
sys.path.insert(0, r"..\svyaz")
sys.stdout.reconfigure(encoding="utf-8")
from views import *
for h in ["30-1", "30-2", "30-3"]:
    n0 = sum(float(v["over_g"]) for k, v in NOW[h].items() if v["N"] == 0)
    np_ = sum(float(v["over_g"]) for k, v in NOW[h].items() if v["N"] > 0)
    top = sorted(((float(v["over_g"]), k) for k, v in NOW[h].items() if v["N"] == 0), reverse=True)[:6]
    print(h, "понад при N=0:", round(n0, 2), "при N>0:", round(np_, 2), top)
# shahmatka D also on done houses
extra = 0
for k, v in NOW_BUY.items():
    pass
# пересчёт докупить, если Треба считать и по №1–№2 (как в отчёте)
tot_now, tot_alt = Z, Z
for k in ON_KEYS:
    N, Ns = ON_NORM.get(k, [Z, Z])
    if N <= 0: continue
    sp = max(Z, NOW_SP[k][0]); pr = Ns / N
    sd = sum((NOW[h][k]["D"] if k in NOW[h] else N) for h in TODO)
    sd12 = sum((NOW[h][k]["econ_q"] if k in NOW[h] else N) for h in DONE)
    tot_now += max(Z, sd - sp) * pr; tot_alt += max(Z, sd + sd12 - sp) * pr
print("докупить now", round(float(tot_now), 2), "если Треба и по №1–№2:", round(float(tot_alt), 2))
