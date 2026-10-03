import sys, json, collections
sys.path.insert(0, r"..\svyaz")
sys.stdout.reconfigure(encoding="utf-8")
from views import *
V = json.load(open("v1.json", encoding="utf-8"))
RKH = {"00000000009": "30-1", "00000000011": "30-2", "00000000012": "30-3", "00000000013": "30-4"}
rk = collections.defaultdict(dict)
for r in V["rk"]:
    rk[RKH[r["Номер"]]][K(r["ОН"])] = r
for h in ["30-1", "30-2", "30-3"]:
    keys = set(rk[h]) | set(NOW[h])
    d_ost = d_n = d_over = 0; over_rk = over_me = 0; ovs_rk = 0
    for k in sorted(keys):
        a = rk[h].get(k); b = NOW[h].get(k)
        ro = a["Остаток"] if a else 0; mo = float(b["F"] + b["Sh"]) if b else 0
        rn = a["Норма"] if a else 0; mn = float(b["N"]) if b else 0
        rp = a["Понад"] if a else 0; mp = float(b["over_q"]) if b else 0
        over_rk += rp; over_me += mp; ovs_rk += a["ПонадСум"] if a else 0
        if abs(ro - mo) > 1e-3: d_ost += 1; print(h, "ОСТ≠", k, ro, mo)
        if abs(rn - mn) > 1e-3: d_n += 1; print(h, "НОРМА≠", k, rn, mn)
        if abs(rp - mp) > 1e-3: d_over += 1; print(h, "ПОНАД≠", k, rp, mp)
    print(h, "ОН:", len(keys), "расх Остаток:", d_ost, "Норма:", d_n, "Понад:", d_over, "ΣПонадСум РК", round(ovs_rk, 2),
          "ΣПонад брутто моё", round(sum(float(v["over_g"]) for v in NOW[h].values()), 2))
