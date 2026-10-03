import sys, json, collections
sys.path.insert(0, r"..\svyaz")
sys.stdout.reconfigure(encoding="utf-8")
from views import *
for cls in ["F:30-1", "F:30-2", "F:30-3", "S:30-1", "S:30-2", "S:30-3", "P"]:
    lines, ex, grp = contributions(cls)
    s_lines = sum(g for v in lines.values() for _, g, _ in v)
    s_ex = sum(v[1] for d in ex.values() for v in d.values())
    s_grp = sum(g for _, g in grp)
    comp, kits = kit_state(cls)
    s_kit = sum(v[1] for v in comp.values())
    total = sum(qg(c, cls)[1] for c in U if c != KIT)
    asm = sum(ASM[(c, cls[2:])][1] for c in KIT_CARDS) if cls.startswith("F:") else Z
    kitcard = qg(KIT, cls)[1] if KIT in U else Z
    # карточки, учтённые дважды: в линиях + в kit
    dup = [c for c in KIT_CARDS if any(cc == c for v in lines.values() for *_, cc in v)]
    print(cls, "Σвиды", round(float(s_lines + s_ex + s_grp + s_kit), 2), "Σкарточек(без KIT)+ASM", round(float(total + asm), 2),
          "KIT карточка", round(float(kitcard), 2), "kits", float(kits), "дубль комплектующих в строках:", dup)
# N:1 / одна карточка в нескольких строках
cnt = collections.Counter()
for c, m in MAP.items():
    for l in m["links"]: cnt[c] += 1
print("карточки с >1 связью:", {c: n for c, n in cnt.items() if n > 1})
# сумма по строкам с N:1 — сверка с суммой карточек
for n, lst in BY_LINE.items():
    if len(lst) > 1:
        cs = [c for c, _, _ in lst]
        for h in ["30-1", "30-2"]:
            r = MAPPED[h].get(n)
            if r is None: continue
            fsum = sum(qg(c, "F:" + h)[0] * [l for l in MAP[c]["links"] if l["line"] == n][0]["coef"] for c in cs if len(MAP[c]["links"]) == 1)
            if abs(fsum - r["F"]) > Decimal("0.001") and not any(len(MAP[c]["links"]) > 1 for c in cs):
                print("N:1 F≠", n, h, float(fsum), float(r["F"]))
print("ok")
