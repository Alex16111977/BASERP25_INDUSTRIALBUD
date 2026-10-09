# -*- coding: utf-8 -*-
import sys, json
sys.stdout.reconfigure(encoding="utf-8")
R = json.load(open("layer_0710.json", encoding="utf-8"))
def f(v): 
    s = "{:,.0f}".format(v).replace(",", " ")
    return "0" if s in ("-0", "0") else s
SPEC = ["Сигналізація виявлення диму", "Двері МДФ внутрішні", "Кутики внутрішні", "Комплект септика", "Гайка",
        "Профільна труба", "Лист ГК (металопрокат)", "лист пвл"]
for P, d in R.items():
    res = d["res"]
    for title, lst in (("економії", d["econ"]), ("перевитрати", d["over"])):
        print("\n**%s — ТОП-15 %s**\n" % (P, title))
        print("| Група | L_g | груп. ціна / обсяг | V1 ΣL_c ціна / обсяг | інші | R_g | перенос V2 | R2 (V2) | карток | кл. гр→V1→V2 |")
        print("|---|--:|--:|--:|--:|--:|--:|--:|--:|---|")
        for g in lst:
            v = res[g]
            ch = v["cls_g"] + "→" + v["cls_c"] + "→" + v["cls_c2"]
            if v["cls_g"] != v["cls_c"] or v["cls_g"] != v["cls_c2"]: ch = "**" + ch + "**"
            het = "" if not v["hetero"] else " (" + ",".join(v["hetero"]) + ")"
            print("| %s%s | %s | %s / %s | %s / %s | %s | %s | %s | %s | %d | %s |" % (
                g, het, f(v["L"]), f(v["gp"]), f(v["gq"]), f(v["pr"]), f(v["qt"]), f(v["others"]), f(v["R"]), f(v["T"]), f(v["R2"]), v["ncards"], ch))
    extra = [g for g in res if any(g.lower() == s.lower() for s in SPEC) and g not in d["econ"] + d["over"]]
    if extra:
        print("\n**%s — особые группы вне ТОП**\n" % P)
        for g in extra:
            v = res[g]
            print("- %s: L_g %s; груп. %s/%s; V1 %s/%s; інші %s; R %s; карток %d" % (g, f(v["L"]), f(v["gp"]), f(v["gq"]), f(v["pr"]), f(v["qt"]), f(v["others"]), f(v["R"]), v["ncards"]))
    # статистика по ТОП
    top = d["econ"] + d["over"]
    het = [g for g in top if res[g]["hetero"]]
    chg = [g for g in top if res[g]["cls_g"] != res[g]["cls_c"]]
    chg2 = [g for g in top if res[g]["cls_g"] != res[g]["cls_c2"]]
    one = [g for g in top if res[g]["ncards"] == 1 and abs(res[g]["gp"] - res[g]["pr"]) > 1]
    print("\nСтат %s: в ТОП %d групп; неоднородных %d; смена класса V1 %d %s; V2 %d %s; однокарточных с факторами ≠ групповых %d %s" % (
        P, len(top), len(het), len(chg), chg, len(chg2), chg2, len(one), one))
    print("Σ|R_g| в ТОП %.2f из Σ|L_g| %.2f" % (sum(abs(res[g]["R"]) for g in top), sum(abs(res[g]["L"]) for g in top)))
