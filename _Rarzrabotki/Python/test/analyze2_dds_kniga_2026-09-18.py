# -*- coding: utf-8 -*-
"""Проверки, умеющие провалиться, над дампами книги и базы (без COM, без записи).

  1. По каждому листу с блоками и месяцу: сумма уровня 2 (бн/нал) == ПОСТУПЛЕНИЯ(r6) + СПИСАНИЕ(r19).
  2. СВОД_Производство == ЦО_Производство + 14 объектных листов, по месяцам, бн и нал.
  3. Карта «имя статьи книги → код статьи ДДС» (все имена уровня 2 с деньгами).
  4. Строка «Суточные», статьи с «юр» в справочнике, комментарии IRC, хвост IRC.
Счётчик: выполнено N / провалов M; код возврата 1 при M > 0.
"""
import sys, json, os
sys.stdout.reconfigure(encoding='utf-8')
from collections import defaultdict

D = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
kniga = json.load(open(os.path.join(D, "dds_kniga_dump.json"), encoding="utf-8"))
baza = json.load(open(os.path.join(D, "dds_baza_dump.json"), encoding="utf-8"))
an = json.load(open(os.path.join(D, "dds_analysis.json"), encoding="utf-8"))
sheets = {s["name"]: s for s in kniga["sheets"]}


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").replace("\t", " ").replace("\n", " ").replace("\r", " ")
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip().lower()


def fmt(x):
    return f"{x:,.2f}".replace(",", " ")


N = 0
M = 0


def check(cond, msg):
    global N, M
    N += 1
    if not cond:
        M += 1
        print(f"  FAIL: {msg}")
    return cond


OBJECTS = ["МД ООН 2026", "МД ООН АДМИН", "МД IRC 2026", "МД Вооз ЕМS Мала 2026", "МД МХП ОРІЛЬ",
           "МД ПРООН Черкаси  ДСНС", "МК Підгірці", "МК Каменец Цемент", "МК Кривий ріг", "МК Чорноморськ",
           "МД ВООЗ  2025", "МД Кернел Кухня", "МД ООН 2025", "МК Глобино"]
CO = "ЦО_Производство"
SVOD = "СВОД_Производство"

print("=" * 100)
print("1. ТОЖДЕСТВО: уровень 2 == ПОСТУПЛЕНИЯ + СПИСАНИЕ (по каждому листу с блоками, каждому месяцу, бн и нал)")
for name, sh in sheets.items():
    if not sh.get("blocks") or name.startswith("Date"):
        continue
    rows = {x["r"]: x for x in sh["rows"]}
    et = an["etalon"].get(name, {})
    r6, r19 = rows.get(6), rows.get(19)
    if not (r6 and r19 and norm(r6["name"]) == "поступления" and norm(r19["name"]) == "списание"):
        print(f"  ! «{name}»: строки 6/19 не ПОСТУПЛЕНИЯ/СПИСАНИЕ: {r6 and r6['name']} / {r19 and r19['name']}")
        continue
    for m, e in et.items():
        exp_bn = r6["vals"][m]["bn"] + r19["vals"][m]["bn"]
        exp_nal = r6["vals"][m]["nal"] + r19["vals"][m]["nal"]
        ok_bn = check(abs(e["bn"] - exp_bn) < 0.02, f"«{name}» {m} бн: L2 {fmt(e['bn'])} ≠ ПОСТ+СПИС {fmt(exp_bn)}")
        ok_nal = check(abs(e["nal"] - exp_nal) < 0.02, f"«{name}» {m} нал: L2 {fmt(e['nal'])} ≠ ПОСТ+СПИС {fmt(exp_nal)}")
        tag = "ok" if ok_bn and ok_nal else "FAIL"
        print(f"  {tag:<4} «{name}» {m}: L2 бн {fmt(e['bn']):>16} нал {fmt(e['nal']):>14} | ПОСТ бн {fmt(r6['vals'][m]['bn']):>14} нал {fmt(r6['vals'][m]['nal']):>13} | СПИС бн {fmt(r19['vals'][m]['bn']):>14} нал {fmt(r19['vals'][m]['nal']):>13}")

print("\n" + "=" * 100)
print("2. СВОД_Производство == ЦО_Производство + 14 объектных (уровень 2, по месяцам)")
svod = an["etalon"][SVOD]
for m in svod:
    s_bn = s_nal = 0.0
    parts = []
    for nm in [CO] + OBJECTS:
        e = an["etalon"][nm].get(m)
        if e:
            s_bn += e["bn"]; s_nal += e["nal"]
            if e["bn"] or e["nal"]:
                parts.append(nm)
    check(abs(s_bn - svod[m]["bn"]) < 0.05, f"СВОД {m} бн {fmt(svod[m]['bn'])} ≠ ЦО+объекты {fmt(s_bn)}")
    check(abs(s_nal - svod[m]["nal"]) < 0.05, f"СВОД {m} нал {fmt(svod[m]['nal'])} ≠ ЦО+объекты {fmt(s_nal)}")
    print(f"  {m}: СВОД бн {fmt(svod[m]['bn'])} / нал {fmt(svod[m]['nal'])}; ЦО+объекты бн {fmt(s_bn)} / нал {fmt(s_nal)}; листов с деньгами {len(parts)}")
svod2 = an["etalon"]["СВОД_Производство -2 "]
print("  «СВОД_Производство -2 » против СВОД_Производство:")
for m in svod2:
    print(f"    {m}: -2 бн {fmt(svod2[m]['bn'])} нал {fmt(svod2[m]['nal'])} | Δбн {fmt(svod2[m]['bn'] - svod[m]['bn'])} Δнал {fmt(svod2[m]['nal'] - svod[m]['nal'])}")
# где «премия производств.персонал» на -2
for x in sheets["СВОД_Производство -2 "]["rows"]:
    if x["name"] and "преми" in norm(x["name"]):
        print(f"    r{x['r']} L{x['level']} «{x['name']}»: " + "; ".join(f"{m}: бн {fmt(v['bn'])} нал {fmt(v['nal'])}" for m, v in x["vals"].items() if v["bn"] or v["nal"]))

print("\n" + "=" * 100)
print("3. КАРТА «имя статьи книги → статья ДДС» (уровень 2 с деньгами по всем листам с блоками)")
st_by_norm = defaultdict(list)
for r in baza["statyi"]:
    st_by_norm[norm(r["Наименование"])].append(r)
# использование по объектным листам + ЦО (то, что реально поедет)
usage = defaultdict(lambda: {"sum": 0.0, "sheets": set(), "side": set()})
for nm in [CO] + OBJECTS:
    sh = sheets[nm]
    for x in sh["rows"]:
        if x["level"] != 2 or not x["name"] or x["r"] >= 156:
            continue
        tot = sum(v["bn"] + v["nal"] for v in x["vals"].values())
        if tot == 0:
            continue
        u = usage[norm(x["name"])]
        u["sum"] += tot; u["sheets"].add(nm); u["side"].add("П" if x["r"] < 19 else "Р")
print(f"  имён с деньгами на ЦО + 14 объектных: {len(usage)}")
for nm, u in sorted(usage.items(), key=lambda kv: -kv[1]["sum"]):
    cand = st_by_norm.get(nm, [])
    if len(cand) == 1:
        c = cand[0]
        print(f"  {''.join(sorted(u['side'])):<2} {nm:<48} {fmt(u['sum']):>16}  → {c['Код']} «{c['Наименование']}» [{c['Тип']}] / {str(c['Родитель'])[:28]}  ({len(u['sheets'])} л.)")
    else:
        print(f"  {''.join(sorted(u['side'])):<2} {nm:<48} {fmt(u['sum']):>16}  → {'НЕ НАЙДЕНО' if not cand else 'НЕОДНОЗНАЧНО ' + str([c['Код'] for c in cand])}  листы {sorted(u['sheets'])}")
check(all(len(st_by_norm.get(nm, [])) <= 1 for nm in usage), "есть неоднозначные имена статей")

print("\n" + "=" * 100)
print("4. ПРОЧЕЕ")
print("  статьи справочника с «юр»/«юри»:", [(r["Код"], r["Наименование"], r["Тип"]) for r in baza["statyi"] if "юр" in norm(r["Наименование"])])
print("  статьи справочника с «преми»:", [(r["Код"], r["Наименование"], r["Тип"]) for r in baza["statyi"] if "преми" in norm(r["Наименование"])])
for nm in ["МД IRC 2026", CO]:
    for x in sheets[nm]["rows"]:
        if x["name"] and norm(x["name"]) in ("суточные", "аренда квартиры", "получение финансовой помощи", "финансовая помощь"):
            print(f"  «{nm}» r{x['r']} L{x['level']} «{x['name']}»: " + ("; ".join(f"{m}: бн {fmt(v['bn'])} нал {fmt(v['nal'])}" for m, v in x["vals"].items() if v["bn"] or v["nal"]) or "нули"))
print("  комментарии IRC (сентябрь):", [(x["r"], x["name"], x["vals"]["2026-09-01"]["comment"]) for x in sheets["МД IRC 2026"]["rows"] if x["r"] < 156 and x["vals"].get("2026-09-01", {}).get("comment")])
print("  IRC верх/хвост колонок:", sheets["МД IRC 2026"]["tail_cols"])
print("  IRC top_rows 2/3:", sheets["МД IRC 2026"]["top_rows"]["2"], sheets["МД IRC 2026"]["top_rows"]["3"])
# скрытые объекты Производства: пусты?
for nm in ["Укрнафта 1 модуль", "МК Бориспільска", "МК Астарта.Тищенки", "МД ПРООН Черкаси  ДСНС"]:
    et = an["etalon"][nm]
    tot = sum(e["bn"] + e["nal"] for e in et.values())
    print(f"  «{nm}»: L2 итого по месяцам {fmt(tot)}")
# Підгірці и ООН 2025 — что за деньги
for nm in ["МК Підгірці", "МД ООН 2025", "МД Вооз ЕМS Мала 2026", "МК Чорноморськ", "МК Глобино", "МД ВООЗ  2025"]:
    print(f"  «{nm}»:", [(x["r"], x["name"], {m: (v["bn"], v["nal"]) for m, v in x["vals"].items() if v["bn"] or v["nal"]}) for x in sheets[nm]["rows"] if x["level"] == 2 and x["r"] < 156 and any(v["bn"] or v["nal"] for v in x["vals"].values())])

print("\n" + "=" * 100)
print(f"ВЫПОЛНЕНО ПРОВЕРОК {N} / ПРОВАЛОВ {M}")
sys.exit(1 if M else 0)
