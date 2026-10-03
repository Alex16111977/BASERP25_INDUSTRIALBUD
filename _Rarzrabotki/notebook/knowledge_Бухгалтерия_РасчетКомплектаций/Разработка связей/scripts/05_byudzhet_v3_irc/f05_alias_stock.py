# -*- coding: utf-8 -*-
"""Замечание 1: где лежат алиасы Lacrysil / «Круг відрізний 230» (склад проекта P, ІБ_МШП, дома) и откуда выдавались на 30 м; G_h по клей / обрізні диски."""
import sys, json, re, collections
sys.stdout.reconfigure(encoding="utf-8")
OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\ekonomiya_ss_irc\out"
DATA = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\Мой расчет\data"
MV = json.load(open(OUT + r"\buh_moves.json", encoding="utf-8"))
F = json.load(open(DATA + r"\irc_fakt.json", encoding="utf-8"))
K = lambda s: " ".join(str(s or "").replace("\xa0", " ").split()).casefold()
is_mat = lambda a: (a or "")[:2] in ("20", "22", "28")
pat = lambda n: "lacrysil" in K(n) or ("круг відрізний" in K(n) and "230" in K(n)) or ("клей" in K(n) and "монтаж" in K(n))
st = collections.defaultdict(lambda: [0.0, 0.0])
for r in MV:
    for side, sg in (("Дт", 1), ("Кт", -1)):
        acc = r["СчДт"] if side == "Дт" else r["СчКт"]
        n = r["Ном" + side]
        if is_mat(acc) and pat(n):
            st[(n, r["Скл" + side], acc)][0] += sg * float(r["Кол" + side] or 0)
            st[(n, r["Скл" + side], acc)][1] += sg * float(r["Сумма"] or 0)
for k, v in sorted(st.items()):
    if abs(v[0]) > 1e-6 or abs(v[1]) > 0.005:
        print("остаток", k, [round(x, 2) for x in v])
for r in MV:
    if is_mat(r["СчКт"]) and not is_mat(r["СчДт"]) and pat(r["НомКт"]) and "30 м" in (r["СклКт"] or ""):
        print("выдача", r["Период"][:10], r["НомКт"], r["СклКт"], r["СчКт"], r["КолКт"], r["Сумма"], r["ОНКт"])
for it in F["deficit"]["items"]:
    if it["key"] in ("клей", "обрізні диски", "дріт", "електроди", "насосна станція"):
        print(it["key"], it["nq"], it["price_ss"], it["def_q"], it["def"], it["sum"], "stockP", it.get("stockP_q"), "c3", it.get("consumed3_q"))
