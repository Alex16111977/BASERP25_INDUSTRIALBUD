import sys, json, openpyxl
from collections import defaultdict, Counter
sys.stdout.reconfigure(encoding="utf-8")
B = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\07102026"
rows = json.load(open("erp_mat_0710.json", encoding="utf-8"))
def key(r):
    if r["on"].strip(): return "ОН|" + r["on"].strip()
    if r["nom_uuid"] and r["nom_uuid"] != "00000000-0000-0000-0000-000000000000": return "НОМ|" + r["nom_name"].strip()
    return "ТЕКСТ|" + r["an"].strip().upper()
def xlsx_groups(p):
    wb = openpyxl.load_workbook(B + "\\" + p + "\План-факт виконання СС на 07-10-2026.xlsx", data_only=True)
    ws = wb.active
    inmat = False; g = {}; tot = None; cur = None; sub = defaultdict(list)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value; ind = ws.cell(r, 1).alignment.indent
        if ind == 2.0:
            inmat = (str(v).strip() == "2. Матеріали")
            if inmat: tot = [ws.cell(r, c).value or 0 for c in (10, 11, 12)]
            continue
        if not inmat: continue
        if ind == 4.0:
            cur = (v or "").strip() if v else ""
            g[cur] = [ws.cell(r, c).value or 0 for c in (10, 11, 12)]
        elif ind == 6.0:
            sub[cur].append((v, [ws.cell(r, c).value or 0 for c in (10, 11, 12)]))
    return tot, g, sub
for P, folder in (("МД IRC 2026", "IRC 2026"), ("МД ООН Админ", "МД ООН Админ")):
    R = [r for r in rows if r["podr"] == P]
    print("=====", P, "строк", len(R), Counter(r["type"] for r in R))
    print(" ПодрОбъекта:", Counter(r["podr_obj"] for r in R))
    print(" СС-документы:", Counter(r["doc"] for r in R if r["type"] == "СС"))
    pg = sum(r["pg"] for r in R); fg = sum(r["fg"] for r in R); pnf = sum(r["pnf"] for r in R)
    print(" План %.2f  Факт %.2f  ETC %.2f  Прогноз(ПнФ) %.2f  Відх %.2f" % (pg, fg, pnf, fg + pnf, fg + pnf - pg))
    tot, xg, sub = xlsx_groups(folder)
    print(" xlsx итог:", tot)
    G = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for r in R:
        k = key(r); G[k][0] += r["pg"]; G[k][1] += r["fg"]; G[k][2] += r["pnf"]
    # сверка групп с xlsx по имени
    def xname(k): return k.split("|", 1)[1]
    ok = bad = 0; miss = []
    xkeys = dict((k, v) for k, v in xg.items())
    used = set()
    for k, v in G.items():
        n = xname(k)
        cand = [x for x in xkeys if x.upper() == n.upper()]
        if not cand:
            miss.append((k, v)); continue
        x = xkeys[cand[0]]; used.add(cand[0])
        d = (round(v[0] - x[0], 2), round(v[1] + v[2] - x[1], 2))
        if abs(d[0]) < 0.011 and abs(d[1]) < 0.011: ok += 1
        else: bad += 1; print("  РАСХ", k, [round(v[0],2), round(v[1]+v[2],2)], x)
    print(" групп ERP", len(G), "xlsx", len(xg), "совпало", ok, "расх", bad, "не найдено в xlsx", len(miss))
    for k, v in miss[:30]: print("   нет в xlsx:", k, [round(x, 2) for x in v])
    nx = [x for x in xkeys if x not in used]
    for x in nx[:30]: print("   нет в ERP:", repr(x), xkeys[x])
