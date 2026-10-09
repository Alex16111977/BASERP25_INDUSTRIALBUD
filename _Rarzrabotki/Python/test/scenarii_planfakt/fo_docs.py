# -*- coding: utf-8 -*-
"""Финотчёт 09.10 с расшифровкой: по объекту — договоры (K, подписанная часть) и документы к получению."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dump import load

P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\09102026\Финансовый_отчет Производства 09-10-2026 с расшифровкой.xlsx"


def fmt(x):
    s = f"{abs(x):,.2f}".replace(",", " ").replace(".", ",")
    return ("−" if x < -0.004 else "") + s


def read():
    wf, wv = load(P, False), load(P, True)
    ws, wsv = wf.worksheets[0], wv.worksheets[0]

    def v(c):
        x = wsv[c].value
        return float(x) if isinstance(x, (int, float)) else 0.0

    def ev(f, hzero):
        def rep(m):
            col, r = m.group(1), int(m.group(2))
            return "0" if (col == "H" and hzero) else repr(v("%s%d" % (col, r)))
        e = re.sub(r"\b([A-Z])(\d+)\b", rep, f[1:]).replace("ROUND", "round").replace("MAX", "max").replace("MIN", "min")
        return eval(e)

    objs, obj, dog = {}, None, None
    for r in range(6, ws.max_row + 1):
        a = ws["A%d" % r].value
        c = ws["C%d" % r].value
        if a and isinstance(c, str) and re.fullmatch(r"=C%d" % (r + 1), c):
            obj = {"name": a.strip(), "K": v("K%d" % r), "G": v("G%d" % r), "H": v("H%d" % r), "I": v("I%d" % r),
                   "D": v("D%d" % r), "F": v("F%d" % r), "dogs": []}
            objs[a.split(" (")[0].strip()] = obj
            dog = None
            continue
        if obj is None or not a:
            continue
        if str(a).strip() == "Итого":
            obj, dog = None, None
            continue
        k = ws["K%d" % r].value
        if isinstance(k, str) and "MAX(0" in k.upper():
            dog = {"name": str(a).strip(), "K": v("K%d" % r), "pess": round(ev(k, True), 2), "G": v("G%d" % r),
                   "H": v("H%d" % r), "I": v("I%d" % r), "J": v("J%d" % r), "docs": []}
            obj["dogs"].append(dog)
            continue
        if dog is not None and str(a).startswith("Реализ"):
            b = str(ws["B%d" % r].value or "")
            n = str(ws["N%d" % r].value or "")
            h, e = v("H%d" % r), v("E%d" % r)
            if h or "не оплачено" in n:
                m = re.search(r"(\S+) от (\d\d\.\d\d\.\d{4})", a)
                doc = (m.group(1) + " от " + m.group(2)[:5]) if m else a
                dog["docs"].append({"doc": doc, "status": b, "H": h, "E": e,
                                    "unpaid": ("не оплачено" in n), "N": n.split("\n")[0]})
    return objs


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for k, o in read().items():
        print(k, fmt(o["K"]))
        for d in o["dogs"]:
            print("   ", d["name"][:50], "K", fmt(d["K"]), "песс", fmt(d["pess"]))
            for x in d["docs"]:
                print("       ", x["doc"], "|", x["status"], "| H", fmt(x["H"]), "| E", fmt(x["E"]), "|", x["N"])
