# -*- coding: utf-8 -*-
"""Финотчёт с расшифровкой: по объекту — договоры (K, условия оплаты) и документы к получению со сроками оплаты.

Пессимистичный сценарий = подписанные акты (оригинал получен) с неоплаченным остатком, срок оплаты которых наступил
или наступит до конца месяца отчёта: такой платёж можем требовать с клиента. Срок берётся из финотчёта (колонка M
«срок DD.MM.YYYY» / «до DD.MM.YYYY»), условие — из строки договора (колонка N «Договор: …»)."""
import sys, os, re, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dump import load

P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\09102026\Финансовый_отчет Производства 09-10-2026 с расшифровкой.xlsx"
DATA_OTCHETA = dt.date(2026, 10, 9)
KONEC_MESYACA = dt.date(2026, 10, 31)
STATUSY = ("оригинал получен", "не подписан", "не проведён")


def fmt(x):
    s = f"{abs(x):,.2f}".replace(",", " ").replace(".", ",")
    return ("−" if x < -0.004 else "") + s


def d2s(d):
    return d.strftime("%d.%m.%Y") if d else "?"


def parse_date(s):
    m = re.search(r"(\d\d)\.(\d\d)\.(\d{4})", s or "")
    return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def parse_sum(s):
    m = re.search(r"не оплачено ([\d\s\u00a0]+,\d\d)", s or "")
    return float(m.group(1).replace(" ", "").replace("\u00a0", "").replace(",", ".")) if m else 0.0


def term(text):
    """«оплата 5 банк. дн. после подписания акта» → (5, 'банк'); None — срока нет."""
    m = re.search(r"оплата (\d+) (банк|раб|кал)\. дн\. после ([^\n;]+)", text or "")
    return (int(m.group(1)), m.group(2), m.group(3).strip()) if m else None


def podpisat_do(t):
    """Последний день подписания акта, при котором оплата по договору придёт до конца месяца."""
    if not t:
        return None
    n, kind, _ = t
    d = KONEC_MESYACA
    if kind == "кал":
        return d - dt.timedelta(days=n)
    while d.weekday() >= 5:  # последний рабочий день месяца — крайний день оплаты
        d -= dt.timedelta(days=1)
    k = 0
    while k < n:  # подпись + n банковских дней ≤ последний рабочий день месяца
        d -= dt.timedelta(days=1)
        if d.weekday() < 5:
            k += 1
    return d


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
            usl = str(ws["N%d" % r].value or "").split("\n")[0]
            dog = {"name": str(a).strip(), "K": v("K%d" % r), "pess_formula": round(ev(k, True), 2), "G": v("G%d" % r),
                   "H": v("H%d" % r), "usl": usl, "term": term(usl), "docs": []}
            obj["dogs"].append(dog)
            continue
        b = str(ws["B%d" % r].value or "")
        if dog is None or str(a).startswith("Поступление") or not b.startswith(STATUSY):
            continue
        n = str(ws["N%d" % r].value or "").split("\n")[0]
        mtext = str(ws["M%d" % r].value or "")
        h = v("H%d" % r)
        m = re.search(r"(\S+) от (\d\d\.\d\d)\.\d{4}", a)
        doc = (m.group(1) + " от " + m.group(2)) if m else a
        signed = b.startswith("оригинал получен")
        unpaid = parse_sum(n) if signed else 0.0
        if not (h or unpaid):
            continue
        srok = None
        mm = re.search(r"(?:срок|до) (\d\d\.\d\d\.\d{4})", mtext)
        if mm:
            srok = parse_date(mm.group(1))
        dog["docs"].append({"doc": doc, "status": b, "signed": signed, "sign_date": parse_date(b) if signed else None,
                            "H": h, "unpaid": unpaid, "M": mtext, "N": n, "srok": srok})
    for o in objs.values():
        o["pess"] = 0.0
        for d in o["dogs"]:
            s = 0.0
            for x in d["docs"]:
                if x["signed"] and x["unpaid"]:
                    x["v_pess"] = (x["srok"] is None) or (x["srok"] <= KONEC_MESYACA)
                    if x["v_pess"]:
                        s += x["unpaid"]
            d["pess"] = round(s, 2)
            o["pess"] += d["pess"]
        o["pess"] = round(o["pess"], 2)
    return objs


def signed_lines(d):
    out = []
    for x in d["docs"]:
        if not (x["signed"] and x["unpaid"]):
            continue
        if x["srok"] is None:
            vyvod = "срок по договору не определён — требовать можем, ставим"
        elif x["srok"] < DATA_OTCHETA:
            vyvod = "срок прошёл, просрочено %d дн. — можем требовать, ставим" % (DATA_OTCHETA - x["srok"]).days
        elif x["srok"] <= KONEC_MESYACA:
            vyvod = "срок наступает в этом месяце — ставим"
        else:
            vyvod = "срок после %s — в этом месяце требовать не можем, не ставим" % d2s(KONEC_MESYACA)
        out.append("  • %s, подписан %s: не оплачено %s; срок оплаты %s → %s"
                   % (x["doc"], d2s(x["sign_date"]), fmt(x["unpaid"]), d2s(x["srok"]), vyvod))
    return out


def unsigned_lines(d):
    out = []
    do = podpisat_do(d["term"])
    for x in d["docs"]:
        if x["signed"] or not x["H"]:
            continue
        if do:
            hint = "оплата придёт в октябре, если подпишут до %s" % d2s(do)
        else:
            hint = "срок оплаты от подписания по договору не задан"
        out.append("  • %s — %s, %s; %s" % (x["doc"], fmt(x["H"]), x["status"], hint))
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for k, o in read().items():
        print(k, "K", fmt(o["K"]), "песс по срокам", fmt(o["pess"]))
        for d in o["dogs"]:
            print("   ", d["name"][:50], "| K", fmt(d["K"]), "| песс", fmt(d["pess"]), "| формула H=0", fmt(d["pess_formula"]), "|", d["usl"])
            for s in signed_lines(d) + unsigned_lines(d):
                print("     ", s)
