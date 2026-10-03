# -*- coding: utf-8 -*-
"""Z3 база: загрузка raw15.json (снимок BuhBud), классы складов по Склады.Подразделение (правило M5 §2.1 п. 10, §6.6),
факт домов 15 м по карточкам (без ПДВ, как Хозрасчетный), склад проекта, последние закупки. Только чтение файлов."""
import os, sys, re, json, collections, datetime
from decimal import Decimal, ROUND_HALF_UP
if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = json.load(open(os.environ.get("Z3_RAW", os.path.join(HERE, "raw15.json")), encoding="utf-8"))
Z = Decimal(0)
VAT = Decimal("1.2")
H15 = [f"15-{i}" for i in range(1, 7)]
HNAME = {f"15-{i}": f"15 м №{i}" for i in range(1, 7)}


def Dd(x):
    return Decimal(str(x)) if x not in (None, "") else Z


def r2(x):
    return float(Dd(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def r3(x):
    return float(Dd(x).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))


def r6(x):
    return float(Dd(x).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP))


K = lambda s: " ".join(str(s or "").replace("\xa0", " ").split()).casefold()


def pdate(s):
    """«dd.mm.yyyy h:mm:ss» (строка COM, уже местное время 1С) -> datetime."""
    s = (s or "").strip()
    if not s:
        return None
    a, *b = s.split(" ")
    d, m, y = a.split(".")
    hh, mi, ss = (b[0].split(":") + ["0", "0", "0"])[:3] if b else ("0", "0", "0")
    return datetime.datetime(int(y), int(m), int(d), int(hh), int(mi), int(ss))


def house_of(podr):
    m = re.search(r"IR[CС] 2026 (15|30) м №\s*(\d+)", podr or "")
    return f"{m.group(1)}-{m.group(2)}" if m else None


# --- склады: класс по Склады.Подразделение (а не по имени)
WH = {w["Имя"]: w for w in RAW["warehouses"]}
PROJECT_PODR = {"МД IRC 2026"}


def wcls(name):
    if not name:
        return "—"
    w = WH.get(name)
    if w:
        podr = (w["Подр"] or "").strip()
        h = house_of(podr)
        if h:
            return "H:" + h
        if podr in PROJECT_PODR:
            return "P"
    if name in ("МД IRС 2026",):          # склад без подразделения, имя проекта (как PROJECT_W прототипа)
        return "P"
    if "МШП" in name:
        return "MSP"
    return "X"


def wstage(name):
    w = WH.get(name)
    v = (w or {}).get("Вид", "")
    return "Монтаж" if v == "Монтаж" else ("Виробничий" if v == "Виробничий" else v or "—")


is_mat = lambda a: (a or "")[:2] in ("20", "22", "28")

# --- карточки
CARD = {}
for c in RAW["cards"]:
    CARD[c["Код"]] = {"name": c["Имя"].strip(), "on": (c["ОН"] or "").strip(), "on_code": c["ОНКод"], "unit": c["Ед"], "ukey": c["ukey"],
                      "del": c["Пом"]}

MV = RAW["moves"]
for r in MV:
    r["dt"] = pdate(r["Период"])
    r["doc"] = r["Рег"]

# --- КН: шапки (вид КН «по нормам/додаткова» из РК; служебные)
KN = {}
for k in RAW["kn"]:
    KN[k["Номер"]] = k


def docnum(rep):
    m = re.search(r"([А-ЯІЇЄA-Z]{2}\d{2}-\d{6})", rep or "")
    return m.group(1) if m else rep


# --- свёртки по карточке
st = collections.defaultdict(lambda: collections.defaultdict(lambda: [Z, Z]))       # код -> класс -> [кол, сумма] (остаток по движениям)
FACT = []        # строки факта домов 15 м: списание со склада дома на нематериальный счёт
ASM = []         # сборки полуфабрикатов на складах домов 15 м (Кт мат -> Дт мат другой карточки)
for r in MV:
    dt_, kt_ = r["СчДт"] or "", r["СчКт"] or ""
    s = Dd(r["Сумма"])
    if is_mat(dt_):
        a = st[r["КодДт"]][wcls(r["СклДт"])]; a[0] += Dd(r["КолДт"] or 0); a[1] += s
    if is_mat(kt_):
        ck = wcls(r["СклКт"])
        a = st[r["КодКт"]][ck]; a[0] -= Dd(r["КолКт"] or 0); a[1] -= s
        if ck.startswith("H:15-"):
            h = ck[2:]
            if not is_mat(dt_):
                FACT.append({"house": h, "dt": r["dt"], "doc": docnum(r["doc"]), "type": r["ТипРег"], "acc_kt": kt_, "acc_dt": dt_,
                             "code": r["КодКт"], "q": Dd(r["КолКт"] or 0), "s": s, "wh": r["СклКт"], "stage": wstage(r["СклКт"])})
            elif r["КодДт"] != r["КодКт"] and wcls(r["СклДт"]) == ck and r["ТипРег"] == "Комплектація номенклатури":
                ASM.append({"house": h, "doc": docnum(r["doc"]), "out": r["КодДт"], "in": r["КодКт"], "q": Dd(r["КолКт"] or 0), "s": s})


def bal(code, cls):
    a = st[code][cls]
    return a[0], a[1]


# --- последние закупки (ПТУ BuhBud, проведённые), без ПДВ
LAST = {}
PURCH = collections.defaultdict(list)
for r in RAW["receipts"]:
    if not r["Проведен"] or not Dd(r["Кол"]):
        continue
    net = Dd(r["Сумма"]) - (Dd(r["НДС"]) if r["ВклНДС"] else Z)
    dts = pdate(r["Дата"])
    rec = {"date": dts, "price_n": net / Dd(r["Кол"]), "q": Dd(r["Кол"]), "kontr": r["Контрагент"], "doc": docnum(r["Док"])}
    PURCH[r["Код"]].append(rec)
    cur = LAST.get(r["Код"])
    if cur is None or dts > cur["date"]:
        LAST[r["Код"]] = rec
