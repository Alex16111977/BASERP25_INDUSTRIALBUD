# -*- coding: utf-8 -*-
"""Общие помощники исследования «экономия по материалам против СС» (МД IRC 2026).

Только чтение: подключения COM к ERP / BuhBud, запуск запросов, чтение xlsx-выгрузок 1С.
"""
import sys, os, io, zipfile, datetime, json
from decimal import Decimal, ROUND_HALF_UP

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = r"C:\Configuration_downloads\BASERP25"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT, exist_ok=True)

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_BUH = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'

_v8 = None


def v8():
    global _v8
    if _v8 is None:
        import win32com.client
        _v8 = win32com.client.Dispatch("V83.COMConnector")
    return _v8


def erp():
    return v8().Connect(CONN_ERP)


def buh():
    return v8().Connect(CONN_BUH)


def q(conn, text, **params):
    """Выполнить запрос, вернуть ТаблицуЗначений (COM). Ошибку — понятным текстом."""
    z = conn.NewObject("Запрос")
    z.Text = text
    for k, v in params.items():
        z.SetParameter(k, v)
    try:
        return z.Execute().Выгрузить()
    except Exception as e:
        msg = e.excepinfo[2] if getattr(e, "excepinfo", None) else str(e)
        raise RuntimeError("ЗАПРОС УПАЛ: " + str(msg))


def rows(conn, vt, cols):
    """ТЗ -> список dict; значения-ссылки -> строка представления через conn.String."""
    res = []
    for i in range(vt.Количество()):
        r = vt.Получить(i)
        d = {}
        for c in cols:
            val = getattr(r, c)
            if val is None or isinstance(val, (int, float, str, bool, Decimal)):
                d[c] = val
            elif isinstance(val, datetime.datetime):
                d[c] = val
            else:
                try:
                    d[c] = conn.String(val)
                except Exception:
                    d[c] = str(val)
        res.append(d)
    return res


def D(x):
    if x is None or x == "":
        return Decimal("0")
    return Decimal(str(x))


def r2(x):
    return D(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def r3(x):
    return D(x).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)


def fmt(x, n=2):
    x = D(x).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP)
    s = f"{x:,.{n}f}".replace(",", " ").replace(".", ",")
    return s


def xlsx_bytes_fixed(path):
    """Выгрузки 1С кладут xl/SharedStrings.xml с заглавной S — переписываем zip для openpyxl."""
    src = zipfile.ZipFile(path)
    buf = io.BytesIO()
    dst = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
    for it in src.infolist():
        data = src.read(it.filename)
        name = it.filename
        if name == "xl/SharedStrings.xml":
            name = "xl/sharedStrings.xml"
        if name in ("[Content_Types].xml", "xl/_rels/workbook.xml.rels"):
            data = data.replace(b"SharedStrings.xml", b"sharedStrings.xml")
        dst.writestr(name, data)
    dst.close()
    buf.seek(0)
    return buf


def load_xlsx(path, data_only=True):
    import openpyxl
    try:
        return openpyxl.load_workbook(path, data_only=data_only)
    except Exception:
        return openpyxl.load_workbook(xlsx_bytes_fixed(path), data_only=data_only)


def dump_json(obj, name):
    p = os.path.join(OUT, name)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, default=str)
    return p


def norm_name(s):
    return " ".join(str(s or "").split())


def classify_row(r, g):
    """Категория строки расхода (g = 'd15' | 'd30') с выделением «единица/состав»: цена < 0,6 нормы при объёме > 2 норм
    (или цена > 1,67 при объёме < 0,5) — в СС и в списании разная единица или комплектация (плинтус, септик, грунт).
    Единое правило для f60/f61/f62/f70."""
    cat = r[f"{g}_cat"]
    if cat == "сопоставимо":
        g_ = lambda f: D(r[f"{g}_{f}"])
        cq, Nq, cg, Ns = g_("cq"), g_("Nq"), g_("cg"), g_("Ns")
        if cq and Nq and Ns:
            pr = (cg / cq) / (Ns / Nq); qr = cq / Nq
            if (pr < D("0.6") and qr > 2) or (pr > D("1.67") and qr < D("0.5")):
                return "единица/состав"
    return cat
