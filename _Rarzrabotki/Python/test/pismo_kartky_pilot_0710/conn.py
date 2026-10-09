# -*- coding: utf-8 -*-
"""Подключения COM (пароли из .v8-project.json), только чтение."""
import sys, json, datetime
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
import win32com.client, pywintypes

PRJ = r"C:\Configuration_downloads\BASERP25\.v8-project.json"


def _cs(db_id):
    d = json.load(open(PRJ, encoding="utf-8-sig"))
    for x in d["databases"]:
        if x["id"] == db_id:
            return 'Srvr="localhost";Ref="%s";Usr="%s";Pwd="%s"' % (x["ref"], x["user"], x["password"])
    raise SystemExit("нет базы " + db_id)


def erp():
    return win32com.client.Dispatch("V83.COMConnector").Connect(_cs("erp"))


def buh():
    return win32com.client.Dispatch("V83.COMConnector").Connect(_cs("buhbud"))


def dt(y, m, d, h=0, mi=0, s=0):
    # COM сдвигает наивную дату на -3 ч: компенсируем
    return pywintypes.Time(datetime.datetime(y, m, d, h, mi, s) + datetime.timedelta(hours=3))


def uid(c, ref):
    try:
        return c.String(ref.УникальныйИдентификатор())
    except Exception:
        return ""
