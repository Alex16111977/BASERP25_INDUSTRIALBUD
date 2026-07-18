# -*- coding: utf-8 -*-
"""
Общий модуль для скриптов knowledge РН.ДенежныеСредстваВПути.
"""
import sys, os, json, csv
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

EDRPOU_ORG_TOV = "40645273"
CONN_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ARTIFACTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
os.makedirs(ARTIFACTS_DIR, exist_ok=True)


def connect_erp():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN_STRING)


def money(x):
    if x is None: return ""
    try:
        s = f"{float(x):,.2f}"
        return s.replace(",", " ").replace(".", ",")
    except: return str(x)


def save_csv(name, rows, headers):
    path = os.path.join(ARTIFACTS_DIR, f"{name}.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    return path


def get_type_name(erp, ref):
    try:
        if ref is None or not erp.ЗначениеЗаполнено(ref): return ""
        return str(ref.Метаданные().Имя)
    except: return ""
