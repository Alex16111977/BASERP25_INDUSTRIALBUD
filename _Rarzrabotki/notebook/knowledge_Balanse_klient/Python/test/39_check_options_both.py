# -*- coding: utf-8 -*-
"""СКРИПТ 39 — Сравнение констант и функциональных опций в BASERP25 vs BaseERPRazr.
Без сравнения остального — только ключевые опции взаиморасчётов."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")

def check(label, conn_str):
    print(f"\n[{label}]")
    erp = v8.Connect(conn_str)
    for c in ("НоваяАрхитектураВзаиморасчетов", "ФормироватьУправленческийБаланс", "ЗаполненыДвиженияАктивовПассивов"):
        try:
            v = erp.Константы[c].Получить()
            print(f"  Константа.{c:<45} {v}")
        except Exception as e:
            print(f"  Константа.{c:<45} ERR {str(e)[:60]}")
    for fo in ("НоваяАрхитектураВзаиморасчетов", "ФормироватьФинансовыйРезультат"):
        try:
            v = erp.ПолучитьФункциональнуюОпцию(fo)
            print(f"  ФО.{fo:<54} {v}")
        except Exception as e:
            print(f"  ФО.{fo:<54} ERR {str(e)[:60]}")

print("=" * 110)
check("BASERP25 (наша)",   'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
check("BaseERPRazr (типовая, эталон)", 'Srvr="localhost";Ref="BaseERPRazr";Usr="Администратор";Pwd="24043"')
