# -*- coding: utf-8 -*-
"""Перевірити чи РКО помічений на удаление."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

with open(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\rko_candidates.json", encoding="utf-8") as f:
    cands = json.load(f)

for k in cands:
    uuid_obj = conn.NewObject("УникальныйИдентификатор", k["uuid"])
    ref = conn.Документы.РасходныйКассовыйОрдер.ПолучитьСсылку(uuid_obj)
    obj = ref.ПолучитьОбъект()
    if obj is None:
        print(f"  №{k['number']}: ОБЪЕКТ НЕ ЗНАЙДЕНО")
        continue
    print(f"  №{k['number']}: Проведен={obj.Проведен}, ПометкаУдаления={obj.ПометкаУдаления}")
