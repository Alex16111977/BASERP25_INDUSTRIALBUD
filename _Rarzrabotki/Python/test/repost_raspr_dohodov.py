# -*- coding: utf-8 -*-
"""
Перепровести документ РаспределениеДоходовПоНаправлениямДеятельности
UUID: a6f0a6dd-fd21-11f0-a2e2-c54425f51b91
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN_ERP)

DOC_UUID = "a6f0a6dd-fd21-11f0-a2e2-c54425f51b91"

doc_ref = conn.Documents.РаспределениеДоходовПоНаправлениямДеятельности.ПолучитьСсылку(
    conn.NewObject("УникальныйИдентификатор", DOC_UUID))

doc = doc_ref.ПолучитьОбъект()
print(f"Документ: {doc.Номер} від {doc.Дата}")
print("Перепроводжу...")

mode = conn.РежимЗаписиДокумента.Проведение
doc.Write(mode)
print("Документ перепроведено успішно!")
