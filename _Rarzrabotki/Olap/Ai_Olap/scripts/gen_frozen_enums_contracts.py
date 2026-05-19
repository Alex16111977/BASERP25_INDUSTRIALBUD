# -*- coding: utf-8 -*-
"""Печатает готовые к вставке FROZEN_ENUMS-литералы (Имя в порядке _EnumOrder)
для 3 enum, нужных Dim_Contracts/Dim_ObjektyRaschetov. READ-ONLY (1С COM)."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
for meta in ("ТипыДоговоров", "ТипыРасчетовСПартнерами", "ТипыОбъектовРасчетов"):
    zn = erp.Метаданные.Перечисления.Найти(meta).ЗначенияПеречисления
    print(f'    "Перечисление.{meta}": [')
    for i in range(zn.Количество()):
        имя = erp.String(zn.Получить(i).Имя)
        print(f'        "{имя}",  # {i}')
    print("    ],")

