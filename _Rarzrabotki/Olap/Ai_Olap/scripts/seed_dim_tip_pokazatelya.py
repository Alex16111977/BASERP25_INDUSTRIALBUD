# -*- coding: utf-8 -*-
"""Seed Dim_TipPokazatelya из Перечисление.ВидыСтатейУправленческогоБаланса.

Измерение А_ОтчетБаланс_Свод.ТипПоказателя заполняется в ПровестиБалансСвод
по формуле штатного Отчет.УправленческийБаланс (АктивПассив→Пассив, иначе
как есть) → в регистр пишется только «Актив»/«Пассив». enum_resolver
маппит ТипПоказателя ref → МЕТАИМЯ ("Актив"/"Пассив"; конвенция как
Source/ТипыНалогов). Dim сидируется из 1С COM Имя (== ключ
Fact_Balance.TipPokazatelya) + Представление 1С UI + _EnumOrder.
+строка "ПустаяСсылка" (порядок 99) на случай незаполненного ТипПоказателя.
Idempotent: применяет DDL (ALTER Fact_Balance + DROP+CREATE Dim) + INSERT.
"""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
import win32com.client
from ai_olap.core.connections import get_olap_sql

DDL = (pathlib.Path(__file__).parent / "ddl_dim_tip_pokazatelya.sql").read_text(encoding="utf-8")

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
md = erp.Метаданные.Перечисления.Найти("ВидыСтатейУправленческогоБаланса")
zn = md.ЗначенияПеречисления
mgr = erp.Перечисления.ВидыСтатейУправленческогоБаланса
rows = []
for i in range(zn.Количество()):
    имя = erp.String(zn.Получить(i).Имя)
    ref = getattr(mgr, имя)
    предст = erp.String(ref).strip() or имя
    rows.append((имя, предст, i))
rows.append(("ПустаяСсылка", "(Не определён)", 99))
print(f"1С Перечисление.ВидыСтатейУправленческогоБаланса: "
      f"{len(rows)-1} значений + ПустаяСсылка")

with get_olap_sql() as c:
    cur = c.cursor()
    for batch in [b.strip() for b in DDL.split("\nGO") if b.strip()]:
        cur.execute(batch)
    c.commit()
    cur.fast_executemany = True
    cur.executemany(
        "INSERT INTO dbo.Dim_TipPokazatelya "
        "(TipPokazatelya, TipPokazatelya_Name, EnumOrder) VALUES (?, ?, ?)",
        rows)
    c.commit()
    cur.execute("SELECT COUNT(*) FROM dbo.Dim_TipPokazatelya")
    n = cur.fetchval()
    cur.execute("SELECT TipPokazatelya, TipPokazatelya_Name, EnumOrder "
                "FROM dbo.Dim_TipPokazatelya ORDER BY EnumOrder")
    for r in cur.fetchall():
        print(f"  [{r.EnumOrder:>2}] {r.TipPokazatelya:<14} | {r.TipPokazatelya_Name}")
print(f"Dim_TipPokazatelya rows={n}")
assert n == len(rows), f"FAIL: вставлено {n} != {len(rows)}"
print("PASS: Dim_TipPokazatelya засидирован из 1С метаданных + "
      "Fact_Balance.TipPokazatelya колонка добавлена")
