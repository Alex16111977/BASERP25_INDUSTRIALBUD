# -*- coding: utf-8 -*-
"""Seed Dim_TipyDogovorov из Перечисление.ТипыДоговоров (1С метаданные).
_Enum1626 в SQL backend имеет только _IDRRef/_EnumOrder (имён НЕТ).
Dim_Contracts.TipDogovora после enum_resolver хранит МЕТАИМЯ. Поэтому Dim
сидируется из 1С COM: ключ=Имя, TipDogovora_Name=Представление (синоним),
EnumOrder=_EnumOrder. +строка "ПустаяСсылка" (порядок 99) для договоров без
типа. Idempotent: применяет DDL (DROP+CREATE) + INSERT."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
import win32com.client
from ai_olap.core.connections import get_olap_sql

DDL = (pathlib.Path(__file__).parent / "ddl_dim_tipy_dogovorov.sql").read_text(encoding="utf-8")

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
md = erp.Метаданные.Перечисления.Найти("ТипыДоговоров")
zn = md.ЗначенияПеречисления
mgr = erp.Перечисления.ТипыДоговоров
rows = []
for i in range(zn.Количество()):
    имя = erp.String(zn.Получить(i).Имя)
    ref = getattr(mgr, имя)
    предст = erp.String(ref).strip() or имя
    rows.append((имя, предст, i))
rows.append(("ПустаяСсылка", "(Не указан)", 99))
print(f"1С Перечисление.ТипыДоговоров: {len(rows)-1} значений + ПустаяСсылка")

with get_olap_sql() as c:
    cur = c.cursor()
    for batch in [b.strip() for b in DDL.split("\nGO") if b.strip()]:
        cur.execute(batch)
    c.commit()
    cur.fast_executemany = True
    cur.executemany(
        "INSERT INTO dbo.Dim_TipyDogovorov (TipDogovora, TipDogovora_Name, EnumOrder) "
        "VALUES (?, ?, ?)", rows)
    c.commit()
    cur.execute("SELECT COUNT(*) FROM dbo.Dim_TipyDogovorov")
    n = cur.fetchval()
    cur.execute("SELECT TipDogovora, TipDogovora_Name, EnumOrder "
                "FROM dbo.Dim_TipyDogovorov ORDER BY EnumOrder")
    for r in cur.fetchall():
        print(f"  [{r.EnumOrder:>2}] {r.TipDogovora:<24} | {r.TipDogovora_Name}")
print(f"Dim_TipyDogovorov rows={n}")
assert n == len(rows), f"FAIL: вставлено {n} != {len(rows)}"
print("PASS: Dim_TipyDogovorov засидирован из 1С метаданных")
