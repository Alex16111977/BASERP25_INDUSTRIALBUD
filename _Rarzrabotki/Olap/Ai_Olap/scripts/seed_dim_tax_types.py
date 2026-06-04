# -*- coding: utf-8 -*-
"""Seed Dim_TaxTypes из Перечисление.ТипыНалогов (1С метаданные).

Перечисление — frozen enum: _Enum1651 в SQL backend имеет только
_IDRRef/_EnumOrder (имён/синонимов НЕТ). Fact_Balance.TaxType после
enum_resolver хранит МЕТАИМЯ ("НДС"/"ВоенныйСбор"/...; "ПустаяСсылка" у
строк НЕ-«Налоги»). Поэтому Dim сидируется из 1С COM: ключ=Имя (== ключ
Fact_Balance.TaxType), TaxType_Name=Представление (синоним 1С UI),
EnumOrder=_EnumOrder. +строка "ПустаяСсылка" (порядок 99) для строк без
налога. Idempotent: применяет DDL (DROP+CREATE) + INSERT.
"""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
import win32com.client
from ai_olap.core.connections import get_olap_sql

DDL = (pathlib.Path(__file__).parent / "ddl_dim_tax_types.sql").read_text(encoding="utf-8")

# 1С: Имя|Представление|Порядок всех значений Перечисление.ТипыНалогов
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
md = erp.Метаданные.Перечисления.Найти("ТипыНалогов")
zn = md.ЗначенияПеречисления
mgr = erp.Перечисления.ТипыНалогов
rows = []
for i in range(zn.Количество()):
    имя = erp.String(zn.Получить(i).Имя)
    ref = getattr(mgr, имя)
    предст = erp.String(ref).strip() or имя
    rows.append((имя, предст, i))
# Пустая ссылка enum: строки НЕ-«Налоги» (Fact_Balance.TaxType="ПустаяСсылка")
rows.append(("ПустаяСсылка", "(Не налог)", 99))
print(f"1С Перечисление.ТипыНалогов: {len(rows)-1} значений + ПустаяСсылка")

with get_olap_sql() as c:
    cur = c.cursor()
    # DDL: DROP+CREATE (выполняем по батчам через GO)
    for batch in [b.strip() for b in DDL.split("\nGO") if b.strip()]:
        cur.execute(batch)
    c.commit()
    cur.fast_executemany = True
    cur.executemany(
        "INSERT INTO dbo.Dim_TaxTypes (TaxType, TaxType_Name, EnumOrder) "
        "VALUES (?, ?, ?)", rows)
    c.commit()
    cur.execute("SELECT COUNT(*) FROM dbo.Dim_TaxTypes")
    n = cur.fetchval()
    cur.execute("SELECT TaxType, TaxType_Name, EnumOrder FROM dbo.Dim_TaxTypes "
                "ORDER BY EnumOrder")
    for r in cur.fetchall():
        print(f"  [{r.EnumOrder:>2}] {r.TaxType:<40} | {r.TaxType_Name}")
print(f"Dim_TaxTypes rows={n}")
assert n == len(rows), f"FAIL: вставлено {n} != {len(rows)}"
print("PASS: Dim_TaxTypes засидирован из 1С метаданных")
