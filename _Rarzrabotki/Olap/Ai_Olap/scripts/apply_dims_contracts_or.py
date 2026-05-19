# -*- coding: utf-8 -*-
"""Применяет DDL recreate Dim_Contracts + Dim_ObjektyRaschetov.
Данные заливает ETL (dim_catalogs full_reload). Idempotent (DROP+CREATE)."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

HERE = pathlib.Path(__file__).parent
# Одно соединение + единая транзакция: парная schema-миграция атомарна
# (падение второго DDL откатывает DROP первого).
with get_olap_sql() as c:
    cur = c.cursor()
    for fname in ("ddl_dim_contracts.sql", "ddl_dim_objekty_raschetov.sql"):
        ddl = (HERE / fname).read_text(encoding="utf-8")
        for batch in [b.strip() for b in ddl.split("\nGO") if b.strip()]:
            cur.execute(batch)
        print(f"applied {fname}")
    c.commit()
print("PASS: Dim_Contracts + Dim_ObjektyRaschetov recreated")
