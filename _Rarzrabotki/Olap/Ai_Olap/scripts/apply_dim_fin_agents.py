# -*- coding: utf-8 -*-
"""Применяет DDL Dim_FinAgents. Данные заливает ETL (dim_catalogs full_reload)."""
import sys, io, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from ai_olap.core.connections import get_olap_sql

ddl = (pathlib.Path(__file__).parent / "ddl_dim_fin_agents.sql").read_text(encoding="utf-8")
with get_olap_sql() as c:
    cur = c.cursor()
    for batch in [b.strip() for b in ddl.split("\nGO") if b.strip()]:
        cur.execute(batch)
    c.commit()
print("PASS: Dim_FinAgents created")
