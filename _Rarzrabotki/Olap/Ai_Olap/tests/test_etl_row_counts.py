"""Etalon row counts for February 2026 (from olap_acceptance_etalons.md)."""
from ai_olap.core.connections import get_olap_sql


def test_fact_pnl_february_2026_row_count():
    with get_olap_sql() as c:
        cur = c.cursor()
        cur.execute("SELECT COUNT(*) FROM Fact_PnL WHERE Period_Month = '2026-02-01'")
        n = cur.fetchval()
    assert n == 3937, f"Fact_PnL Feb 2026: expected 3937, got {n}"


def test_fact_cashflow_february_2026_row_count():
    with get_olap_sql() as c:
        cur = c.cursor()
        cur.execute("SELECT COUNT(*) FROM Fact_Cashflow WHERE Period_Month = '2026-02-01'")
        n = cur.fetchval()
    assert n == 4652, f"Fact_Cashflow Feb 2026: expected 4652, got {n}"


def test_fact_pnl_distinct_sources():
    """Knowledge base says 7 distinct Source values appear in February
    (ERP_БезPL_Расх may have 0 rows -> 7 not 8)."""
    with get_olap_sql() as c:
        cur = c.cursor()
        cur.execute("SELECT COUNT(DISTINCT Source) FROM Fact_PnL WHERE Period_Month = '2026-02-01'")
        n = cur.fetchval()
    assert n == 7, f"distinct Source: expected 7, got {n}"


def test_dim_organizations_only_own_company():
    """The whole pipeline is scoped to KodEDRPOU=40645273 (TOВ ІНДАСТРІАЛБУД)."""
    with get_olap_sql() as c:
        cur = c.cursor()
        cur.execute("SELECT COUNT(*), MIN(Organization_Name) FROM Dim_Organizations")
        row = cur.fetchone()
    assert row[0] == 1, f"expected exactly 1 own organization, got {row[0]}"
    assert "ІНДАСТРІАЛБУД" in (row[1] or ""), f"unexpected org name: {row[1]!r}"
