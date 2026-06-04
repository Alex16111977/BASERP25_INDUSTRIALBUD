"""🎯 BULLET-PROOF ACCEPTANCE GATE.

Глобино-2 / Source=ERP_Income / Period_Month=2026-02-01 must equal 38 432 968.66 ₴.
This is the canonical etalon from olap_acceptance_etalons.md.
"""
from decimal import Decimal

import pytest

from ai_olap.core.connections import get_olap_sql

EXPECTED = Decimal("38432968.66")
TOLERANCE = Decimal("0.01")


def test_globyno_erp_income_february_2026():
    with get_olap_sql() as c:
        cur = c.cursor()
        cur.execute(
            """
            SELECT SUM(F.Sum_Fact) AS Total
            FROM Fact_PnL F
            JOIN Dim_Departments D ON F.Department_ID = D.Department_ID
            JOIN Dim_Income_Articles I ON F.Income_Article_ID = I.Income_Article_ID
            WHERE D.Department_Name = N'Глобино-2'
              AND F.Source = N'PL_ЕРП'
              AND I.Income_Article_Name LIKE N'Выручка от продаж%'
              AND F.Period_Month = '2026-02-01'
            """
        )
        total = cur.fetchval()

    assert total is not None, "no rows for Глобино-2 / ERP_Income / Feb 2026"
    delta = abs(total - EXPECTED)
    assert delta <= TOLERANCE, (
        f"acceptance gate failed: total={total}, expected={EXPECTED}, delta={delta}"
    )
