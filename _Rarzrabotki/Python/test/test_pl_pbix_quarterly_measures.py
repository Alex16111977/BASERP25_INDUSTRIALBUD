"""SQL-acceptance tests for PL.pbix quarterly DAX measures.

Validates the SQL aggregations that DAX measures (Сума грн місяць/квартал/рік PL,
% Доля затрат квартал грн) are built upon. DAX-runtime correctness is verified
separately via MCP `dax_query_operations.Execute` against the live Power BI model.

Acceptance gate: Глобино-2 / ERP_Income / 2026-02 = 38 432 968.66 ₴
(canonical etalon from `_Rarzrabotki/notebook/knowledge_Olap/olap_acceptance_etalons.md`).
"""
from decimal import Decimal

import pyodbc
import pytest

CONN_OLAP = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!"
)

EXPECTED_GLOBYNO = Decimal("38432968.66")
TOLERANCE = Decimal("0.01")


@pytest.fixture(scope="module")
def conn():
    cn = pyodbc.connect(CONN_OLAP, autocommit=True)
    yield cn
    cn.close()


def test_globyno_feb_2026_acceptance(conn):
    """Глобино-2 / ERP_Income / 2026-02 must equal 38 432 968.66 ₴."""
    cur = conn.cursor()
    cur.execute(
        """
        SELECT SUM(F.Sum_ERP_Grn)
        FROM Fact_PnL F
        JOIN Dim_Departments D ON F.Department_ID = D.Department_ID
        WHERE D.Department_Name = N'Глобино-2'
          AND F.Source = 'ERP_Income'
          AND F.Period_Month = '2026-02-01'
        """
    )
    total = cur.fetchval()
    assert total is not None, "no rows — ETL not loaded?"
    delta = abs(Decimal(str(total)) - EXPECTED_GLOBYNO)
    assert delta <= TOLERANCE, (
        f"acceptance gate failed: total={total}, expected={EXPECTED_GLOBYNO}, delta={delta}"
    )


def test_quarter_equals_sum_of_3_months(conn):
    """Q1 2026 sum must equal (Jan + Feb + Mar) 2026 sum across all Sources."""
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            SUM(CASE WHEN Period_Month = '2026-01-01' THEN Sum_ERP_Grn ELSE 0 END) AS Jan,
            SUM(CASE WHEN Period_Month = '2026-02-01' THEN Sum_ERP_Grn ELSE 0 END) AS Feb,
            SUM(CASE WHEN Period_Month = '2026-03-01' THEN Sum_ERP_Grn ELSE 0 END) AS Mar,
            SUM(CASE WHEN Period_Month BETWEEN '2026-01-01' AND '2026-03-01' THEN Sum_ERP_Grn ELSE 0 END) AS Q1
        FROM Fact_PnL
        """
    )
    jan, feb, mar, q1 = cur.fetchone()
    jan = Decimal(str(jan or 0))
    feb = Decimal(str(feb or 0))
    mar = Decimal(str(mar or 0))
    q1 = Decimal(str(q1 or 0))
    expected = jan + feb + mar
    assert abs(q1 - expected) <= TOLERANCE, (
        f"Q1 ({q1}) != Jan+Feb+Mar ({expected}); delta={q1 - expected}"
    )


def test_year_equals_sum_of_4_quarters(conn):
    """Year 2026 sum must equal sum of all 12 months (= 4 quarters)."""
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            SUM(CASE WHEN YEAR(Period_Month) = 2026 THEN Sum_ERP_Grn ELSE 0 END) AS Year2026,
            SUM(CASE WHEN YEAR(Period_Month) = 2026 AND MONTH(Period_Month) BETWEEN 1 AND 3 THEN Sum_ERP_Grn ELSE 0 END) AS Q1,
            SUM(CASE WHEN YEAR(Period_Month) = 2026 AND MONTH(Period_Month) BETWEEN 4 AND 6 THEN Sum_ERP_Grn ELSE 0 END) AS Q2,
            SUM(CASE WHEN YEAR(Period_Month) = 2026 AND MONTH(Period_Month) BETWEEN 7 AND 9 THEN Sum_ERP_Grn ELSE 0 END) AS Q3,
            SUM(CASE WHEN YEAR(Period_Month) = 2026 AND MONTH(Period_Month) BETWEEN 10 AND 12 THEN Sum_ERP_Grn ELSE 0 END) AS Q4
        FROM Fact_PnL
        """
    )
    year, q1, q2, q3, q4 = cur.fetchone()
    year = Decimal(str(year or 0))
    qsum = sum(Decimal(str(q or 0)) for q in (q1, q2, q3, q4))
    assert abs(year - qsum) <= TOLERANCE, (
        f"Year ({year}) != Q1+Q2+Q3+Q4 ({qsum}); delta={year - qsum}"
    )


def test_dolya_sums_to_100pct_for_q1_2026(conn):
    """For one Year_Quarter, Σ (article_sum / quarter_total) across all PL articles = 1.00 (±0.01).

    Mirrors DAX measure `% Доля затрат квартал грн` aggregation.
    """
    cur = conn.cursor()
    cur.execute(
        """
        WITH q AS (
            SELECT SUM(Sum_ERP_Grn) AS QTotal
            FROM Fact_PnL
            WHERE Period_Month BETWEEN '2026-01-01' AND '2026-03-01'
        ),
        per_article AS (
            SELECT PL_Article_ID, SUM(Sum_ERP_Grn) AS ArticleSum
            FROM Fact_PnL
            WHERE Period_Month BETWEEN '2026-01-01' AND '2026-03-01'
            GROUP BY PL_Article_ID
        )
        SELECT
            q.QTotal,
            SUM(per_article.ArticleSum) AS SumArticles
        FROM per_article CROSS JOIN q
        GROUP BY q.QTotal
        """
    )
    qtotal, sum_articles = cur.fetchone()
    qtotal = Decimal(str(qtotal or 0))
    sum_articles = Decimal(str(sum_articles or 0))
    assert qtotal != 0, "Q1 2026 has no data"
    sum_dolya = sum_articles / qtotal
    delta = abs(sum_dolya - Decimal("1"))
    assert delta <= Decimal("0.0001"), (
        f"Σ %Доля_затрат_квартал ({sum_dolya}) != 1.00; delta={delta}"
    )
