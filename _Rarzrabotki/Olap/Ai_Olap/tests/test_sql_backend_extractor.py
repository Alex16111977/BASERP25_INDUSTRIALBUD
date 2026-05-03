"""SqlBackendExtractor smoke — read at least one row from a known catalog."""
from ai_olap.extractors.sql_backend import SqlBackendExtractor


def test_organizations_yields_at_least_one_row():
    ex = SqlBackendExtractor(
        {
            "type": "sql_backend",
            "object": "Справочник.Организации",
            "fields": ["Ссылка", "Наименование"],
            "where": "_Fld32799 = '40645273'",
            "limit": 5,
        }
    )
    rows = ex.extract()
    assert rows, "no rows returned"
    assert "Наименование" in rows[0]
    assert "ІНДАСТРІАЛБУД" in (rows[0]["Наименование"] or "")


def test_raw_sql_path():
    ex = SqlBackendExtractor(
        {
            "type": "raw_sql",
            "raw_sql": "SELECT TOP 1 1 AS One",
        }
    )
    rows = ex.extract()
    assert rows == [{"One": 1}]
