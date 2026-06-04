"""Regenerate dim_catalogs.json with raw_sql + recursive CTE (Stage v3.7).

Replaces 4 hierarchical + 3 flat pipeline steps with raw_sql that pre-computes
Hierarchy_Path, Hierarchy_Depth, Level1..Level5 (and Direction_ID cascade for
Departments) in SQL Server. Adds unknown-member ('(Пусто)' row with
ID = 0x...0001) to every Dim that previously lacked it.

Replaced / extended steps (8 total):
  - dim_dds_articles      (_Reference529, depth=5)
        + А_РазделCFS (_Fld55969RRef → CFS_Section via enum_resolver)
        + А_ИсключатьИзОтчетаCashflow (_Fld56081 → Is_Excluded_From_Cashflow)
  - dim_departments       (_Reference540, depth=3, has_folder=False)
        + Direction_ID (_Fld54614RRef) COALESCE cascade self→ancestor→unknown
  - dim_items             (_Reference306, depth=5) — 41k rows, MAXRECURSION 0
  - dim_expense_articles  (_Chrc1772 ChartOfCharacteristics, depth=4)
  - dim_partners          (_Reference360, depth=5, has_folder=False) — NEW
  - dim_counterparties    (_Reference263, flat + Partner_ID/EDRPOU/Tax_Code + unknown)
  - dim_directions        (_Reference292, hierarchy + unknown)
  - dim_income_articles   (_Chrc1771, flat + unknown — no hierarchy)

Run once: `python scripts/update_pipeline_hierarchies.py`
Then: `python main.py --run-once dim_catalogs`

SQL field numbers verified via sqlcmd probe + baserp_storage.json (2026-05-11).
"""
from __future__ import annotations
import json
from pathlib import Path

PIPELINE = Path(__file__).resolve().parents[1] / "pipelines" / "dim_catalogs.json"

UNKNOWN_ID_HEX = "0x00000000000000000000000000000001"
ZERO_UUID_HEX = "0x00000000000000000000000000000000"


def hierarchy_cte_sql(
    table: str,
    code_col: str = "_Code",
    has_folder: bool = True,
    extra_select_root: str = "",
    extra_select_recurse: str = "",
    extra_select_unknown: str = "",
    extra_columns_in_alias: str = "",
) -> str:
    """Build recursive CTE that produces:
       Ссылка, Код, Наименование, Родитель, ЭтоГруппа, ПометкаУдаления,
       Hierarchy_Path, Hierarchy_Depth, Level1..Level5 [+ extra_columns]
       + UNION ALL unknown-member row (0x...0001 "(Пусто)").
    """
    folder_root = "d._Folder AS ЭтоГруппа" if has_folder else "CAST(0x00 AS varbinary(1)) AS ЭтоГруппа"
    folder_recurse = "d._Folder" if has_folder else "CAST(0x00 AS varbinary(1))"
    return (
        f"WITH RecCTE AS ("
        f"SELECT d._IDRRef AS Ссылка, d.{code_col} AS Код, d._Description AS Наименование, "
        f"d._ParentIDRRef AS Родитель, {folder_root}, d._Marked AS ПометкаУдаления, "
        f"CAST(CONVERT(varchar(32), d._IDRRef, 2) AS nvarchar(500)) AS Hierarchy_Path, "
        f"1 AS Hierarchy_Depth, "
        f"CAST(d._Description AS nvarchar(150)) AS Level1, "
        f"CAST(NULL AS nvarchar(150)) AS Level2, "
        f"CAST(NULL AS nvarchar(150)) AS Level3, "
        f"CAST(NULL AS nvarchar(150)) AS Level4, "
        f"CAST(NULL AS nvarchar(150)) AS Level5"
        f"{extra_select_root} "
        f"FROM {table} d "
        f"WHERE d._ParentIDRRef = {ZERO_UUID_HEX} "
        f"UNION ALL "
        f"SELECT d._IDRRef, d.{code_col}, d._Description, d._ParentIDRRef, {folder_recurse}, d._Marked, "
        f"CAST(r.Hierarchy_Path + N'|' + CONVERT(varchar(32), d._IDRRef, 2) AS nvarchar(500)), "
        f"r.Hierarchy_Depth + 1, "
        f"r.Level1, "
        f"CASE WHEN r.Hierarchy_Depth + 1 = 2 THEN CAST(d._Description AS nvarchar(150)) ELSE r.Level2 END, "
        f"CASE WHEN r.Hierarchy_Depth + 1 = 3 THEN CAST(d._Description AS nvarchar(150)) ELSE r.Level3 END, "
        f"CASE WHEN r.Hierarchy_Depth + 1 = 4 THEN CAST(d._Description AS nvarchar(150)) ELSE r.Level4 END, "
        f"CASE WHEN r.Hierarchy_Depth + 1 = 5 THEN CAST(d._Description AS nvarchar(150)) ELSE r.Level5 END"
        f"{extra_select_recurse} "
        f"FROM {table} d INNER JOIN RecCTE r ON d._ParentIDRRef = r.Ссылка"
        f") "
        f"SELECT Ссылка, Код, Наименование, Родитель, ЭтоГруппа, ПометкаУдаления, "
        f"Hierarchy_Path, Hierarchy_Depth, Level1, Level2, Level3, Level4, Level5"
        f"{extra_columns_in_alias} "
        f"FROM RecCTE "
        f"UNION ALL "
        f"SELECT CAST({UNKNOWN_ID_HEX} AS varbinary(16)) AS Ссылка, "
        f"'00-UNK' AS Код, N'(Пусто)' AS Наименование, "
        f"CAST({ZERO_UUID_HEX} AS varbinary(16)) AS Родитель, "
        f"CAST(0x00 AS varbinary(1)) AS ЭтоГруппа, "
        f"CAST(0x00 AS varbinary(1)) AS ПометкаУдаления, "
        f"N'(Пусто)' AS Hierarchy_Path, 1 AS Hierarchy_Depth, "
        f"N'(Пусто)' AS Level1, "
        f"CAST(NULL AS nvarchar(150)) AS Level2, "
        f"CAST(NULL AS nvarchar(150)) AS Level3, "
        f"CAST(NULL AS nvarchar(150)) AS Level4, "
        f"CAST(NULL AS nvarchar(150)) AS Level5"
        f"{extra_select_unknown} "
        f"OPTION (MAXRECURSION 0)"
    )


def flat_with_unknown_sql(
    table: str,
    code_col: str = "_Code",
    has_parent: bool = True,
    has_folder: bool = True,
    extra_select_real: str = "",
    extra_select_unknown: str = "",
) -> str:
    """Flat (no recursive CTE) catalog + UNION ALL unknown-member row."""
    parent_real = "d._ParentIDRRef" if has_parent else f"CAST({ZERO_UUID_HEX} AS varbinary(16))"
    folder_real = "d._Folder" if has_folder else "CAST(0x00 AS varbinary(1))"
    code_real = f"d.{code_col}" if code_col else "CAST(N'' AS nvarchar(50))"
    return (
        f"SELECT d._IDRRef AS Ссылка, {code_real} AS Код, d._Description AS Наименование, "
        f"{parent_real} AS Родитель, {folder_real} AS ЭтоГруппа, d._Marked AS ПометкаУдаления"
        f"{extra_select_real} "
        f"FROM {table} d "
        f"UNION ALL "
        f"SELECT CAST({UNKNOWN_ID_HEX} AS varbinary(16)), '00-UNK', N'(Пусто)', "
        f"CAST({ZERO_UUID_HEX} AS varbinary(16)), CAST(0x00 AS varbinary(1)), CAST(0x00 AS varbinary(1))"
        f"{extra_select_unknown}"
    )


# ===== Step SQL expressions =====

DDS_EXTRA_ROOT = (
    ", d._Fld55969RRef AS А_РазделCFS, "
    "COALESCE(d._Fld56081, 0x00) AS А_ИсключатьИзОтчетаCashflow"
)
DDS_EXTRA_RECURSE = ", d._Fld55969RRef, COALESCE(d._Fld56081, 0x00)"
DDS_EXTRA_OUTER = ", А_РазделCFS, А_ИсключатьИзОтчетаCashflow"
DDS_EXTRA_UNKNOWN = (
    f", CAST({ZERO_UUID_HEX} AS varbinary(16)) AS А_РазделCFS, "
    f"CAST(0x00 AS varbinary(1)) AS А_ИсключатьИзОтчетаCashflow"
)
DDS_SQL = hierarchy_cte_sql(
    "_Reference529",
    extra_select_root=DDS_EXTRA_ROOT,
    extra_select_recurse=DDS_EXTRA_RECURSE,
    extra_select_unknown=DDS_EXTRA_UNKNOWN,
    extra_columns_in_alias=DDS_EXTRA_OUTER,
)

EXPENSE_SQL = hierarchy_cte_sql("_Chrc1772")
ITEMS_SQL = hierarchy_cte_sql("_Reference306")

DEPT_EXTRA_ROOT = (
    f", COALESCE(NULLIF(d._Fld54614RRef, {ZERO_UUID_HEX}), CAST({UNKNOWN_ID_HEX} AS varbinary(16))) AS Direction_ID"
)
DEPT_EXTRA_RECURSE = (
    f", COALESCE(NULLIF(d._Fld54614RRef, {ZERO_UUID_HEX}), r.Direction_ID, CAST({UNKNOWN_ID_HEX} AS varbinary(16)))"
)
DEPT_EXTRA_OUTER = ", Direction_ID"
DEPT_EXTRA_UNKNOWN = f", CAST({UNKNOWN_ID_HEX} AS varbinary(16)) AS Direction_ID"
DEPT_SQL = hierarchy_cte_sql(
    "_Reference540",
    has_folder=False,
    extra_select_root=DEPT_EXTRA_ROOT,
    extra_select_recurse=DEPT_EXTRA_RECURSE,
    extra_columns_in_alias=DEPT_EXTRA_OUTER,
    extra_select_unknown=DEPT_EXTRA_UNKNOWN,
)

PARTNERS_SQL = hierarchy_cte_sql("_Reference360", has_folder=False)

DIRECTIONS_SQL = hierarchy_cte_sql("_Reference292")

CP_EXTRA_REAL = (
    ", d._Fld31169RRef AS Партнер, d._Fld31175 AS КодПоЕДРПОУ, d._Fld31173 AS НалоговыйНомер"
)
CP_EXTRA_UNKNOWN = (
    f", CAST({ZERO_UUID_HEX} AS varbinary(16)) AS Партнер, "
    f"CAST(N'' AS nvarchar(20)) AS КодПоЕДРПОУ, "
    f"CAST(N'' AS nvarchar(20)) AS НалоговыйНомер"
)
COUNTERPARTIES_SQL = flat_with_unknown_sql(
    "_Reference263",
    code_col="",
    has_parent=False,
    has_folder=False,
    extra_select_real=CP_EXTRA_REAL,
    extra_select_unknown=CP_EXTRA_UNKNOWN,
)

INCOME_SQL = flat_with_unknown_sql("_Chrc1771")


# ===== Pipeline patch =====

HIER_COL_MAP_BASE = {
    "Родитель": "Parent_ID",
    "ЭтоГруппа": "Is_Group",
    "ПометкаУдаления": "Marked_For_Deletion",
    "Hierarchy_Path": "Hierarchy_Path",
    "Hierarchy_Depth": "Hierarchy_Depth",
    "Level1": "Level1",
    "Level2": "Level2",
    "Level3": "Level3",
    "Level4": "Level4",
    "Level5": "Level5",
}

DEFAULTS_BOOL = {"Is_Group": False, "Marked_For_Deletion": False}


def make_hier_step(
    step_id: str,
    sql: str,
    target_table: str,
    pk_dst: str,
    code_dst: str,
    name_dst: str,
    extra_map: dict | None = None,
    transformer_steps: list[str] | None = None,
    enum_resolver_opts: dict | None = None,
) -> dict:
    cm = dict(HIER_COL_MAP_BASE)
    cm["Ссылка"] = pk_dst
    cm["Код"] = code_dst
    cm["Наименование"] = name_dst
    if extra_map:
        cm.update(extra_map)
    steps = transformer_steps or ["varbinary_to_uuid", "column_mapper"]
    options: dict = {"column_mapper": {"column_map": cm, "defaults": DEFAULTS_BOOL}}
    if enum_resolver_opts:
        options["enum_resolver"] = enum_resolver_opts
    return {
        "step_id": step_id,
        "extractor": {"type": "raw_sql", "sql": sql},
        "transformer": {"steps": steps, "options": options},
        "loader": {"target_table": target_table, "mode": "full_reload"},
    }


def make_flat_step(
    step_id: str,
    sql: str,
    target_table: str,
    pk_dst: str,
    name_dst: str,
    code_dst: str | None = None,
    extra_map: dict | None = None,
) -> dict:
    cm = {
        "Ссылка": pk_dst,
        "Наименование": name_dst,
        "Родитель": "Parent_ID",
        "ЭтоГруппа": "Is_Group",
        "ПометкаУдаления": "Marked_For_Deletion",
    }
    if code_dst:
        cm["Код"] = code_dst
    if extra_map:
        cm.update(extra_map)
    return {
        "step_id": step_id,
        "extractor": {"type": "raw_sql", "sql": sql},
        "transformer": {
            "steps": ["varbinary_to_uuid", "column_mapper"],
            "options": {
                "column_mapper": {"column_map": cm, "defaults": DEFAULTS_BOOL},
            },
        },
        "loader": {"target_table": target_table, "mode": "full_reload"},
    }


def replace_step(steps: list, step_id: str, new_step: dict) -> bool:
    for i, s in enumerate(steps):
        if s["step_id"] == step_id:
            steps[i] = new_step
            return True
    return False


def insert_step_after(steps: list, after_step_id: str, new_step: dict) -> bool:
    for i, s in enumerate(steps):
        if s["step_id"] == after_step_id:
            if any(x["step_id"] == new_step["step_id"] for x in steps):
                return False
            steps.insert(i + 1, new_step)
            return True
    return False


def main():
    cfg = json.loads(PIPELINE.read_text(encoding="utf-8"))
    steps = cfg["steps"]

    replace_step(steps, "dim_dds_articles", make_hier_step(
        "dim_dds_articles", DDS_SQL, "Dim_DDS_Articles",
        "DDS_Article_ID", "DDS_Article_Code", "DDS_Article_Name",
        extra_map={
            "А_РазделCFS": "CFS_Section",
            "А_ИсключатьИзОтчетаCashflow": "Is_Excluded_From_Cashflow",
        },
        transformer_steps=["varbinary_to_uuid", "enum_resolver", "column_mapper"],
        enum_resolver_opts={"column_to_enum": {"А_РазделCFS": "Перечисление.А_РазделыCFS"}},
    ))

    replace_step(steps, "dim_expense_articles", make_hier_step(
        "dim_expense_articles", EXPENSE_SQL, "Dim_Expense_Articles",
        "Expense_Article_ID", "Expense_Article_Code", "Expense_Article_Name",
    ))

    replace_step(steps, "dim_items", make_hier_step(
        "dim_items", ITEMS_SQL, "Dim_Items",
        "Item_ID", "Item_Code", "Item_Name",
    ))

    replace_step(steps, "dim_departments", make_hier_step(
        "dim_departments", DEPT_SQL, "Dim_Departments",
        "Department_ID", "Department_Code", "Department_Name",
        extra_map={"Direction_ID": "Direction_ID"},
    ))

    replace_step(steps, "dim_income_articles", make_flat_step(
        "dim_income_articles", INCOME_SQL, "Dim_Income_Articles",
        "Income_Article_ID", "Income_Article_Name",
        code_dst="Income_Article_Code",
    ))

    replace_step(steps, "dim_directions", make_hier_step(
        "dim_directions", DIRECTIONS_SQL, "Dim_Directions",
        "Direction_ID", "Direction_Code", "Direction_Name",
    ))

    replace_step(steps, "dim_counterparties", {
        "step_id": "dim_counterparties",
        "extractor": {"type": "raw_sql", "sql": COUNTERPARTIES_SQL},
        "transformer": {
            "steps": ["varbinary_to_uuid", "column_mapper"],
            "options": {
                "column_mapper": {
                    "column_map": {
                        "Ссылка": "Counterparty_ID",
                        "Наименование": "Counterparty_Name",
                        "ПометкаУдаления": "Marked_For_Deletion",
                        "Партнер": "Partner_ID",
                        "КодПоЕДРПОУ": "Code_EDRPOU",
                        "НалоговыйНомер": "Tax_Code",
                    },
                    "defaults": DEFAULTS_BOOL,
                },
            },
        },
        "loader": {"target_table": "Dim_Counterparties", "mode": "full_reload"},
    })

    insert_step_after(steps, "dim_counterparties", make_hier_step(
        "dim_partners", PARTNERS_SQL, "Dim_Partners",
        "Partner_ID", "Partner_Code", "Partner_Name",
    ))

    PIPELINE.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"OK: rewrote {PIPELINE}")
    print(f"Total steps: {len(steps)}")
    print(
        "Modified/added: dim_dds_articles, dim_expense_articles, dim_items, "
        "dim_departments, dim_income_articles, dim_directions, dim_counterparties, "
        "+ dim_partners (NEW)"
    )


if __name__ == "__main__":
    main()
