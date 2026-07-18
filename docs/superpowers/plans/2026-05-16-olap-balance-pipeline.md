# OLAP-конвеєр управлінського балансу (Fact_Balance → PL.pbix) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Перенести `РегистрСведений.А_ОтчетБаланс_Свод` з BaseERP у SQL `OlapBASERP` (нова `Fact_Balance` + `Dim_PAP_Articles`) і додати у `PL.pbix` модель балансу (зв'язки + DAX [Opening]/[Inflow]/[Outflow]/[Closing] + розріз Актив/Пасив + сторінка Balance), дзеркалячи готовий патерн DDS→Fact_Cashflow.

**Architecture:** Свёртка вже зроблена в 1С (Етапи 1-4, live). Python ETL `Ai_Olap` (SQL-first pyodbc до MSSQL backend BaseERP) лише КОПІЮЄ готовий регістр у SQL — нічого не рахує. Шлях: `А_ОтчетБаланс_Свод` → `refresh_mapping` (resolve `_InfoRg/_Fld`) → `pipelines/fact_balance.json` → SQL `Fact_Balance` → `PL.pbix` (Import mode). Розріз Актив/Пасив — атрибут `Dim_PAP_Articles.AktivPassiv` (канон OD-9), рахується у DAX, не в ETL.

**Tech Stack:** SQL Server 2022 (`OlapBASERP`, `sa`/`Brw739182465!`, ODBC Driver 17), Python 3 + pyodbc (`Ai_Olap`), декларативні JSON-pipeline, Power BI Desktop + MCP `powerbi-modeling-mcp`, pytest.

**Reference (knowledge_Olap):** `olap_etl_pipeline.md`, `olap_sql_schema.md`, `olap_powerbi_pl_pbix.md`, `pipelines/fact_cashflow.json` (точний шаблон). Канон балансу: `docs/superpowers/specs/2026-05-15-balans-svod-canonical-design.md` (v1.4, §10/Roadmap, OD-9).

**Шляхи:** worktree `C:\Configuration_downloads\BASERP25\.claude\worktrees\peaceful-knuth-51295b`; main `C:\Configuration_downloads\BASERP25`. Ai_Olap: `_Rarzrabotki/Olap/Ai_Olap/`. DDL: `_Rarzrabotki/Python/Olap/ddl/`. PL.pbix: `_Rarzrabotki/Olap/PowerBi/PL.pbix`.

---

## Передумови (хард-ворота, до Task 1)

1. **Етапи 1-4 А_ФинРез_Баланс — live у BaseERP** (січень 2026/ТОВ ІНДАСТРІАЛБУД, Σ==ПАП до копійки). Перевірка: `python _Rarzrabotki/Python/test/test_balans_s4_verify.py` → `PASS Етап 4`. Якщо ні — STOP, спершу завершити баланс-цикл.
2. **OlapBASERP доступна:** `python -c "import pyodbc; pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!').close(); print('OK')"` → `OK`.
3. **Свёртка тільки в 1С** — ETL НЕ рахує суми, лише копіює регістр. Суми Fact_Balance по `Статья` ОБОВ'ЯЗКОВО == ПАП/Етап4 до копійки (acceptance gate Task 4/5).
4. **Типові 1С об'єкти не міняти** — лише читання SQL backend BaseERP. Зміни — тільки у `OlapBASERP` (нові таблиці) та `Ai_Olap`/`PL.pbix`.
5. **Rule-#-1-стиль:** спершу `refresh_mapping` resolve реальні `_InfoRg/_Fld` (Task 2), потім `raw_sql` під resolved імена (Task 3). НЕ вгадувати `_Fld` номери.
6. **Не чіпати БД/PL.pbix поза задачами плану.** Кожна задача — окремий commit.

---

## File Structure

| Файл | Відповідальність | Дія |
|---|---|---|
| `_Rarzrabotki/Python/Olap/ddl/07_balance.sql` | DDL `Fact_Balance` + `Dim_PAP_Articles` (ідемпотентно) | Create |
| `_Rarzrabotki/Python/Olap/ddl/apply_07_balance.py` | Застосувати 07_balance.sql через pyodbc + verify | Create |
| `_Rarzrabotki/Olap/Ai_Olap/scripts/introspect_balance_fields.py` | Дамп resolved `_InfoRg/_Fld` для А_ОтчетБаланс_Свод | Create |
| `_Rarzrabotki/Olap/Ai_Olap/pipelines/fact_balance.json` | ETL pipeline Fact_Balance (шаблон fact_cashflow.json) | Create |
| `_Rarzrabotki/Olap/Ai_Olap/pipelines/dim_pap_articles.json` | ETL pipeline Dim_PAP_Articles (СтатьиАктивовПассивов+АктивПассив) | Create |
| `_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json` | +запис А_ОтчетБаланс_Свод (через refresh_mapping.py) | Modify (авто) |
| `_Rarzrabotki/Olap/Ai_Olap/tests/test_etl_acceptance_balance.py` | Acceptance: ΣFact_Balance по Статья == ПАП; Актив=Пасив | Create |
| `PL.pbix` (live) | +Fact_Balance +Dim_СтатьиАктивовПассивов +relationships +DAX +сторінка | Modify (MCP) |
| `_Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md` / `olap_etl_pipeline.md` / `olap_powerbi_pl_pbix.md` / `KNOWLEDGE_MAP_OLAP.md` | Оновити під Balance Stage | Modify |

---

## Task 1: SQL DDL — Fact_Balance + Dim_PAP_Articles

**Files:**
- Create: `_Rarzrabotki/Python/Olap/ddl/07_balance.sql`
- Create: `_Rarzrabotki/Python/Olap/ddl/apply_07_balance.py`

- [ ] **Step 1: Написати DDL** `_Rarzrabotki/Python/Olap/ddl/07_balance.sql`

```sql
-- 07_balance.sql — Fact_Balance + Dim_PAP_Articles (ідемпотентно). ASCII-only (без Cyrillic).
IF OBJECT_ID('dbo.Fact_Balance','U') IS NULL
CREATE TABLE Fact_Balance (
    Balance_ID          bigint IDENTITY(1,1) PRIMARY KEY,
    Period_Month        date NOT NULL,                 -- '2026-01-01'
    Period              datetime2 NULL,                -- дата документа А_ФинРез_Баланс
    Source              varchar(40) NOT NULL,          -- enum А_ИсточникБаланса (8 значень)
    Recorder_Balance_ID char(32) NULL,                 -- UUID Документ.А_ФинРез_Баланс
    -- 17 вимірів (char(32) UUID-ключі, як Fact_Cashflow)
    Organization_ID     char(32) NOT NULL,
    Department_ID       char(32) NULL,
    PAP_Article_ID      char(32) NULL,                 -- Статья (СтатьиАктивовПассивов) -> Dim_PAP_Articles
    Item_ID             char(32) NULL,                 -- Номенклатура
    Counterparty_ID     char(32) NULL,
    Partner_ID          char(32) NULL,
    Warehouse_ID        char(32) NULL,                 -- Склад
    OperObject_ID       char(32) NULL,                 -- ОбъектыЭксплуатации
    Contract_ID         char(32) NULL,                 -- Договор
    Individual_ID       char(32) NULL,                 -- ФизическоеЛицо
    Cash_ID             char(32) NULL,                 -- ДенежныеСредства (composite UUID)
    SettlementObj_ID    char(32) NULL,                 -- ОбъектРасчетов
    Intangible_ID       char(32) NULL,                 -- НематериальныйАктив
    Analytics1          nvarchar(150) NULL,
    Analytics2          nvarchar(150) NULL,
    Analytics3          nvarchar(150) NULL,
    -- 4 ресурси (decimal(15,2), 1:1 з регістром)
    Sum_Open            decimal(15,2) NOT NULL DEFAULT 0,   -- СуммаНачальныйОстаток
    Sum_Inflow          decimal(15,2) NOT NULL DEFAULT 0,   -- СуммаПриход
    Sum_Outflow         decimal(15,2) NOT NULL DEFAULT 0,   -- СуммаРасход
    Sum_Close           decimal(15,2) NOT NULL DEFAULT 0,   -- СуммаКонечныйОстаток
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME(),
    INDEX IX_Bal_Period_Source (Period_Month, Source),
    INDEX IX_Bal_Article       (PAP_Article_ID, Period_Month),
    INDEX IX_Bal_Individual    (Individual_ID, Period_Month),
    INDEX IX_Bal_SettlementObj (SettlementObj_ID, Period_Month)
);

IF OBJECT_ID('dbo.Dim_PAP_Articles','U') IS NULL
CREATE TABLE Dim_PAP_Articles (
    PAP_Article_ID      char(32) PRIMARY KEY,
    PAP_Article_Code    varchar(50) NULL,
    PAP_Article_Name    nvarchar(150) NOT NULL,
    Parent_ID           char(32) NULL,
    Is_Group            bit NOT NULL DEFAULT 0,
    AktivPassiv         varchar(15) NULL,              -- 'Aktiv'|'Passiv'|'AktivPassiv' (з реквізиту АктивПассив)
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At           datetime2 NOT NULL DEFAULT SYSDATETIME()
);
```

- [ ] **Step 2: Написати applier** `_Rarzrabotki/Python/Olap/ddl/apply_07_balance.py`

```python
# -*- coding: utf-8 -*-
"""Застосувати 07_balance.sql у OlapBASERP + verify (ідемпотентно)."""
import sys, io, pyodbc, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
CONN = ("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
        "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;")
sql = pathlib.Path(__file__).with_name("07_balance.sql").read_text(encoding="utf-8")
cx = pyodbc.connect(CONN, autocommit=True); cu = cx.cursor()
for batch in [b.strip() for b in sql.split("\nGO") if b.strip()]:
    cu.execute(batch)
for t in ("Fact_Balance", "Dim_PAP_Articles"):
    n = cu.execute(f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME=?", t).fetchval()
    print(f"{t}: {'OK' if n==1 else 'MISSING'}")
    assert n == 1, f"FAIL: {t} не створено"
cx.close()
print("PASS Task 1 DDL")
```

- [ ] **Step 3: Запустити**

Run: `python "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\Olap\ddl\apply_07_balance.py"`
Expected: `Fact_Balance: OK`, `Dim_PAP_Articles: OK`, `PASS Task 1 DDL`.

- [ ] **Step 4: Copy back у main + Commit**

```bash
cp "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b/_Rarzrabotki/Python/Olap/ddl/07_balance.sql" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Python/Olap/ddl/07_balance.sql"
cp "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b/_Rarzrabotki/Python/Olap/ddl/apply_07_balance.py" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Python/Olap/ddl/apply_07_balance.py"
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b"
git add _Rarzrabotki/Python/Olap/ddl/07_balance.sql _Rarzrabotki/Python/Olap/ddl/apply_07_balance.py
git commit -m "feat(olap): DDL Fact_Balance + Dim_PAP_Articles в OlapBASERP"
```

---

## Task 2: Mapping — resolve А_ОтчетБаланс_Свод + enum А_ИсточникБаланса (Rule #-1)

**Files:**
- Modify (авто): `_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json`
- Create: `_Rarzrabotki/Olap/Ai_Olap/scripts/introspect_balance_fields.py`

- [ ] **Step 1: Refresh mapping (1С метадані → SQL backend)**

```bash
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b/_Rarzrabotki/Olap/Ai_Olap"
python mapping/refresh_mapping.py
python mapping/refresh_enum_values.py
```
Expected: оновлено `mapping/baserp_storage.json`; з'явився запис для `РегистрСведений.А_ОтчетБаланс_Свод`; перелік `Перечисление.А_ИсточникБаланса` містить 8 значень (ПрочиеАктивыПассивы, РасчетыСКлиентами, РасчетыСПоставщиками, СебестоимостьТоваров, ДенежныеСредстваБезналичные, ДенежныеСредстваНаличные, ПрочиеРасходы, та ін. — усі що є в БД).

- [ ] **Step 2: Написати інтроспектор** `_Rarzrabotki/Olap/Ai_Olap/scripts/introspect_balance_fields.py`

```python
# -*- coding: utf-8 -*-
"""Друкує resolved SQL backend _InfoRg + _Fld для всіх 17 вимірів + 4 ресурсів
+ Регистратор/Период РегистрСведений.А_ОтчетБаланс_Свод. Результат -> у raw_sql Task 3."""
import sys, io, json, pathlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
mp = json.loads(pathlib.Path(__file__).resolve().parents[1].joinpath(
    "mapping", "baserp_storage.json").read_text(encoding="utf-8"))
key = next(k for k in mp if "А_ОтчетБаланс_Свод" in k)
obj = mp[key]
print(f"1С: {key}")
print(f"SQL table: {obj.get('sql_table')}")
print("fields (логічне_ім'я -> SQL колонка):")
for logical, col in obj.get("fields", {}).items():
    print(f"  {logical:<24} -> {col}")
# контроль: усі 21 поле резолвлені
need = ["Организация","Подразделение","Статья","Номенклатура","Контрагент","Партнер",
        "Склад","ОбъектыЭксплуатации","Договор","ФизическоеЛицо","ДенежныеСредства",
        "Source","Аналитика1","Аналитика2","Аналитика3","ОбъектРасчетов","НематериальныйАктив",
        "СуммаНачальныйОстаток","СуммаПриход","СуммаРасход","СуммаКонечныйОстаток"]
miss = [n for n in need if n not in obj.get("fields", {})]
print(f"\nНе резолвлено: {miss if miss else 'НЕМАЄ (усі 21 OK)'}")
assert not miss, f"FAIL: refresh_mapping не дав поля {miss} — повторити Step 1"
print("PASS Task 2 mapping")
```

> Якщо ключ/структура `baserp_storage.json` інша (напр. поля під `column_map` чи інший рівень) — прочитати реальний JSON-блок А_ОтчетБаланс_Свод (`python -c "import json;d=json.load(open(r'...baserp_storage.json',encoding='utf-8'));print(json.dumps([k for k in d if 'Баланс' in k],ensure_ascii=False))"`) і підлаштувати шлях у скрипті під фактичну схему (НЕ вгадувати — дивитись файл).

- [ ] **Step 3: Запустити + зафіксувати мапу**

Run: `python "C:\Configuration_downloads\BASERP25\.claude\worktrees\peaceful-knuth-51295b\_Rarzrabotki\Olap\Ai_Olap\scripts\introspect_balance_fields.py"`
Expected: `SQL table: _InfoRg<N>`, перелік `21` пар `логічне → _Fld<...>`, `PASS Task 2 mapping`.
**Зберегти вивід** — це джерело істини для `raw_sql` у Task 3 (точні `_Fld` імена). Якщо `FAIL: не резолвлено` → повторити Step 1 (mapping не побачив новий регістр).

- [ ] **Step 4: Commit**

```bash
cp "...worktree.../_Rarzrabotki/Olap/Ai_Olap/scripts/introspect_balance_fields.py" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap/Ai_Olap/scripts/introspect_balance_fields.py"
cp "...worktree.../_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json"
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b"
git add _Rarzrabotki/Olap/Ai_Olap/scripts/introspect_balance_fields.py _Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json
git commit -m "feat(olap): refresh_mapping resolve А_ОтчетБаланс_Свод + introspect script"
```

---

## Task 3: ETL pipelines — fact_balance.json + dim_pap_articles.json

**Files:**
- Create: `_Rarzrabotki/Olap/Ai_Olap/pipelines/fact_balance.json`
- Create: `_Rarzrabotki/Olap/Ai_Olap/pipelines/dim_pap_articles.json`

- [ ] **Step 1: Написати `pipelines/fact_balance.json`** (шаблон 1:1 з `fact_cashflow.json`)

У `raw_sql` ПІДСТАВИТИ `_InfoRg<N>` і `_Fld<...>` з виводу Task 2 Step 3 (кожне логічне поле → його `_Fld`-колонка). Структура (приклад; `<...>` замінити РЕАЛЬНИМИ іменами з Task 2, документ-регістратор `_Document<M>` — теж з mapping для `Документ.А_ФинРез_Баланс`):

```json
{
  "pipeline_id": "fact_balance",
  "description": "Fact_Balance idempotent reload з РегистрСведений.А_ОтчетБаланс_Свод. 17 dim + 4 res. Свёртка в 1С (Етапи 1-4), ETL лише копіює. Source enum: А_ИсточникБаланса (8).",
  "schedule": "0 2 1 * *",
  "steps": [{
    "step_id": "load",
    "extractor": {
      "type": "raw_sql",
      "auto_period_params": true,
      "raw_sql": "SELECT r._RecorderRRef AS Recorder_Balance_ID, d._Date_Time AS Period, r.<Fld_Source>RRef AS Source, r.<Fld_Орг>RRef AS Organization_ID, r.<Fld_Подр>RRef AS Department_ID, r.<Fld_Статья>RRef AS PAP_Article_ID, r.<Fld_Номенкл>RRef AS Item_ID, r.<Fld_Контр>RRef AS Counterparty_ID, r.<Fld_Партн>RRef AS Partner_ID, r.<Fld_Склад>RRef AS Warehouse_ID, r.<Fld_ОЭ>RRef AS OperObject_ID, r.<Fld_Догов>RRef AS Contract_ID, r.<Fld_ФЛ>RRef AS Individual_ID, r.<Fld_ДенСр>_RRRef AS Cash_ID, r.<Fld_ОР>RRef AS SettlementObj_ID, r.<Fld_НМА>RRef AS Intangible_ID, r.<Fld_А1> AS Analytics1, r.<Fld_А2> AS Analytics2, r.<Fld_А3> AS Analytics3, r.<Fld_НачОст> AS Sum_Open, r.<Fld_Приход> AS Sum_Inflow, r.<Fld_Расход> AS Sum_Outflow, r.<Fld_КонОст> AS Sum_Close FROM _InfoRg<N> r INNER JOIN _Document<M> d ON d._IDRRef = r._RecorderRRef WHERE r._Active = 0x01 AND r.<Fld_Орг>RRef = 0x80D3000C29BBAC2311E653F06BEE36B2 AND d._Date_Time >= ? AND d._Date_Time < ?"
    },
    "transformer": {
      "steps": ["varbinary_to_uuid", "enum_resolver", "period_offset_fix", "column_mapper"],
      "options": {
        "enum_resolver": {"column_to_enum": {"Source": "Перечисление.А_ИсточникБаланса"}},
        "period_offset_fix": {"source_column": "Period", "period_month_column": "Period_Month"},
        "column_mapper": {"column_map": {
          "Recorder_Balance_ID":"Recorder_Balance_ID","Period":"Period","Period_Month":"Period_Month",
          "Source":"Source","Organization_ID":"Organization_ID","Department_ID":"Department_ID",
          "PAP_Article_ID":"PAP_Article_ID","Item_ID":"Item_ID","Counterparty_ID":"Counterparty_ID",
          "Partner_ID":"Partner_ID","Warehouse_ID":"Warehouse_ID","OperObject_ID":"OperObject_ID",
          "Contract_ID":"Contract_ID","Individual_ID":"Individual_ID","Cash_ID":"Cash_ID",
          "SettlementObj_ID":"SettlementObj_ID","Intangible_ID":"Intangible_ID",
          "Analytics1":"Analytics1","Analytics2":"Analytics2","Analytics3":"Analytics3",
          "Sum_Open":"Sum_Open","Sum_Inflow":"Sum_Inflow","Sum_Outflow":"Sum_Outflow","Sum_Close":"Sum_Close"
        }}
      }
    },
    "loader": {"target_table": "Fact_Balance", "mode": "idempotent_period", "period_column": "Period_Month"}
  }]
}
```

> `_RRRef` (потрійне R) для composite-типів (ДенежныеСредства) — як `Cash_Account_ID` у fact_cashflow.json (`_Fld56068_RRRef`). Точний суфікс брати з Task 2 виводу. Строкові `Аналитика1/2/3` — БЕЗ `RRef`.

- [ ] **Step 2: Написати `pipelines/dim_pap_articles.json`** (Dim_PAP_Articles з реквізитом АктивПассив)

Джерело: `ПланВидовХарактеристик.СтатьиАктивовПассивов` (resolved SQL таблиця з mapping — підставити з Task 2 аналогічно). `AktivPassiv` — з реквізиту `АктивПассив` (enum `ВидыСтатейУправленческогоБаланса`), резолвиться `enum_resolver`.

```json
{
  "pipeline_id": "dim_pap_articles",
  "description": "Dim_PAP_Articles з ПланВидовХарактеристик.СтатьиАктивовПассивов + реквізит АктивПассив (Aktiv/Passiv/AktivPassiv) для розрізу балансу (канон OD-9).",
  "schedule": "0 2 1 * *",
  "steps": [{
    "step_id": "load",
    "extractor": {"type": "raw_sql",
      "raw_sql": "SELECT _IDRRef AS PAP_Article_ID, _Code AS PAP_Article_Code, _Description AS PAP_Article_Name, _ParentIDRRef AS Parent_ID, _Folder AS Is_Group, <Fld_АктивПассив>RRef AS AktivPassiv, _Marked AS Marked_For_Deletion FROM _Chrc<K>"},
    "transformer": {"steps": ["varbinary_to_uuid","enum_resolver","column_mapper"],
      "options": {"enum_resolver": {"column_to_enum": {"AktivPassiv": "Перечисление.ВидыСтатейУправленческогоБаланса"}},
        "column_mapper": {"column_map": {"PAP_Article_ID":"PAP_Article_ID","PAP_Article_Code":"PAP_Article_Code","PAP_Article_Name":"PAP_Article_Name","Parent_ID":"Parent_ID","Is_Group":"Is_Group","AktivPassiv":"AktivPassiv","Marked_For_Deletion":"Marked_For_Deletion"}}}},
    "loader": {"target_table": "Dim_PAP_Articles", "mode": "full_reload"}
  }]
}
```

> `_Chrc<K>` — SQL таблиця ПВХ.СтатьиАктивовПассивов (з mapping Task 2; перевірити чи existing dim_catalogs.json уже має схожий ПВХ-патерн — СтатьиРасходов/СтатьиДоходов — і скопіювати точні стандартні колонки `_Code/_Description/_ParentIDRRef/_Folder/_Marked`). `<Fld_АктивПассив>` — реквізит, з Task 2 виводу для `СтатьиАктивовПассивов`.

- [ ] **Step 3: Валідувати конфіги**

Run: `cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b/_Rarzrabotki/Olap/Ai_Olap" && python main.py --validate`
Expected: `fact_balance` і `dim_pap_articles` у списку валідних pipeline, `0 schema errors`. Помилка схеми → виправити JSON, повторити.

- [ ] **Step 4: Copy back + Commit**

```bash
cp "...worktree.../_Rarzrabotki/Olap/Ai_Olap/pipelines/fact_balance.json" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap/Ai_Olap/pipelines/fact_balance.json"
cp "...worktree.../_Rarzrabotki/Olap/Ai_Olap/pipelines/dim_pap_articles.json" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap/Ai_Olap/pipelines/dim_pap_articles.json"
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b"
git add _Rarzrabotki/Olap/Ai_Olap/pipelines/fact_balance.json _Rarzrabotki/Olap/Ai_Olap/pipelines/dim_pap_articles.json
git commit -m "feat(olap): pipelines fact_balance + dim_pap_articles (шаблон fact_cashflow)"
```

---

## Task 4: Запуск ETL + acceptance (Σ == ПАП/Етап4 до копійки)

**Files:**
- Create: `_Rarzrabotki/Olap/Ai_Olap/tests/test_etl_acceptance_balance.py`

- [ ] **Step 1: Прогнати pipelines за січень 2026**

Run:
```bash
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b/_Rarzrabotki/Olap/Ai_Olap"
python main.py --pipeline dim_pap_articles --run-once
python main.py --pipeline fact_balance --period 2026-01
```
Expected: `dim_pap_articles` завантажив ~N рядків (СтатьиАктивовПассивов); `fact_balance` — рядки за січень (idempotent DELETE WHERE Period_Month + INSERT), `Status=OK` у ETL_Runs. (Якщо CLI-прапорці інші — звірити `python main.py --help`; режим за замовч. `python main.py` робить повний прогон усіх pipeline.)

- [ ] **Step 2: Написати acceptance-тест** `_Rarzrabotki/Olap/Ai_Olap/tests/test_etl_acceptance_balance.py`

```python
# -*- coding: utf-8 -*-
"""Acceptance Fact_Balance: Σ по PAP_Article == ПАП БаERP (січень/ТОВ) до копійки;
Актив=Пасив (Σ Sum_Close ≈ 0). Свёртка в 1С — ETL лише копіює."""
import pyodbc, win32com.client, pythoncom
from datetime import datetime
OLAP = ("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
        "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;")
ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
TOL = 0.01

def test_balance_sum_matches_pap_and_zero():
    cx = pyodbc.connect(OLAP)
    rows = cx.execute("""SELECT d.PAP_Article_Name, SUM(f.Sum_Close)
        FROM Fact_Balance f JOIN Dim_PAP_Articles d ON f.PAP_Article_ID=d.PAP_Article_ID
        WHERE f.Period_Month='2026-01-01' GROUP BY d.PAP_Article_Name""").fetchall()
    olap = {r[0]: float(r[1] or 0) for r in rows}
    total = sum(olap.values()); cx.close()
    assert abs(total) < 1.0, f"Актив!=Пасив: Σ Sum_Close={total:,.2f}"

    pythoncom.CoInitialize()
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(ERP)
    q0 = conn.NewObject("Запрос")
    q0.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации ГДЕ КодПоЕДРПОУ="40645273"'
    s=q0.Выполнить().Выбрать(); s.Следующий(); org=s.С
    qb=conn.NewObject("Запрос")
    qb.Текст="ВЫБРАТЬ НАЧАЛОПЕРИОДА(ДАТАВРЕМЯ(2026,1,1),МЕСЯЦ) КАК НМ, КОНЕЦПЕРИОДА(ДАТАВРЕМЯ(2026,1,31,23,59,59),ДЕНЬ) КАК КД"
    rb=qb.Выполнить().Выбрать(); rb.Следующий()
    qi=conn.NewObject("Запрос")
    qi.Текст='ВЫБРАТЬ Ссылка КАК С ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование В (&Сп)'
    a=conn.NewObject("Массив")
    for n in ["Собственные средства","Доходы текущего периода","Расходы текущего периода"]: a.Добавить(n)
    qi.УстановитьПараметр("Сп",a); искл=qi.Выполнить().Выгрузить().ВыгрузитьКолонку("С")
    qp=conn.NewObject("Запрос")
    qp.Текст=("ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Б.Статья) КАК Н, СУММА(Б.СуммаКонечныйОстаток) КАК sK "
      "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(&Д1,&Д2,Авто,,"
      "Организация=&Орг И НЕ Статья В ИЕРАРХИИ(&Искл)) КАК Б "
      "СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(Б.Статья) ИМЕЮЩИЕ СУММА(Б.СуммаКонечныйОстаток)<>0")
    qp.УстановитьПараметр("Д1",rb.НМ); qp.УстановитьПараметр("Д2",rb.КД)
    qp.УстановитьПараметр("Орг",org); qp.УстановитьПараметр("Искл",искл)
    t=qp.Выполнить().Выгрузить()
    pap={str(t.Получить(i).Н): float(t.Получить(i).sK or 0) for i in range(t.Количество())}
    diffs=[(k,olap.get(k,0),pap.get(k,0)) for k in set(pap)
           if abs(olap.get(k,0)-pap.get(k,0))>1.0]   # ОТ розклад: допуск 1.0₴
    assert not diffs, f"Fact_Balance != ПАП по статтях: {diffs[:5]}"
```

- [ ] **Step 3: Запустити acceptance**

Run: `cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b/_Rarzrabotki/Olap/Ai_Olap" && python -m pytest tests/test_etl_acceptance_balance.py -v`
Expected: `PASSED`. **FAIL «Fact_Balance != ПАП»** → діагностувати: (а) refresh_mapping не оновив `_Fld` (Task 2); (б) фільтр організації UUID невірний; (в) period_offset_fix зсув. Виправити pipeline → Task 4 Step 1 → Step 3.

- [ ] **Step 4: Row counts + Commit**

Run: `python -c "import pyodbc;c=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;');print('Fact_Balance',c.execute('SELECT COUNT(*) FROM Fact_Balance').fetchval(),'Dim_PAP_Articles',c.execute('SELECT COUNT(*) FROM Dim_PAP_Articles').fetchval())"`
Expected: `Fact_Balance >0`, `Dim_PAP_Articles >0`.
```bash
cp "...worktree.../_Rarzrabotki/Olap/Ai_Olap/tests/test_etl_acceptance_balance.py" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap/Ai_Olap/tests/test_etl_acceptance_balance.py"
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b"
git add _Rarzrabotki/Olap/Ai_Olap/tests/test_etl_acceptance_balance.py
git commit -m "test(olap): acceptance Fact_Balance Σ==ПАП січень + Актив=Пасив (PASS)"
```

---

## Task 5: PL.pbix — модель балансу (MCP powerbi-modeling-mcp)

> ⚠️ PL.pbix має бути ВІДКРИТИЙ у Power BI Desktop і підключений до MCP. Power BI MCP tools — deferred: спершу `ToolSearch` query `"powerbi-modeling"` max_results 30 (завантажити весь toolkit). Усі зміни — через MCP; наприкінці `model_operations Refresh refreshType=Full` + Save .pbix.

**Files:** `PL.pbix` (live, через MCP).

- [ ] **Step 1: Завантажити MCP toolkit + перевірити з'єднання**

`ToolSearch` query `"powerbi-modeling"` max_results 30. Потім `connection_operations` (list / connect до відкритого PL.pbix). `model_operations` info → переконатися модель = Fact_PnL + Fact_Cashflow + Dim + Calendar + Table_Measures (compatibility 1600).

- [ ] **Step 2: Додати Power Query партиції Fact_Balance + Dim_СтатьиАктивовПассивов**

`table_operations` create з Power Query M (Import, Sql.Database localhost/OlapBASERP, як існуючі):
- Партиція `Fact_Balance`: `let S=Sql.Database("localhost","OlapBASERP"){[Schema="dbo",Item="Fact_Balance"]}[Data] in S`. Модельна назва таблиці `Fact_Balance`.
- Партиція `Dim_PAP_Articles`: аналогічно `Item="Dim_PAP_Articles"`. Модельна назва **`СтатьиАктивовПассивов`** (1С-нотація, як інші Dim).
Auth: SQL Server, `sa`/`Brw739182465!`.

- [ ] **Step 3: Зв'язки (relationships) Fact_Balance → Dim** (Many-to-One, single filter)

`relationship_operations` create:
- `Fact_Balance[PAP_Article_ID]` → `СтатьиАктивовПассивов[PAP_Article_ID]`
- `Fact_Balance[Organization_ID]` → `Организации[Organization_ID]`
- `Fact_Balance[Department_ID]` → `СтруктураПредприятия[Department_ID]`
- `Fact_Balance[Counterparty_ID]` → `Контрагенты[Counterparty_ID]`
- `Fact_Balance[Partner_ID]` → `Партнеры[Partner_ID]`
- `Fact_Balance[Item_ID]` → `Номенклатура[Item_ID]`
- `Fact_Balance[Individual_ID]` → `ФизическиеЛица[Individual_ID]`
- `Fact_Balance[SettlementObj_ID]` → (нова Dim_ОбъектыРасчетов якщо є; інакше лишити без зв'язку — drill-down пізніше)
- `Fact_Balance[Period_Month]` → `Calendar[Date_Key]`
(Warehouse/OperObject/Contract/Cash/Intangible — без зв'язку поки немає відповідних Dim; додати у наступному циклі. НЕ блокує.)

- [ ] **Step 4: DAX-міри у Table_Measures** (`measure_operations` create)

```dax
[Баланс Вх]        = SUM(Fact_Balance[Sum_Open])
[Баланс Прихід]    = SUM(Fact_Balance[Sum_Inflow])
[Баланс Розхід]    = SUM(Fact_Balance[Sum_Outflow])
[Баланс Вих]       = SUM(Fact_Balance[Sum_Close])
[Перевірка обороту]= [Баланс Вих] - ([Баланс Вх] + [Баланс Прихід] - [Баланс Розхід])
[Актив]   = CALCULATE([Баланс Вих], 'СтатьиАктивовПассивов'[AktivPassiv]="Aktiv")
          + CALCULATE([Баланс Вих], 'СтатьиАктивовПассивов'[AktivPassiv]="AktivPassiv", Fact_Balance[Sum_Close] > 0)
[Пассив]  = -CALCULATE([Баланс Вих], 'СтатьиАктивовПассивов'[AktivPassiv]="Passiv")
          - CALCULATE([Баланс Вих], 'СтатьиАктивовПассивов'[AktivPassiv]="AktivPassiv", Fact_Balance[Sum_Close] < 0)
[Контроль Актив=Пассив] = [Актив] - [Пассив]
```
(Двосторонні `АктивПассив` як `Налоги` — діляться по знаку `Sum_Close`, канон OD-9.)

- [ ] **Step 5: Сторінка «Баланс»** (дзеркало сторінки Cashflow з `olap_cashflow_dashboard_design_2026_05.md`)

Створити звітну сторінку: matrix `СтатьиАктивовПассивов`(рядки, drill Parent→child) × міри `[Баланс Вх]/[Прихід]/[Розхід]/[Вих]`; KPI-карти `[Актив]`, `[Пассив]`, `[Контроль Актив=Пассив]` (має бути ≈0); слайсери `Calendar[Year_Month]`, `СтатьиАктивовПассивов[AktivPassiv]`, `Fact_Balance[Source]` (drill-down джерела), `Организации`. (Візуали додаються вручну у Power BI Desktop або через MCP якщо підтримує; зафіксувати у olap_powerbi_pl_pbix.md.)

- [ ] **Step 6: Refresh + Save + acceptance DAX**

`model_operations Refresh refreshType=Full`. Потім `dax_query_operations`:
```dax
EVALUATE ROW("Контроль", [Контроль Актив=Пассив], "ВихСальдо", [Баланс Вих])
```
Expected: `Контроль ≈ 0`, `ВихСальдо ≈ 0` (січень закритий, Актив=Пасив). Зберегти PL.pbix (Power BI Desktop Ctrl+S — MCP не зберігає файл сам).

- [ ] **Step 7: Commit (тести/доки; .pbix бінарний — окремо)**

```bash
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b"
git add -A && git commit -m "feat(olap): PL.pbix +Fact_Balance +СтатьиАктивовПассивов +DAX баланс +сторінка"
```

---

## Task 6: Оновити knowledge_Olap + KNOWLEDGE_MAP

**Files:** `_Rarzrabotki/notebook/knowledge_Olap/{olap_sql_schema.md, olap_etl_pipeline.md, olap_powerbi_pl_pbix.md, KNOWLEDGE_MAP_OLAP.md}`

- [ ] **Step 1:** `olap_sql_schema.md` — додати DDL `Fact_Balance` + `Dim_PAP_Articles`, оновити лічильник таблиць (24→26) і список verify.
- [ ] **Step 2:** `olap_etl_pipeline.md` — додати pipelines `fact_balance` + `dim_pap_articles`, правило refresh_mapping перед прогоном.
- [ ] **Step 3:** `olap_powerbi_pl_pbix.md` — нова Fact (Fact_Balance), нова Dim (СтатьиАктивовПассивов з AktivPassiv), зв'язки, DAX-міри балансу, сторінка «Баланс».
- [ ] **Step 4:** `KNOWLEDGE_MAP_OLAP.md` — новий рядок Stage «Balance» (DONE, дата 2026-05-16), row counts Fact_Balance/Dim_PAP_Articles, acceptance-еталон січень.
- [ ] **Step 5: cp у main + Commit**

```bash
for f in olap_sql_schema olap_etl_pipeline olap_powerbi_pl_pbix KNOWLEDGE_MAP_OLAP; do
 cp "...worktree.../_Rarzrabotki/notebook/knowledge_Olap/$f.md" "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Olap/$f.md"; done
cd "C:/Configuration_downloads/BASERP25/.claude/worktrees/peaceful-knuth-51295b"
git add _Rarzrabotki/notebook/knowledge_Olap
git commit -m "knowledge_Olap: Balance Stage — Fact_Balance pipeline + PL.pbix модель"
```

---

## Acceptance-еталон (січень 2026 / ТОВ ІНДАСТРІАЛБУД)

- `Σ Fact_Balance.Sum_Close` по кожній `PAP_Article` == ПАП/Етап4 до копійки (tol 0.01; «Оплата труда» tol 1.0 — розклад по фізлицю округлення).
- `Σ Fact_Balance.Sum_Close` (весь набір) ≈ 0 → Актив=Пасив.
- `[Контроль Актив=Пассив]` у PL.pbix ≈ 0.
- Контрольна цифра 289 064 974,43 — січень/ВСІ організації (для ТОВ — своя; критерій = збіг з ПАП, не абсолют).
- pytest `tests/test_etl_acceptance_balance.py` PASS; регрес `tests/test_etl_acceptance_globyno2.py` (PnL) PASS — Fact_PnL/Cashflow не зачеплені.

---

## Self-Review

**1. Spec coverage (канон §10/Roadmap + OD-9):**
- SQL Fact_Balance (4 res + 17 dim) → Task 1 ✅
- Dim_Статья з атрибутом АктивПассив (OD-9 розріз downstream) → Task 1 (Dim_PAP_Articles.AktivPassiv) + Task 3 dim_pap_articles.json + Task 5 DAX [Актив]/[Пассив] ✅
- ETL перенос (свёртка в 1С, ETL копіює) → Task 2 (refresh_mapping) + Task 3 (fact_balance.json дзеркало fact_cashflow) + Task 4 ✅
- DAX [Opening]/[Inflow]/[Outflow]/[Closing] → Task 5 Step 4 ([Баланс Вх/Прихід/Розхід/Вих]) ✅
- сторінка Balance поряд із PL/DDS → Task 5 Step 5 ✅
- drill-down по Source → Task 5 Step 5 (слайсер Source) ✅
- Acceptance Σ==ПАП/Етап4 + Актив=Пасив → Task 4 + Task 5 Step 6 + еталон-секція ✅
- Rule-#-1 (resolve перед raw_sql) → Task 2 перед Task 3 ✅
- Типові 1С не міняти → лише читання SQL backend (Передумови 4) ✅

**2. Placeholder scan:** Конкретні DDL/JSON/Python/DAX/команди з очікуваним виводом. `<Fld_*>`/`_InfoRg<N>`/`_Document<M>`/`_Chrc<K>` у Task 3 — НЕ плейсхолдери, а **обов'язкова підстановка з конкретного виводу Task 2** (Rule-#-1 discovery-first; вгадувати _Fld номери заборонено каноном/memory). Task 2 дає точну мапу + assert що всі 21 поле резолвлені.

**3. Type/naming consistency:** `Fact_Balance`/`Dim_PAP_Articles`/`PAP_Article_ID`/`Sum_Open|Inflow|Outflow|Close`/`AktivPassiv` однакові у Task 1 (DDL) ↔ Task 3 (pipeline column_map) ↔ Task 4 (acceptance) ↔ Task 5 (PBI relationships/DAX). Модельна назва Dim у PBI = `СтатьиАктивовПассивов` (1С-нотація, консистентно з іншими Dim). `Source` enum = `Перечисление.А_ИсточникБаланса` скрізь.

Гепів немає (Task 3 _Fld — discovery-first за дизайном пайплайну, Task 2 його забезпечує). План готовий.
