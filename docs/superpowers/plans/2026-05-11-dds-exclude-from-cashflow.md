# А_ИсключатьИзОтчетаCashflow — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Прокинути новий 1С Boolean-реквізит `СтатьиДвиженияДенежныхСредств.А_ИсключатьИзОтчетаCashflow` через ETL у `Dim_DDS_Articles.Is_Excluded_From_Cashflow` (OlapBASERP) і далі у Power BI модель PL.pbix; на сторінках Cashflow налаштувати page-level filter `=False` + slicer override.

**Architecture:** 3-шарова інтеграція (1С → SQL OlapBASERP → Power BI Tabular) з фільтруванням на рівні Layout сторінок Power BI (варіант B: page-filter + slicer, без зміни базових DAX-мір).

**Tech Stack:** 1C COM (`V83.COMConnector` via `pywin32`), Ai_Olap ETL (Python pipeline JSON + `main.py --run-once`), SQL Server `OlapBASERP` (T-SQL DDL), Power BI Desktop UI + `mcp__powerbi-modeling-mcp` для rename, `mcp__1c-workerp` для діагностики 1С.

**Reference spec:** [docs/superpowers/specs/2026-05-11-dds-exclude-from-cashflow-design.md](../specs/2026-05-11-dds-exclude-from-cashflow-design.md)

---

## File Structure

| Файл | Дія | Призначення |
|------|-----|-------------|
| `_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json` | Regenerate | Виявлення нового `_Fld<NNN>` через `refresh_mapping.py` |
| (Ad-hoc SQL via MCP) | DDL | `ALTER TABLE Dim_DDS_Articles ADD Is_Excluded_From_Cashflow bit` |
| `_Rarzrabotki/Olap/Ai_Olap/pipelines/dim_dds_articles.json` | Modify | Додати SELECT + column_mapper |
| `_Rarzrabotki/Olap/PowerBi/PL.pbix` | Modify via UI | Power Query Refresh + rename column |
| `_Rarzrabotki/Olap/PowerBi/PL.pbix` | Modify via UI | Page-filter + Slicer на 3-4 Cashflow-сторінках |
| `_Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md` | Update | DDL `Dim_DDS_Articles` + нова колонка |
| `_Rarzrabotki/notebook/knowledge_Olap/olap_powerbi_pl_pbix.md` | Update | §3.1 структура моделі, STATUS Stage v3.5 |
| `_Rarzrabotki/notebook/knowledge_Olap/olap_changelog_2026_05.md` | Update | Новий §Stage v3.5 |

---

## Task 1: Перевірити що реквізит у 1С Designer і знайти SQL колонку `_Fld<NNN>`

**Files:**
- Create: `_Rarzrabotki/Python/test/test_dds_exclude_attribute_exists.py`

- [ ] **Step 1: Створити Python-тест перевірки**

```python
# -*- coding: utf-8 -*-
"""
Перевірка що реквізит А_ИсключатьИзОтчетаCashflow доданий до
Справочник.СтатьиДвиженияДенежныхСредств у конфігурації 1С BAS ERP.
"""
import sys, io
import win32com.client, pythoncom

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)

mdo = erp.Метаданные.Справочники.НайтиПоИмени("СтатьиДвиженияДенежныхСредств")
assert mdo is not None, "Справочник не знайдено"

found = None
for rekv in mdo.Реквизиты:
    if str(rekv.Имя) == "А_ИсключатьИзОтчетаCashflow":
        found = rekv
        break

assert found, "Реквізит А_ИсключатьИзОтчетаCashflow ще НЕ доданий у конфігурацію — потрібно зробити в 1С Designer"
print(f"OK: реквізит знайдено, тип={found.Тип}")
print(f"     Синонім={found.Синоним}")
```

- [ ] **Step 2: Запустити тест**

```bash
python "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_dds_exclude_attribute_exists.py"
```

Expected output:
```
OK: реквізит знайдено, тип=...
     Синонім=...
```

Якщо `AssertionError` — реквізит ще не доданий у 1С Designer; STOP plan і повернутись до користувача.

- [ ] **Step 3: Запит до 1С SQL backend щоб знайти `_Fld<NNN>` мепінг**

Через `mcp__1c-workerp__get_metadata_structure`:

```javascript
get_metadata_structure({
  metaType: "Catalogs",
  name: "СтатьиДвиженияДенежныхСредств"
})
```

У відповіді знайти секцію Attributes → `А_ИсключатьИзОтчетаCashflow`. Це поверне ідентифікатор UUID, але не SQL `_Fld<NNN>`. Для SQL ідентифікатора:

```bash
sqlcmd -S SQLSERVER -d BaseERP -U sa -P "Brw739182465!" -Q "SELECT TOP 0 * FROM _Reference529" -y 0 -h -1
```

(або через SSMS) — побачимо схему таблиці. Шукати колонку чий бінарний eq схожий на `0x01`/`0x00` (Boolean). Або точніше через INFORMATION_SCHEMA після `refresh_mapping.py` (Task 2).

- [ ] **Step 4: Commit Python-тест**

```bash
cd C:\Configuration_downloads\BASERP25
git add _Rarzrabotki/Python/test/test_dds_exclude_attribute_exists.py
git commit -m "test: verify А_ИсключатьИзОтчетаCashflow attribute exists in 1С config"
```

---

## Task 2: Regenerate ETL mapping (`baserp_storage.json`)

**Files:**
- Modify: `_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json` (auto-regenerated)

- [ ] **Step 1: Backup поточного mapping**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap
cp mapping/baserp_storage.json mapping/baserp_storage.json.bak
```

- [ ] **Step 2: Запустити refresh_mapping.py**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap
python mapping/refresh_mapping.py
```

Expected: скрипт виконається без помилок, `baserp_storage.json` оновиться.

- [ ] **Step 3: Перевірити що нове поле виявлено**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap
python -c "import json; m=json.load(open('mapping/baserp_storage.json', encoding='utf-8')); art=m.get('СтатьиДвиженияДенежныхСредств', m.get('Справочник.СтатьиДвиженияДенежныхСредств', {})); print('Fields:'); [print(' ', k, '→', v) for k,v in (art.get('fields') or art.get('Реквизиты') or {}).items() if 'Cashflow' in str(k)]"
```

Expected output (точний `_FldNNNNN` номер залежить від конфігурації):
```
Fields:
  А_ИсключатьИзОтчетаCashflow → _Fld<NNNNN>
```

Зафіксувати точний `_FldNNNNN` номер — потрібен для Task 4.

- [ ] **Step 4: Diff vs backup**

```bash
diff mapping/baserp_storage.json.bak mapping/baserp_storage.json
```

Expected: лише одна додаткова строка з `А_ИсключатьИзОтчетаCashflow → _Fld<NNNNN>` (плюс можливо інші reorderings — переконатись що нічого важливого не зникло).

- [ ] **Step 5: Видалити backup і commit**

```bash
rm mapping/baserp_storage.json.bak
cd C:\Configuration_downloads\BASERP25
git add _Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json
git commit -m "mapping: regenerate baserp_storage.json with new А_ИсключатьИзОтчетаCashflow field"
```

---

## Task 3: SQL DDL — додати колонку `Is_Excluded_From_Cashflow` у `Dim_DDS_Articles`

**Files:**
- Modify: SQL OlapBASERP (ad-hoc DDL)
- Update: `_Rarzrabotki/Olap/Ai_Olap/sql/02_dim_tables.sql` (синхронізувати з production)

- [ ] **Step 1: Перевірити поточну схему `Dim_DDS_Articles`**

Через SSMS або psqlcmd:

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Dim_DDS_Articles' ORDER BY ORDINAL_POSITION"
```

Expected: 10-12 колонок (`DDS_Article_ID`, `DDS_Article_Code`, `Name`, `Parent_ID`, `Is_Group`, `Marked_For_Deletion`, `CFS_Section`, `Hierarchy_Path`, `Hierarchy_Depth`, `Level1..Level5`, `Loaded_At`). Колонки `Is_Excluded_From_Cashflow` НЕ повинно бути.

- [ ] **Step 2: ALTER TABLE додати колонку**

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "ALTER TABLE Dim_DDS_Articles ADD Is_Excluded_From_Cashflow bit NOT NULL DEFAULT 0"
```

Expected output: `Команды успешно выполнены.` (без повідомлень про помилки)

- [ ] **Step 3: Створити індекс для page-filter (часто фільтруватиметься)**

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "CREATE INDEX IX_DDS_Articles_Excluded ON Dim_DDS_Articles(Is_Excluded_From_Cashflow) INCLUDE (DDS_Article_ID, Name)"
```

- [ ] **Step 4: Перевірити що колонка створена і базово заповнена 0**

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "SELECT COUNT(*) AS total, SUM(CAST(Is_Excluded_From_Cashflow AS int)) AS excluded FROM Dim_DDS_Articles WHERE Marked_For_Deletion = 0"
```

Expected: `total = ~425`, `excluded = 0` (бо за замовчуванням FALSE; ETL ще не запускався).

- [ ] **Step 5: Оновити DDL-файл у репо**

Read `_Rarzrabotki/Olap/Ai_Olap/sql/02_dim_tables.sql`, знайти секцію `CREATE TABLE Dim_DDS_Articles`, додати рядок після `CFS_Section`:

```sql
    CFS_Section varchar(15) NULL,
    Is_Excluded_From_Cashflow bit NOT NULL DEFAULT 0,
    Hierarchy_Path nvarchar(500) NULL,
```

- [ ] **Step 6: Commit DDL update**

```bash
cd C:\Configuration_downloads\BASERP25
git add _Rarzrabotki/Olap/Ai_Olap/sql/02_dim_tables.sql
git commit -m "sql: add Is_Excluded_From_Cashflow column to Dim_DDS_Articles DDL"
```

---

## Task 4: Оновити ETL pipeline `dim_dds_articles.json`

**Files:**
- Modify: `_Rarzrabotki/Olap/Ai_Olap/pipelines/dim_dds_articles.json` (або `dim_catalogs.json` якщо step там)

- [ ] **Step 1: Знайти pipeline-файл**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\pipelines
ls -la *.json
```

Шукати файл де є `"step_id": "dim_dds_articles"` або таблиця `Dim_DDS_Articles`:

```bash
grep -l "Dim_DDS_Articles\|dim_dds_articles" *.json
```

- [ ] **Step 2: Прочитати поточний raw_sql**

Через `Read tool` файл який знайдено у Step 1. Знайти секцію `extractor.sql` для DDS Articles. Має бути SELECT з `_Reference529` (або з view).

- [ ] **Step 3: Modify raw_sql — додати поле в SELECT**

Замінити (приклад — точні рядки залежать від реального вмісту):

```sql
    -- old:
    CAST(a._Code AS varchar(50)) AS DDS_Article_Code,
    a._Description AS Name,
    CAST(a._Fld<XXX> AS varchar(15)) AS CFS_Section,
```

На:

```sql
    -- new:
    CAST(a._Code AS varchar(50)) AS DDS_Article_Code,
    a._Description AS Name,
    CAST(a._Fld<XXX> AS varchar(15)) AS CFS_Section,
    CAST(CASE WHEN a._Fld<NNNNN> = 0x00 THEN 0 ELSE 1 END AS bit) AS Is_Excluded_From_Cashflow,
```

Де `<NNNNN>` — номер з Task 2 Step 3.

- [ ] **Step 4: Modify column_mapper — додати маппінг**

Знайти секцію `transformer.column_mapper.mapping` (або `columns`):

```json
{
    "DDS_Article_ID": "DDS_Article_ID",
    "DDS_Article_Code": "DDS_Article_Code",
    "Name": "Name",
    ...
    "CFS_Section": "CFS_Section",
    "Is_Excluded_From_Cashflow": "Is_Excluded_From_Cashflow",
    "Hierarchy_Path": "Hierarchy_Path",
    ...
}
```

- [ ] **Step 5: Перевірити валідність JSON**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap
python -c "import json; json.load(open('pipelines/dim_dds_articles.json', encoding='utf-8')); print('JSON valid')"
```

Expected: `JSON valid`

- [ ] **Step 6: Commit**

```bash
cd C:\Configuration_downloads\BASERP25
git add _Rarzrabotki/Olap/Ai_Olap/pipelines/dim_dds_articles.json
git commit -m "etl: add Is_Excluded_From_Cashflow to dim_dds_articles pipeline"
```

---

## Task 5: Запустити ETL full reload + валідація

**Files:**
- Дані у SQL `OlapBASERP.Dim_DDS_Articles`

- [ ] **Step 1: Запустити ETL для `dim_dds_articles`**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap
python main.py --run-once dim_dds_articles
```

Expected output:
```
[OK] Extracted N rows from _Reference529
[OK] Loaded N rows into Dim_DDS_Articles
[OK] step dim_dds_articles completed
```

Де `N ≈ 426`.

- [ ] **Step 2: Verify SQL-рівнем — є колонка і заповнена для частини рядків**

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "SELECT TOP 10 DDS_Article_ID, Name, Is_Excluded_From_Cashflow FROM Dim_DDS_Articles WHERE Marked_For_Deletion = 0 ORDER BY Is_Excluded_From_Cashflow DESC, Name"
```

Expected: 10 рядків, перші мають `Is_Excluded_From_Cashflow = 1` якщо фінансист уже відмітив якісь статті в 1С; решта `0`. Якщо `excluded = 0` всюди — це нормально (просто треба заповнити в 1С).

- [ ] **Step 3: Sanity-check загальний counter**

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "SELECT COUNT(*) AS total, SUM(CAST(Is_Excluded_From_Cashflow AS int)) AS excluded FROM Dim_DDS_Articles WHERE Marked_For_Deletion = 0"
```

Записати числа в коммент Task 6 (для документації).

- [ ] **Step 4: Перевірити що Fact_Cashflow посилається коректно**

```bash
sqlcmd -S SQLSERVER -d OlapBASERP -U sa -P "Brw739182465!" -Q "SELECT TOP 5 fc.DDS_Article_ID, da.Name, da.Is_Excluded_From_Cashflow FROM Fact_Cashflow fc JOIN Dim_DDS_Articles da ON fc.DDS_Article_ID = da.DDS_Article_ID WHERE da.Is_Excluded_From_Cashflow = 1"
```

Expected: якщо є виключені статті — побачимо їх Fact-рядки; якщо ні — `0 rows affected` (також ок).

---

## Task 6: Power BI — підтягнути нову колонку через Power Query Refresh

**Files:**
- Modify via UI: `_Rarzrabotki/Olap/PowerBi/PL.pbix`

**Передумова:** PL.pbix відкритий у Power BI Desktop.

- [ ] **Step 1: Відкрити Power Query Editor**

У Power BI Desktop: ribbon Home → **Transform data** → відкривається Power Query Editor у новому вікні.

- [ ] **Step 2: Знайти query `Dim_DDS_Articles`**

У лівій панелі Queries знайти `Dim_DDS_Articles` (а не "СтатьиДвиженияДенежныхСредств" — це назва партиції у Power Query, не назва моделі).

- [ ] **Step 3: Refresh Preview**

Правий клік на `Dim_DDS_Articles` → **Refresh Preview**. Це підтягне нову колонку зі SQL.

У правій панелі Applied Steps підсвітити Source — у Preview має з'явитися колонка `Is_Excluded_From_Cashflow` (Boolean — True/False).

- [ ] **Step 4: Перевірити що нова колонка має правильний тип**

У стовпчику `Is_Excluded_From_Cashflow` заголовок повинен містити іконку `True/False` (Boolean). Якщо там `Whole Number` (1/0) — клік на іконку → змінити на `True/False`.

- [ ] **Step 5: Close & Apply**

Ribbon Home → **Close & Apply**. Power Query Editor закриється, модель оновиться.

- [ ] **Step 6: Перевірити що колонка з'явилась у моделі**

Через `mcp__powerbi-modeling-mcp__connection_operations` (Connect → знайти localhost:port → Connect):

```
connection_operations.ListLocalInstances
connection_operations.Connect (connectionString: "Data Source=localhost:<port>")
```

Потім перевірити колонки таблиці:

```
column_operations.List (filter: { tableNames: ["СтатьиДвиженияДенежныхСредств"] })
```

Expected: колонка `Is_Excluded_From_Cashflow` присутня.

- [ ] **Step 7: Перейменувати на display name `А_ИсключатьИзОтчетаCashflow`**

Через MCP:

```
column_operations.Rename
  renameDefinitions: [{
    tableName: "СтатьиДвиженияДенежныхСредств",
    currentName: "Is_Excluded_From_Cashflow",
    newName: "А_ИсключатьИзОтчетаCashflow"
  }]
```

Verify через `column_operations.Get`:

```
column_operations.Get
  references: [{ tableName: "СтатьиДвиженияДенежныхСредств", name: "А_ИсключатьИзОтчетаCashflow" }]
```

Expected: повертає колонку з `sourceColumn = "Is_Excluded_From_Cashflow"`, `dataType = "Boolean"`, `isHidden = false`.

- [ ] **Step 8: Save PL.pbix**

У Power BI Desktop: File → **Save** (або Ctrl+S).

---

## Task 7: Page-filter + Slicer на сторінках Cashflow

**Files:**
- Modify via UI: `_Rarzrabotki/Olap/PowerBi/PL.pbix` Report Layout

**Цільові сторінки** (4):
1. `Cashflow`
2. `Дублікат Cashflow місяць(казна)`
3. `Дублікат Cashflow об'єкт (казна) по місячно`
4. `Дублікат Cashflow об'єкт (казна) по місяч...` (другий дублікат якщо існує)

**Для КОЖНОЇ сторінки повторити Steps 1-7:**

- [ ] **Step 1: Відкрити цільову сторінку**

У Power BI Desktop внизу — таб з назвою сторінки. Кліком переключитись.

- [ ] **Step 2: Зняти виділення з усіх візуалів**

Клік на пусте місце canvas (не на візуал) щоб виділити сторінку.

- [ ] **Step 3: Додати Page-level filter**

У Filters pane (праворуч):
1. Знайти секцію **"Filters on this page"** (друга секція згори)
2. У панелі Fields (Дані) розгорнути `Table_Measures` → не туди! Розгорнути таблицю **`СтатьиДвиженияДенежныхСредств`**
3. Перетягнути поле **`А_ИсключатьИзОтчетаCashflow`** у "Filters on this page"
4. У dropdown фільтра: тип `Basic filtering`
5. Зняти галку з `True`, залишити галку на `False`
6. Натиснути pin (закріпити) щоб користувач не міг прибрати фільтр випадково

- [ ] **Step 4: Додати Slicer для override**

1. Insert → **Visualizations** → клік на іконку **Slicer** (вертикальні смужки)
2. Не клікати на canvas — новий пустий slicer з'являється
3. У панелі Visualizations → Fields → перетягнути `СтатьиДвиженияДенежныхСредств[А_ИсключатьИзОтчетаCashflow]`
4. Формат slicer (Format pane → Slicer settings):
   - Style: **Vertical list** (за замовчуванням) або **Dropdown** для компактності
   - Selection controls: enable "Show 'Select all' option"
5. Перейменувати заголовок: Format → Header → Text: `Показати виключені`

- [ ] **Step 5: Розмістити slicer**

Перемістити slicer у правий верхній кут сторінки (~150×80 px). Не накладати на існуючі візуали.

- [ ] **Step 6: Перевірити взаємодію slicer і фільтра**

1. У slicer натиснути на `True` — повинні з'явитися виключені статті у матрицях сторінки
2. Зняти галку — статті знову мають зникнути
3. Прибрати всі галки в slicer — поведінка має повернутись до page-filter (тільки False показано)

- [ ] **Step 7: Save PL.pbix**

File → Save (Ctrl+S).

**Після всіх 4 сторінок:**

- [ ] **Step 8: Final save + git commit (бінарного pbix не комітимо, тільки метадані)**

```bash
cd C:\Configuration_downloads\BASERP25
git status _Rarzrabotki/Olap/PowerBi/PL.pbix
```

Бінарний `.pbix` файл буде в `git status`, але не комітимо його — занадто великий. Лише doc-зміни (Task 8).

---

## Task 8: Оновити документацію

**Files:**
- Modify: `_Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md`
- Modify: `_Rarzrabotki/notebook/knowledge_Olap/olap_powerbi_pl_pbix.md`
- Modify: `_Rarzrabotki/notebook/knowledge_Olap/olap_changelog_2026_05.md`

- [ ] **Step 1: Оновити `olap_sql_schema.md`**

Знайти секцію `Dim_DDS_Articles` (грепом). У DDL знайти `CFS_Section varchar(15) NULL,` і додати після:

```sql
CFS_Section varchar(15) NULL,
Is_Excluded_From_Cashflow bit NOT NULL DEFAULT 0,    -- Stage v3.5 (2026-05-11): з 1С реквізиту А_ИсключатьИзОтчетаCashflow
Hierarchy_Path nvarchar(500) NULL,
```

- [ ] **Step 2: Оновити `olap_powerbi_pl_pbix.md` STATUS-заголовок**

Знайти STATUS-блок і додати рядок:

```markdown
> **STATUS:** 🔄 IN PROGRESS — **Stage v3.5 (2026-05-11)** додано колонку `Is_Excluded_From_Cashflow` (Power BI alias `А_ИсключатьИзОтчетаCashflow`) у `Dim_DDS_Articles` з 1С реквізиту. На сторінках Cashflow застосовано page-level filter `=False` + slicer для override.
```

- [ ] **Step 3: Оновити §3.1 у `olap_powerbi_pl_pbix.md` (структура таблиці)**

Знайти таблицю опису `СтатьиДвиженияДенежныхСредств` (9 колонок per docs). Оновити на 10 колонок — додати рядок:

```markdown
| `СтатьиДвиженияДенежныхСредств` | 10 (Stage v3.5: +Is_Excluded_From_Cashflow alias `А_ИсключатьИзОтчетаCashflow`) | Dim_DDS_Articles | Dim_DDS_Articles |
```

- [ ] **Step 4: Додати запис у `olap_changelog_2026_05.md`**

Знайти секцію TODO (наприкінці файлу) і ВСТАВИТИ перед нею новий розділ:

```markdown
## Stage v3.5 (2026-05-11) — Is_Excluded_From_Cashflow integration

### Симптом / Мотивація
Фінансист потребує тимчасово приховувати окремі статті ДДС (наприклад внутрішньогрупові обороти) у звітах Cashflow без видалення їх з первинних даних.

### Зміни
- **1С:** додано Boolean-реквізит `Справочник.СтатьиДвиженияДенежныхСредств.А_ИсключатьИзОтчетаCashflow` (зроблено фінансистом)
- **mapping:** `mapping/refresh_mapping.py` виявив нове поле, оновлено `baserp_storage.json`
- **SQL DDL:** `ALTER TABLE Dim_DDS_Articles ADD Is_Excluded_From_Cashflow bit NOT NULL DEFAULT 0` + індекс `IX_DDS_Articles_Excluded`
- **ETL pipeline:** `pipelines/dim_dds_articles.json` — додано поле у raw_sql SELECT (з `CAST(...AS bit)`) і column_mapper
- **Power BI модель:** Power Query Refresh підтягнув нову колонку → rename на display `А_ИсключатьИзОтчетаCashflow` (SQL sourceColumn=`Is_Excluded_From_Cashflow`)
- **Power BI Layout:** на 4 сторінках Cashflow (`Cashflow`, `Дублікат Cashflow місяць(казна)`, `Дублікат Cashflow об'єкт (казна) по місячно` + інші дублікати) додано page-level filter `=False` + slicer `Показати виключені`

### Архітектурне рішення
Обрано **варіант B** (page-filter + slicer) замість A (DAX-фільтр у базових мірах) — для гнучкості та збереження повноти даних у моделі. Базові SUM-міри не змінювались. Користувач може через slicer тимчасово показати виключені статті.

### Перевірка
- SQL: `SELECT COUNT(*), SUM(CAST(Is_Excluded_From_Cashflow AS int)) FROM Dim_DDS_Articles` — повертає total/excluded
- PBI: відмітити статтю в 1С → ETL re-run → Refresh PL.pbix → стаття зникла з матриці; slicer "True" → знову з'являється

### Out of scope
- Fact_PnL фільтрація — окрема задача
- Заповнення значень True для статей — робота фінансиста
```

- [ ] **Step 5: Скопіювати docs з worktree у main config (Rule #4)**

```bash
cd C:\Configuration_downloads\BASERP25
cp .claude/worktrees/<worktree>/_Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md _Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md
cp .claude/worktrees/<worktree>/_Rarzrabotki/notebook/knowledge_Olap/olap_powerbi_pl_pbix.md _Rarzrabotki/notebook/knowledge_Olap/olap_powerbi_pl_pbix.md
cp .claude/worktrees/<worktree>/_Rarzrabotki/notebook/knowledge_Olap/olap_changelog_2026_05.md _Rarzrabotki/notebook/knowledge_Olap/olap_changelog_2026_05.md
```

(Якщо редагували напряму в main config — навпаки, копіювати з main у worktree.)

- [ ] **Step 6: Git commit docs**

```bash
cd C:\Configuration_downloads\BASERP25
git add _Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md \
        _Rarzrabotki/notebook/knowledge_Olap/olap_powerbi_pl_pbix.md \
        _Rarzrabotki/notebook/knowledge_Olap/olap_changelog_2026_05.md \
        docs/superpowers/specs/2026-05-11-dds-exclude-from-cashflow-design.md \
        docs/superpowers/plans/2026-05-11-dds-exclude-from-cashflow.md
git commit -m "docs: Stage v3.5 — А_ИсключатьИзОтчетаCashflow integration through OLAP chain"
```

---

## Task 9: End-to-end acceptance test

**Files:**
- Create: `_Rarzrabotki/Python/test/test_dds_exclude_e2e.py`

- [ ] **Step 1: Створити acceptance test**

```python
# -*- coding: utf-8 -*-
"""
End-to-end перевірка інтеграції А_ИсключатьИзОтчетаCashflow:
1. У 1С реквізит є (повторно)
2. ETL поле успішно прокинуло у SQL Dim_DDS_Articles
3. Power BI модель має нову колонку (через XMLA — опціонально, можна manual check)
"""
import sys, io
import pyodbc

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# === SQL OlapBASERP перевірка ===
cn = pyodbc.connect(
    "Driver={SQL Server};Server=SQLSERVER;Database=OlapBASERP;UID=sa;PWD=Brw739182465!"
)
cur = cn.cursor()

# 1. Колонка є
cur.execute("""
    SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'Dim_DDS_Articles'
      AND COLUMN_NAME = 'Is_Excluded_From_Cashflow'
""")
assert cur.fetchone()[0] == 1, "Колонка Is_Excluded_From_Cashflow не знайдена в Dim_DDS_Articles"
print("[OK] SQL: колонка Is_Excluded_From_Cashflow присутня")

# 2. Тип bit, NOT NULL, DEFAULT 0
cur.execute("""
    SELECT DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'Dim_DDS_Articles'
      AND COLUMN_NAME = 'Is_Excluded_From_Cashflow'
""")
row = cur.fetchone()
assert row.DATA_TYPE == 'bit', f"Тип повинен bit, отримано {row.DATA_TYPE}"
assert row.IS_NULLABLE == 'NO', "Має бути NOT NULL"
print(f"[OK] SQL: тип={row.DATA_TYPE}, nullable={row.IS_NULLABLE}, default={row.COLUMN_DEFAULT}")

# 3. ETL заповнив дані
cur.execute("""
    SELECT COUNT(*) AS total,
           SUM(CAST(Is_Excluded_From_Cashflow AS int)) AS excluded
    FROM Dim_DDS_Articles
    WHERE Marked_For_Deletion = 0
""")
row = cur.fetchone()
print(f"[OK] SQL: total={row.total}, excluded={row.excluded}")
assert row.total > 0, "Dim_DDS_Articles порожня — ETL впав"

# 4. Індекс є
cur.execute("""
    SELECT COUNT(*) FROM sys.indexes
    WHERE name = 'IX_DDS_Articles_Excluded'
""")
assert cur.fetchone()[0] == 1, "Індекс IX_DDS_Articles_Excluded не знайдено"
print("[OK] SQL: індекс IX_DDS_Articles_Excluded створено")

cn.close()
print("\n=== Acceptance E2E test passed ===")
```

- [ ] **Step 2: Запустити test**

```bash
cd C:\Configuration_downloads\BASERP25
python _Rarzrabotki/Python/test/test_dds_exclude_e2e.py
```

Expected:
```
[OK] SQL: колонка Is_Excluded_From_Cashflow присутня
[OK] SQL: тип=bit, nullable=NO, default=((0))
[OK] SQL: total=425, excluded=0  (або >0 якщо фінансист уже заповнив)
[OK] SQL: індекс IX_DDS_Articles_Excluded створено

=== Acceptance E2E test passed ===
```

- [ ] **Step 3: Manual PBI test (не автоматизовується)**

1. У 1С Designer / Enterprise: відкрити будь-яку статтю ДДС, поставити `А_ИсключатьИзОтчетаCashflow = Истина`, записати
2. ETL: `python main.py --run-once dim_dds_articles`
3. У Power BI Desktop: відкрити PL.pbix → Home → Refresh → wait
4. Перейти на сторінку `Cashflow` → знайти цю статтю у матрицях → **має бути приховано**
5. У slicer "Показати виключені" → клік на `True` → стаття **з'являється**
6. Зняти галку → стаття знову прихована
7. Зайти у Дублікат Cashflow місяць(казна) → той самий патерн

Якщо все працює — implementation готова.

- [ ] **Step 4: Commit acceptance test**

```bash
cd C:\Configuration_downloads\BASERP25
git add _Rarzrabotki/Python/test/test_dds_exclude_e2e.py
git commit -m "test: e2e acceptance for А_ИсключатьИзОтчетаCashflow integration"
```

---

## Final checklist (для self-review)

- [ ] Реквізит у 1С виявлений (Task 1)
- [ ] mapping/baserp_storage.json оновлений з `_Fld<NNN>` (Task 2)
- [ ] `Dim_DDS_Articles.Is_Excluded_From_Cashflow` колонка створена (Task 3)
- [ ] Індекс `IX_DDS_Articles_Excluded` створений (Task 3)
- [ ] DDL файл `02_dim_tables.sql` оновлений (Task 3)
- [ ] Pipeline `dim_dds_articles.json` оновлений (Task 4)
- [ ] ETL запущений, дані в SQL (Task 5)
- [ ] PBI модель має колонку `А_ИсключатьИзОтчетаCashflow` через Power Query Refresh (Task 6)
- [ ] На 4 сторінках Cashflow page-filter + slicer (Task 7)
- [ ] PL.pbix збережений (Task 7)
- [ ] 3 doc-файли оновлені (Task 8)
- [ ] Acceptance test passed (Task 9)
- [ ] Manual PBI test з реальним прикладом виключення (Task 9 Step 3)
