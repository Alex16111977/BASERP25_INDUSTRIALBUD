# OLAP — SQL-схема OlapBASERP (Stage 2, ✅ DONE)

> 24 таблиці на SQL Server `localhost / OlapBASERP`. DDL, індекси, connection deta. Commit `df732a73f` (Stage 2).

---

## Connection Details

| Параметр | Значення |
|---|---|
| **SQL Server version** | Microsoft SQL Server 2022 (RTM-GDR) Standard Edition (16.0.1165.1) |
| **Server name (hostname)** | `SQLSERVER` (Windows hostname) |
| **Connection** | `localhost` (default instance, port 1433) — `(local)` / `127.0.0.1` теж OK |
| **DB** | `OlapBASERP` (recovery SIMPLE) |
| **Auth** | SQL Server authentication (Mixed Mode) |
| **User** | `sa` (sysadmin server role — auto db_owner у всіх БД) |
| **Password** | `Brw739182465!` |
| **ODBC driver** | `ODBC Driver 17 for SQL Server` |
| **TCP/IP** | Listening 1433 на 0.0.0.0 і :: |

### Connection strings

**Python pyodbc:**
```python
CONN_OLAP = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;DATABASE=OlapBASERP;"
    "UID=sa;PWD=Brw739182465!;"
)
```

**sqlcmd:**
```bash
sqlcmd -S localhost -U sa -P "Brw739182465!" -d OlapBASERP
```

**Power BI:**
- Сервер: `localhost` (НЕ `localhost;OlapBASERP` — це окремі поля)
- База даних: `OlapBASERP`
- ⚠️ Вкладка **"База данных"** (НЕ Windows!) — sa це SQL-логін, не Windows
- Имя: `sa` / Пароль: `Brw739182465!`

---

## Загальна структура (24 таблиці)

```
OlapBASERP/
├── Fact (3)
│   ├── Fact_PnL              — рядки PnL з 9 dim FK + Source
│   ├── Fact_Cashflow         — рядки Cashflow з CFS_Section
│   └── Fact_CF_Balance       — opening/closing balance per period+account
│
├── Dim (16)
│   ├── Dim_Organizations
│   ├── Dim_Departments
│   ├── Dim_Directions
│   ├── Dim_Counterparties    — + Code_EDRPOU + Tax_Code
│   ├── Dim_Contracts
│   ├── Dim_Items
│   ├── Dim_ItemGroups
│   ├── Dim_Individuals
│   ├── Dim_ObjektyRaschetov  — Справочник.ОбъектыРасчетов (плоский, Fact_Balance.SettlementObj_ID)
│   ├── Dim_Users
│   ├── Dim_BankAccounts      — + Account_Type ('Bank'/'Cash') + Currency_ID
│   ├── Dim_Currencies
│   ├── Dim_DDS_Articles      — + CFS_Section денормалізація
│   ├── Dim_Expense_Articles
│   ├── Dim_Income_Articles
│   ├── Dim_PL_Articles       — + Sort_Order + Group_ID
│   └── Dim_PL_ArticleGroups
│
├── Bridge (1)
│   └── PLArticle_DDS         — мапа N:M PL-стаття ↔ ДДС
│
└── Util (4)
    ├── Calendar              — 2191 днів 2025-2030
    ├── CFS_Sections          — 4 рядки (Operating/Investing/Financing/Internal)
    ├── Table_Measures        — порожня hub-таблиця для DAX-мір у PBIX
    └── ETL_Runs              — логування ETL запусків
```

---

## Fact_PnL (DDL)

```sql
CREATE TABLE Fact_PnL (
    Fact_ID                 bigint IDENTITY(1,1) PRIMARY KEY,
    Period_Month            date NOT NULL,                     -- '2026-02-01'
    Period                  datetime2 NULL,                    -- точна дата руху
    Source                  varchar(20) NOT NULL,              -- 'PL_Excel'|'ERP_OpEx'|'ERP_CoGS'|...
    Recorder_PL_ID          char(32) NULL,                     -- UUID Документ.А_ФинРез_PL

    -- Вимірювання (7 dim FK + Source у dim — итого 8 з логічної точки зору)
    Organization_ID         char(32) NOT NULL,
    Department_ID           char(32) NULL,
    Direction_ID            char(32) NULL,
    PL_Article_ID           char(32) NULL,
    PL_Group_ID             char(32) NULL,
    DDS_Article_ID          char(32) NULL,
    Counterparty_ID         char(32) NULL,

    -- Реквізити
    Income_Article_ID       char(32) NULL,
    Expense_Article_ID      char(32) NULL,
    Currency_ID             char(32) NULL,
    Exchange_Rate           decimal(15,4) NULL,

    -- Drill-down (для гіперпосилань у Power BI)
    Source_Recorder_ID      char(32) NULL,                     -- UUID первинного документа
    Source_Recorder_Type    varchar(60) NULL,                  -- 'РеализацияТоваровУслуг'
    Source_Recorder_Url     varchar(500) NULL,                 -- 'e1cib/data/Документ.X?ref=...'
    Source_Recorder_Presentation nvarchar(500) NULL,           -- 'Реалізація 0Ц-000123 ...'

    -- Ресурси (5 + 1 валюта)
    Sum_Plan_Grn            decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Plan_F1_Grn         decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Plan_F2_Grn         decimal(15,2) NOT NULL DEFAULT 0,
    Sum_ERP_Grn             decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Kazna_Grn           decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Original            decimal(15,4) NULL,                -- у валюті оригіналу

    Loaded_At               datetime2 NOT NULL DEFAULT SYSDATETIME(),

    INDEX IX_PnL_Period_Source (Period_Month, Source),
    INDEX IX_PnL_Department    (Department_ID, Period_Month),
    INDEX IX_PnL_PL_Article    (PL_Article_ID, Period_Month),
    INDEX IX_PnL_DDS_Article   (DDS_Article_ID, Period_Month)
);
```

**Тип ключа `char(32)`:** UUID-рядок без дефісів (32 hex символи). 1С UUID `f28832a2-b83c-499c-8099-f001efb4aa20` зберігається як `f28832a2b83c499c8099f001efb4aa20`. Конвертація у Python ETL через `utils/uuid_utils.py`.

**Чому char(32) а не uniqueidentifier:** SQL Server `uniqueidentifier` потребує конвертації при JOIN; `char(32)` — простіший і однотипний з 1С.

---

## Fact_Cashflow (DDL)

```sql
CREATE TABLE Fact_Cashflow (
    Fact_ID                 bigint IDENTITY(1,1) PRIMARY KEY,
    Period                  datetime2 NOT NULL,
    Period_Month            date NOT NULL,
    Source                  varchar(20) NOT NULL,              -- 'ERP_Безнал'|'ERP_Нал'|'Казна'
    CFS_Section             varchar(15) NULL,                  -- 'Operating'|...
    Recorder_DDS_ID         char(32) NULL,                     -- UUID Документ.А_ФинРез_DDS

    -- Вимірювання (8)
    Organization_ID         char(32) NOT NULL,
    Department_ID           char(32) NULL,
    Direction_ID            char(32) NULL,
    BankAccount_ID          char(32) NOT NULL,
    Account_Type            varchar(10) NOT NULL,              -- 'Bank'|'Cash'
    DDS_Article_ID          char(32) NULL,
    Counterparty_ID         char(32) NULL,
    Contract_ID             char(32) NULL,

    -- Реквізити
    Direction               varchar(10) NOT NULL,              -- 'Inflow'|'Outflow'
    Currency_ID             char(32) NULL,
    Exchange_Rate           decimal(15,4) NULL,

    -- Drill-down
    Source_Recorder_ID      char(32) NULL,
    Source_Recorder_Type    varchar(60) NULL,
    Source_Recorder_Url     varchar(500) NULL,
    Source_Recorder_Presentation nvarchar(500) NULL,

    -- Ресурси
    Sum_Grn                 decimal(15,2) NOT NULL,            -- абсолютна сума
    Sum_Original            decimal(15,4) NULL,

    Loaded_At               datetime2 NOT NULL DEFAULT SYSDATETIME(),

    INDEX IX_CF_Period_Source (Period_Month, Source),
    INDEX IX_CF_Section       (CFS_Section, Period_Month),
    INDEX IX_CF_Account       (BankAccount_ID, Period_Month),
    INDEX IX_CF_Article       (DDS_Article_ID, Period_Month)
);
```

**Direction:** `'Inflow'` | `'Outflow'` — обчислюється Python ETL з `Vnyzhe.ВидДвижения` (Поступление → Inflow, Списание → Outflow). Дублює інформацію з ВидДвижения (denormalized для DAX simplicity).

---

## Fact_CF_Balance (DDL)

```sql
CREATE TABLE Fact_CF_Balance (
    Balance_ID         bigint IDENTITY(1,1) PRIMARY KEY,
    Period_Month       date NOT NULL,
    Organization_ID    char(32) NOT NULL,
    BankAccount_ID     char(32) NOT NULL,
    Account_Type       varchar(10) NOT NULL,
    Currency_ID        char(32) NULL,
    Sum_Grn_Open       decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Grn_Close      decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Grn_Inflow     decimal(15,2) NOT NULL DEFAULT 0,
    Sum_Grn_Outflow    decimal(15,2) NOT NULL DEFAULT 0,
    Loaded_At          datetime2 NOT NULL DEFAULT SYSDATETIME(),
    INDEX IX_CFB_Period (Period_Month, BankAccount_ID)
);
```

**Призначення:** opening/closing балансу на кожен місяць по кожному рахунку. Завантажується через `.ОстаткиИОбороты` віртуальну таблицю регістрів `ДенежныеСредстваБезналичные` і `ДенежныеСредстваНаличные` у Python ETL `41_cashflow_balance.py`.

**DAX use:** `[Reconciliation] = [Balance Close] - [Balance Open] - [CFS Total]` — має бути ≈ 0 для перевірки консистентності.

---

## 17 Dim таблиць

> +1 з 2026-05-17: **Dim_ObjektyRaschetov** (`Справочник.ОбъектыРасчетов`,
> `_Reference319` — фізично лише `_IDRRef/_Description/_Marked` → плоский, без
> коду/ієрархії; sql_backend full_reload; FK `Fact_Balance.SettlementObj_ID`,
> джерело `Свод_РасчетыСПартнерами`). DDL:
> `scripts/ddl_dim_objekty_raschetov.sql` (idempotent DROP+CREATE). Колонки:
> `SettlementObj_ID char(32) PK, SettlementObj_Name nvarchar(150),
> Marked_For_Deletion bit, Loaded_At datetime2`.

### Common pattern (для більшості Dim)

```sql
CREATE TABLE Dim_<Name> (
    <Name>_ID         char(32) PRIMARY KEY,
    <Name>_Code       varchar(50) NULL,
    <Name>_Name       nvarchar(150) NOT NULL,
    Parent_ID         char(32) NULL,                            -- для ієрархії
    Is_Group          bit NOT NULL DEFAULT 0,
    Marked_For_Deletion bit NOT NULL DEFAULT 0,
    Loaded_At         datetime2 NOT NULL DEFAULT SYSDATETIME()
);
```

### Variants

#### Dim_DDS_Articles
```sql
+ CFS_Section varchar(15) NULL  -- денормалізовано з реквізиту А_РазделCFS
```

#### Dim_BankAccounts
```sql
+ Account_Type varchar(10) NOT NULL  -- 'Bank' (з БанковскиеСчетаОрганизаций) або 'Cash' (з Кассы)
+ Currency_ID  char(32) NULL          -- FK на Dim_Currencies
```

UNION 2 справочників 1С: `БанковскиеСчетаОрганизаций` + `Кассы`.

#### Dim_Counterparties
```sql
+ Code_EDRPOU varchar(20) NULL  -- код ЄДРПОУ (для юр. осіб)
+ Tax_Code    varchar(20) NULL  -- ІНН/податковий код
```

#### Dim_PL_Articles
```sql
+ Sort_Order  int          NULL  -- з реквізиту Сорт
+ Group_ID    char(32)     NULL  -- FK на Dim_PL_ArticleGroups
+ Type_Statya nvarchar(50) NULL  -- 2026-05-21: ТипСтатьи (Доход/Расход/ОперационныйИтог/Информационный),
                                 --              для mirror-знака у Fact_PnL та фільтрів PBIX
```

ETL: `pipelines/dim_catalogs.json` step `dim_pl_articles` → `enum_resolver` мапить
`Перечисление.А_ТипСтатьиPL` через `FROZEN_ENUMS` (`enum_resolver.py`).

### Список 16 Dim

| Dim | 1С Source | Special |
|---|---|---|
| Dim_Organizations | `Справочник.Организации` | — |
| Dim_Departments | `Справочник.СтруктураПредприятия` | + Parent_ID для ієрархії |
| Dim_Directions | `Справочник.НаправленияДеятельности` | — |
| Dim_Counterparties | `Справочник.Контрагенты` | + Code_EDRPOU, Tax_Code |
| Dim_Contracts | `Справочник.ДоговорыКонтрагентов` | — |
| Dim_Items | `Справочник.Номенклатура` | — |
| Dim_ItemGroups | `Справочник.Номенклатура` (filter ЭтоГруппа=Истина) | — |
| Dim_Individuals | `Справочник.ФизическиеЛица` | — |
| Dim_ObjektyRaschetov | `Справочник.ОбъектыРасчетов` (`_Reference319`, плоский) | Fact_Balance.SettlementObj_ID |
| Dim_Users | `Справочник.Пользователи` | — |
| Dim_BankAccounts | UNION (`БанковскиеСчетаОрганизаций` + `Кассы`) | + Account_Type, Currency_ID |
| Dim_Currencies | `Справочник.Валюты` | — |
| Dim_DDS_Articles | `Справочник.СтатьиДвиженияДенежныхСредств` | + CFS_Section |
| Dim_Expense_Articles | `ПланВидовХарактеристик.СтатьиРасходов` | — |
| Dim_Income_Articles | `ПланВидовХарактеристик.СтатьиДоходов` | — |
| Dim_PL_Articles | `Справочник.А_Статьи_PL` | + Sort_Order, Group_ID |
| Dim_PL_ArticleGroups | `Справочник.А_ГруппаСтатей_PL` | — |

---

## Bridge: PLArticle_DDS

```sql
CREATE TABLE PLArticle_DDS (
    PL_Article_ID    char(32) NOT NULL,
    DDS_Article_ID   char(32) NOT NULL,
    PRIMARY KEY (PL_Article_ID, DDS_Article_ID)
);
```

**Призначення:** мапа N:M між PL-статтями і ДДС-статтями. Завантажується з `Справочник.А_Статьи_PL.Статьи` (ТЧ).

**Унікальність:** одна ДДС в одній PL (1:N), але PRIMARY KEY на пару дозволяє bridge для DAX.

---

## Util-таблиці (4)

### Calendar (2191 рядків)

```sql
CREATE TABLE Calendar (
    Date_Key       date PRIMARY KEY,
    Year           int NOT NULL,
    Quarter        int NOT NULL,
    Month_Num      int NOT NULL,
    Month_Name     nvarchar(20) NOT NULL,         -- 'Февраль' / 'Лютий' (зачежить від locale)
    Month_Start    date NOT NULL,
    Month_End      date NOT NULL,
    Day_Of_Month   int NOT NULL,
    Day_Of_Week    int NOT NULL,
    Day_Name       nvarchar(20) NOT NULL,
    Is_Weekend     bit NOT NULL,
    Year_Month     varchar(7) NOT NULL,           -- '2026-02'
    INDEX IX_Calendar_Month (Year_Month)
);
```

**Заповнено:** 2025-01-01 → 2030-12-31 = **2191 днів**. Згенеровано CTE з `MAXRECURSION 3000` у `03_bridge_util.sql`.

**DAX:** використовується як time-intelligence таблиця у Power BI — основа для всіх часових слайсерів і MTD/YTD calculations.

### CFS_Sections (4 рядки)

```sql
CREATE TABLE CFS_Sections (
    Section_Code    varchar(15) PRIMARY KEY,      -- 'Operating'|'Investing'|'Financing'|'Internal'
    Section_Name    nvarchar(50) NOT NULL,
    Sort_Order      int NOT NULL,
    Is_In_CFS_Total bit NOT NULL                  -- Internal=0, інші=1
);
```

**Seed data** (через Python+pyodbc — sqlcmd обрізає Cyrillic у `N'...'` literals):
```python
INSERT INTO CFS_Sections VALUES
    ('Operating', 'Операционная деятельность', 1, 1),
    ('Investing', 'Инвестиционная деятельность', 2, 1),
    ('Financing', 'Финансовая деятельность', 3, 1),
    ('Internal',  'Внутренние перемещения',     4, 0);
```

**DAX use:** JOIN до Fact_Cashflow для отримання Section_Name, фільтр по Is_In_CFS_Total для CFS_Total.

### Table_Measures (порожня)

```sql
CREATE TABLE Table_Measures (
    Measure_ID      int IDENTITY(1,1) PRIMARY KEY,
    Placeholder     varchar(10) NULL
);
```

**Призначення:** порожня hub-таблиця у PBIX для організації DAX-мір. Power BI best-practice: створити окрему таблицю для всіх measures щоб вони не плавали між Fact-таблицями. У SQL вона порожня; у PBIX до неї додаються 70+ DAX measures для PnL.

### ETL_Runs

```sql
CREATE TABLE ETL_Runs (
    Run_ID         bigint IDENTITY(1,1) PRIMARY KEY,
    Script         varchar(100) NOT NULL,                      -- '30_fact_pnl' / '10_dim_organizations' / ...
    Period_Month   date NULL,
    Started_At     datetime2 NOT NULL DEFAULT SYSDATETIME(),
    Finished_At    datetime2 NULL,
    Rows_Loaded    int NULL,
    Status         varchar(20) NOT NULL DEFAULT 'Running',     -- 'Running'|'OK'|'Failed'
    Error          nvarchar(2000) NULL,
    INDEX IX_ETLRuns_Script_Period (Script, Period_Month)
);
```

**Призначення:** логування Python ETL запусків. Кожен скрипт-loader робить `INSERT INTO ETL_Runs` на старті (Status='Running'), `UPDATE` на завершенні (Status='OK' або 'Failed' з error text).

**Use case:** оркестратор перевіряє чи попередня spuštění завершилась OK перед наступною; адмін бачить історію запусків і час виконання.

---

## DDL файли

| Файл | Призначення |
|---|---|
| `_Rarzrabotki/Python/Olap/ddl/00_create_database.sql` | CREATE DATABASE OlapBASERP + RECOVERY SIMPLE |
| `_Rarzrabotki/Python/Olap/ddl/01_fact_tables.sql` | 3 Fact таблиці з індексами |
| `_Rarzrabotki/Python/Olap/ddl/02_dim_tables.sql` | 16 Dim таблиць |
| `_Rarzrabotki/Python/Olap/ddl/03_bridge_util.sql` | Bridge + Calendar + Table_Measures + ETL_Runs |
| `_Rarzrabotki/Python/Olap/ddl/seed_cfs_sections.py` | Python-seed для CFS_Sections (Cyrillic encoding fix) |
| `_Rarzrabotki/Python/Olap/ddl/99_verify.sql` | Verification: SELECT COUNT(*) tables = 24 |
| `_Rarzrabotki/Python/Olap/tests/test_olap_connectivity.py` | pyodbc smoke test |

---

## Common DDL patterns

### IF NOT EXISTS

Кожен CREATE TABLE обгорнуто у `IF OBJECT_ID('dbo.X', 'U') IS NULL` для ідемпотентного re-run скрипта. CREATE DATABASE — у `IF DB_ID('OlapBASERP') IS NULL`.

### Cyrillic encoding ⚠️

**Проблема:** sqlcmd Windows codepage не UTF-8, тому `INSERT INTO X VALUES (N'Кирилиця', ...)` обрізає текст. Workaround:
- DDL без Cyrillic у SQL-файлах (тільки ASCII в кодових іменах)
- Cyrillic seed через **Python + pyodbc parameterized query** — драйвер сам передає текст у Unicode

**Альтернативно:** запускати sqlcmd з `-f 65001` для UTF-8 input або зберегти SQL-файли як UTF-8 with BOM. Не використовуємо — Python надійніше.

---

## Повний список таблиць (verify)

```sql
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE='BASE TABLE'
ORDER BY TABLE_NAME;
```

Очікуваний результат (24 рядки, alphabetical):
1. Calendar
2. CFS_Sections
3. Dim_BankAccounts
4. Dim_Contracts
5. Dim_Counterparties
6. Dim_Currencies
7. Dim_DDS_Articles
8. Dim_Departments
9. Dim_Directions
10. Dim_Expense_Articles
11. Dim_Income_Articles
12. Dim_Individuals
13. Dim_Items
14. Dim_ItemGroups
15. Dim_Organizations
16. Dim_PL_Articles
17. Dim_PL_ArticleGroups
18. Dim_Users
19. ETL_Runs
20. Fact_Cashflow
21. Fact_CF_Balance
22. Fact_PnL
23. PLArticle_DDS
24. Table_Measures

---

## Cross-references

- Python ETL який вантажить ці таблиці: [olap_etl_pipeline.md](olap_etl_pipeline.md)
- Power BI який читає ці таблиці: [olap_powerbi_model.md](olap_powerbi_model.md)
- Який 1С контент потрапляє у які SQL колонки: [olap_data_sources_erp.md](olap_data_sources_erp.md)
- Spec v3 §5: `docs/superpowers/specs/2026-05-01-olap-baserp-architecture-design-v3-final.md`

---

## Balance Stage (2026-05-16) — Fact_Balance + Dim_PAP_Articles

Канон §10/Roadmap (OD-9). Джерело: `РегистрСведений.А_ОтчетБаланс_Свод`
(Етапи 1-4 свёртки в 1С, ETL лише копіює). Таблиць OlapBASERP: **24 → 26**.

DDL: `_Rarzrabotki/Python/Olap/ddl/07_balance.sql` (applier `apply_07_balance.py`,
ідемпотентно `IF OBJECT_ID IS NULL`).

**`Fact_Balance`** (**27 кол.**, 2026-05-18 +TaxType): `Balance_ID bigint IDENTITY PK`, `Period_Month date`,
`Period datetime2`, `Source varchar(40)` (enum **ИсточникиУправленческогоБаланса**, 31 знач., _EnumOrder 0..30 — типове перечислення ERP; виправлено 2026-05-17, НЕ кастомне А_ИсточникБаланса; **`ПустаяСсылка`** для прямих рухів ПАП — `Свод_ПрочиеАктивыПассивы_Прямой`),
`Recorder_Balance_ID char(32)` (Док.А_ФинРез_Баланс); 13 dim `char(32)`
(Organization/Department/PAP_Article/Item/Counterparty/Partner/Warehouse/
OperObject/Contract/Individual/Cash/SettlementObj/Intangible);
**`TaxType varchar(50) COLLATE Cyrillic_General_CI_AS NULL`** (enum
`Перечисление.ТипыНалогов` 14 знач., метаімена; заповнено лише у статті
«Налоги», решта `"ПустаяСсылка"`; ALTER ADD 2026-05-18 для
`Свод_ПрочиеАктивыПассивы_Прямой` розшифровки ПАП.Аналитика); `Analytics1..3
nvarchar(150)`; 4 ресурси `decimal(15,2)` `Sum_Open/Inflow/Outflow/Close`;
`Loaded_At`; 4 індекси (Period_Source, Article, Individual, SettlementObj).

**`Dim_PAP_Articles`** (ПВХ.СтатьиАктивовПассивов): `PAP_Article_ID char(32) PK`,
`PAP_Article_Code`, `PAP_Article_Name nvarchar(150)`, `Parent_ID`, `Is_Group bit`,
`AktivPassiv varchar(15)` (Aktiv/Passiv/AktivPassiv — реквізит АктивПассив, OD-9),
`Marked_For_Deletion bit`, `Loaded_At`.

**`Dim_TaxTypes`** (2026-05-18, `Перечисление.ТипыНалогов` — образець
Dim_ObjektyRaschetov): `TaxType varchar(50) COLLATE Cyrillic_General_CI_AS
PK` (метаім'я == Fact_Balance.TaxType), `TaxType_Name nvarchar(100)`
(синонім 1С UI), `EnumOrder int`, `Loaded_At`. **15 рядків** (14 enum +
`ПустаяСсылка`/«(Не налог)» для НЕ-«Налоги»). DDL
`scripts/ddl_dim_tax_types.sql`; сидер `scripts/seed_dim_tax_types.py`
(імена/синоніми з 1С COM — `_Enum1651` у SQL backend не має імен, frozen).
FK `Fact_Balance.TaxType→Dim_TaxTypes.TaxType` 100% (verify 0 orphans).

Row counts (live, січень 2026/ТОВ): `Fact_Balance` 7 643 (2026-01),
`Dim_PAP_Articles` 54, `Dim_TaxTypes` 15.

### Оновлення 2026-05-17 — `Свод_ДенежныеСредства` LIVE + `Dim_Warehouses` (26 → 27 таблиць)

**Fact_Balance** тепер містить 5 Source (Себест + 4 ден.): окрім
`СебестоимостьТоваров` — `ДенежныеСредстваБезналичные`/`Наличные`/`ВПути`/
`УПодотчетныхЛиц` (Свод_ДенежныеСредства в 1С). Колонки незмінні; заповнюються
`Cash_ID` (безнал→БанкСчёт, нал→Касса), `Individual_ID` (підзвіт→ПодотчетноеЛицо),
`Warehouse_ID`/`Item_ID` (Себест). Січень/ТОВ: Σ ден. групи Sum_Close=
**75 265 344,95**; «безнал» КО=50 435 887,99. ETL `fact_balance.json` без змін
(raw_sql фільтрує лише Орг+період, Source НЕ фільтрує — нові рядки тягне сам).

**`Dim_Warehouses`** (нова, 27-а таблиця; для `Fact_Balance.Warehouse_ID`):
```sql
CREATE TABLE Dim_Warehouses (
    Warehouse_ID         char(32) PRIMARY KEY,   -- Справочник.Склады
    Warehouse_Name       nvarchar(250) NOT NULL,
    Parent_ID            char(32) NULL,          -- ієрархія (Иерархия груп і елементів)
    Is_Group             bit NOT NULL DEFAULT 0,
    Marked_For_Deletion  bit NOT NULL DEFAULT 0,
    Hierarchy_Path       nvarchar(500) NULL,
    Hierarchy_Depth      int NULL,
    Level1..Level5       nvarchar(150) NULL,      -- для PBI ієрархії ИерархияСкладов
    Loaded_At            datetime2 NOT NULL DEFAULT SYSDATETIME()
);
```
DDL `ddl/08_dim_warehouses.sql` (applier `apply_08_dim_warehouses.py`).
⚠️ `Справочник.Склады` **БЕЗ Кода** (Длина кода=0 → `_Reference502` не має
`_Code`, але має `_ParentIDRRef`/`_Folder` — ієрархічний). ETL крок
`dim_warehouses` у `dim_catalogs.json` = recursive-CTE по `_Reference502`
БЕЗ `_Code` (паттерн dim_organizations). Live: **347 рядків, 303 групи,
Level1..5 заповнені**. `refresh_mapping.py` WHITELIST += `Справочник.Склады`
(→ `_Reference502` у baserp_storage.json, 84 об'єкти).

> **Урок:** відсутність `_Code` ≠ неієрархічний довідник (catalog може мати
> Длина кода=0). Ієрархічність — по Конфігуратору / фізичних
> `_ParentIDRRef`+`_Folder`, перевірка `scripts/probe_ref502_columns.py`.

Повний список тепер **27 рядків** (24 базові + Fact_Balance + Dim_PAP_Articles
+ Dim_Warehouses).

### Оновлення 2026-05-19 — розширення Dim_Contracts/Dim_ObjektyRaschetov + нові Dim_TipyDogovorov/Dim_FinAgents (29 таблиць)

**`Dim_Contracts`** пересоздан (idempotent DROP+CREATE; DDL `scripts/ddl_dim_contracts.sql`).
Нові колонки (додано до всіх старих):

| Колонка | Тип | Опис |
|---|---|---|
| `Is_FinAgent_Contract` | `bit NULL` | Признак договора фінагента (NULL-able — інакше транзит старого sql_backend ламався) |
| `TipDogovora` | `varchar(50)` | Метаімя `Перечисление.ТипыДоговоров` (FK→Dim_TipyDogovorov) |
| `FinAgent_ID` | `char(32)` | FK→Dim_FinAgents (Справочник.А_ФинАгенты) |
| `Department_Name` | `nvarchar(300)` | Денорм. назва підрозділу (бух.) |
| `Dept_OkazUslug_Name` | `nvarchar(300)` | Підрозділ-надавач послуг між підрозділами |
| `Partner_Name` | `nvarchar(300)` | Партнер договора |
| `Counterparty_Name` | `nvarchar(300)` | Контрагент договора |
| `DDS_Article_Forced_Name` | `nvarchar(300)` | Стаття ДДС (основна примусово) |
| `Org_Buh_Name` | `nvarchar(300)` | Організація (бух.) |

Джерело: `_Reference171` (ДоговорыКонтрагентов); денорм через LEFT JOIN `_Reference540/_Reference360/_Reference263/_Reference529/_Reference329`.

**`Dim_ObjektyRaschetov`** пересоздан (DDL `scripts/ddl_dim_objekty_raschetov.sql`). Нові колонки:
`TipRaschetov varchar(50)` / `TipObjektaRaschetov varchar(50)` (enum метаімена),
`Partner_Name` / `Department_Name` (nvarchar(300)),
`Counterparty_ID` / `Contract_ID` / `Object_ID` (char(32), composite UUID з `_Fld...RRef`),
`Object_Type_Name nvarchar(300)` (назва типу посилання об'єкта через `ТипСсылки`→`_Reference211`).
Джерело: `_Reference319`.

**`Dim_TipyDogovorov`** (нова, DDL `scripts/ddl_dim_tipy_dogovorov.sql`):
`TipDogovora varchar(50) COLLATE Cyrillic_General_CI_AS PK` (метаімя),
`TipDogovora_Name nvarchar(100)`, `EnumOrder int`, `Loaded_At`.
**12 рядків** (11 знач. `Перечисление.ТипыДоговоров` + ПустаяСсылка).
Сид: `scripts/seed_dim_tipy_dogovorov.py` (COM, паттерн Dim_TaxTypes). **Вне DEFAULT pipelines** — запускати окремо.

**`Dim_FinAgents`** (нова, DDL `scripts/ddl_dim_fin_agents.sql`):
`FinAgent_ID char(32) PK`, `FinAgent_Name nvarchar(300)`, `Loaded_At`.
**13 рядків** (`_Reference54722` + unknown-member ПустаяСсылка). ETL-шаг `dim_fin_agents`
у `pipelines/dim_catalogs.json` (`raw_sql`-крок; входить у default dim_catalogs).

FK: `Dim_Contracts.TipDogovora→Dim_TipyDogovorov.TipDogovora` і
`Dim_Contracts.FinAgent_ID→Dim_FinAgents.FinAgent_ID` (verify 0 orphans).

**WHITELIST** `mapping/refresh_mapping.py` +5: `Справочник.А_ФинАгенты`,
`Справочник.ИдентификаторыОбъектовМетаданных`, `Перечисление.ТипыДоговоров`,
`Перечисление.ТипыРасчетовСПартнерами`, `Перечисление.ТипыОбъектовРасчетов`;
`baserp_storage.json` перегенерований.

Row counts (verify PASS): Dim_Contracts=8248, Dim_ObjektyRaschetov=14109, Dim_TipyDogovorov=12, Dim_FinAgents=13.
Повний список тепер **29 таблиць** (+ Dim_TipyDogovorov + Dim_FinAgents).

---

## 2026-06-12 — Dim_VidyKontragentov + 3 колонки Dim_Contracts (виды контрагентов для баланса)

Нова таблиця **`Dim_VidyKontragentov`** (`scripts/ddl_dim_vidy_kontragentov.sql`, idempotent):

| колонка | тип | джерело |
|---|---|---|
| `VidKontragenta_ID` | char(32) PK | `_Reference56330._IDRRef` (Справочник.А_ВидыКонтрагентовДляБаланса) |
| `VidKontragenta_Name` | nvarchar(150) | `_Description` |
| `Code` | varchar(20) | `_Code` |
| `Marked_For_Deletion` | bit | `_Marked` |
| `Loaded_At` | datetime2 DEFAULT | — |

Рядків **6** = 5 предопределённых (Внутригрупповые / Внутренние подразделения /
Собственники / Внешние / Кредиторы) + синтетичний «(Пусто)» (0x…01).

**`Dim_Contracts` +3 колонки** (guarded ALTER):
- `VidKontragenta_ID` char(32) — FK → Dim_VidyKontragentov (`_Fld56332RRef` А_ВидКонтрагента)
- `NapravlenieUslug_ID` char(32) — FK → Dim_Directions (`_Fld56331RRef` А_НаправлениеОказаниеУслуг); у PBIX НЕ используется (см. нижче)
- `NapravlenieUslug_Name` nvarchar(150) — денорм-имя направления (JOIN `_Reference292`), для PBIX-колонки «Направление (услуги между подр.)» — рішення фінансиста: НЕ окреме вимірювання

Розкладка договорів по видах (== 1С до штуки, verify PASS): Внутригрупповые 331,
Внутренние підрозділи 1106, Собственники 10, Внешние 7162, Кредиторы 17;
NapravlenieUslug_* — 1104. FK orphans 0. Повний список тепер **30 таблиць**.
Verify: `scripts/verify_olap_dim_vid_kontragenta.py` (PASS, динамічні звірки з 1С COM).
