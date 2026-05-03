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

## 16 Dim таблиць

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
+ Sort_Order int NULL          -- з реквізиту Сорт
+ Group_ID   char(32) NULL     -- FK на Dim_PL_ArticleGroups
```

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
