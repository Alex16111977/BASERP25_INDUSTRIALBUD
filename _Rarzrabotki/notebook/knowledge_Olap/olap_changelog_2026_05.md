# OLAP Changelog — травень 2026

> Хроніка змін у BI-конвеєрі BASERP25 (Ai_Olap + OlapBASERP SQL + PL.pbix).

---

## §11 — `Сорт` (Sort_Order) для Dim_PL_Articles / Dim_PL_ArticleGroups + PBIX SortByColumn

**2026-05-21 (продолження §10)**

### Що зробили

Реквізит `Сорт` (Число) з обох довідників 1С тепер прокидається у OLAP і використовується для нативного впорядкування у візуалах PBIX.

1. **ETL `pipelines/dim_catalogs.json`**:
   - step `dim_pl_articles`: `fields += ["Сорт"]`, `column_map += {"Сорт": "Sort_Order"}`
   - step `dim_pl_article_groups`: те саме
   - Прогон: `python main.py --run-once dim_catalogs` → 68 PL-статей + 8 груп.

2. **PBIX live-правки через `powerbi-modeling-mcp`**:
   - `partition_operations Update` для `Dim_PL_ArticleGroups` → перехід M-вираз на native query
     (`Sql.Database("SQLSERVER", "OlapBASERP", [Query="SELECT ... Sort_Order FROM dbo.Dim_PL_ArticleGroups"])`)
     щоб обійти PBI schema cache.
   - `partition_operations Refresh` обох таблиць (`А_Статьи_PL`, `А_ГруппаСтатей_PL`) — `useTransaction: false`.
   - `column_operations Update sortByColumn: "Sort_Order"` для `ГруппаСтатейPL` і `СтатьяPL`.
   - Acceptance через `INFO.COLUMNS()`: `SortByColumnID` ≠ NULL для обох → колонки прив'язані до Sort_Order.

### Acceptance: PnL за групами в правильному порядку

```dax
EVALUATE SUMMARIZECOLUMNS(
    'А_ГруппаСтатей_PL'[ГруппаСтатейPL],
    'А_ГруппаСтатей_PL'[Sort_Order],
    "PL", SUM(Fact_PnL[Sum_Fact])
) ORDER BY [Sort_Order]
```

| Sort | Група | PnL |
|---:|---|---:|
| 100 | Операционный доход | +360 619 237,36 |
| 200 | Себестоимость проданой продукции | −325 019 351,64 |
| 300 | Дополнительные расходы | −31 545,93 |
| 400 | Общепроизводственные затраты | −2 469 744,07 |
| 500 | Маркетинговые затраты | −303 561,91 |
| 600 | Административные затраты | −39 466 132,11 |
| 700 | Налоги и сборы | 0 |
| 800 | Финансовая деятельность | +1 148 189,47 |
| | **P&L** | **−5 522 909 ₽** |

Тепер у візуалах PL.pbix `ГруппаСтатейPL` сортується **в порядку Sort_Order** (не алфавітно). Аналогічно для `СтатьяPL` — фінансист може ставити пріоритет статей всередині групи через реквізит `Сорт` у 1С.

### Документація
- Knowledge_Olap: `olap_sql_schema.md` — `Dim_PL_Articles.Sort_Order` і `Dim_PL_ArticleGroups.Sort_Order` тепер заповнюються (раніше були NULL).
- Knowledge_PL: `pl_finrez_sign_by_type.md` — ETL extension згадує також Sort_Order.

---

## §10 — Mirror-знак суми за `ТипСтатьи` у Fact_PnL (cross-layer)

**2026-05-21**

### Що зробили

Сквозна реалізація mirror-семантики знака суми за реквізитом `Справочник.А_Статьи_PL.ТипСтатьи` (Перечисление.А_ТипСтатьиPL: Доход/Расход/ОперационныйИтог/Информационный):

1. **BSL** `Documents/А_ФинРез_PL/Ext/ObjectModule.bsl`:
   У фінальному SELECT `втРезультат` 6 ресурсів обернуто у CASE: `Расход → -Сумма`, `ИНАЧЕ → Сумма`. Σ |abs| зберігається до 0.01 ₽.

2. **Перепровод 28 проведених А_ФинРез_PL** за 2024-01..2026-04 через `_Rarzrabotki/Python/scripts/_reprovesti_finrez_pl.py` (28/28 OK).

3. **SQL DDL**: `ALTER TABLE Dim_PL_Articles ADD Type_Statya nvarchar(50) NULL` в OlapBASERP.

4. **ETL extension**:
   - `mapping/refresh_mapping.py`: whitelist += `Перечисление.А_ТипСтатьиPL`
   - `ai_olap/transformers/enum_resolver.py`: `FROZEN_ENUMS["Перечисление.А_ТипСтатьиPL"] = ["Доход", "Расход", "ОперационныйИтог", "Информационный"]`
   - `pipelines/dim_catalogs.json` step `dim_pl_articles`: + fields["ТипСтатьи"], + enum_resolver.column_to_enum, + column_map ТипСтатьи→Type_Statya
   - Прогон: `python main.py --run-once dim_catalogs` (68 PL-статей) + `--run-once fact_pnl` (55 146 рядків).

5. **PL.pbix**: M-вираз партиції `Dim_PL_Articles` переведено з table-mode на native query `Sql.Database("SQLSERVER", "OlapBASERP", [Query="SELECT ... Type_Statya FROM dbo.Dim_PL_Articles"])` — обхід PBI schema cache (memory `pbix_sql_database_schema_cache.md`).

### Acceptance (2026-05-21)

| Шар | Доход Σsigned | Расход Σsigned | P&L |
|---|---:|---:|---:|
| 1С `РегистрСведений.А_ОтчетPL_Свод` | +361 767 426,83 | −367 290 335,66 | −5 522 908,83 ₽ |
| OLAP `Fact_PnL.Sum_Fact` | +361 767 426,83 | −367 290 335,66 | −5 522 908,83 ₽ |

**Збіг 1С == OLAP до 0,01 ₽.**

### DAX-наслідок

```dax
-- БУЛО:
Маржинальный доход = 
    CALCULATE(SUM(Fact_PnL[Sum_Fact]), Dim_PL_Articles[Type_Statya] = "Доход") -
    CALCULATE(SUM(Fact_PnL[Sum_Fact]), Dim_PL_Articles[Type_Statya] = "Расход")
-- СТАЛО:
Маржинальный доход = SUM(Fact_PnL[Sum_Fact])
```

### Документація
- Knowledge_PL: новий файл `pl_finrez_sign_by_type.md` (повний опис патерну і acceptance).
- Knowledge_PL: оновлено `KNOWLEDGE_MAP_PL.md`, `pl_report_architecture_analyst.md`.
- Knowledge_Olap: оновлено `olap_sql_schema.md` (Dim_PL_Articles + Type_Statya), `olap_powerbi_pl_pbix.md` (А_Статьи_PL 10→11 колонок).

---

## §1 — Pipeline `dim_catalogs.json` — перехід на `raw_sql` + recursive CTE

**Stage v3.7, 2026-05-11**

### Що було

Усі 17 кроків `dim_catalogs.json` використовували `sql_backend` extractor — простий `SELECT ... FROM _Reference{N}` через `mapping/baserp_storage.json`. Це означало:

- Ієрархія `Родитель → Дочірні` зберігалась тільки як `Parent_ID` FK; розгортання у Level1..5 / Hierarchy_Path робилося б на DAX-side (повільно, нестабільно).
- 7 з 8 ключових Dim не мали unknown-member `0x...0001 "(Пусто)"` для NULL FK з Fact-таблиць → broken refs у моделі.
- `Dim_Partners` як SQL-таблиця існувала (4918 рядків + 1 unknown) але pipeline-step для неї **не було** — оновлення не відбувалось.
- `dim_dds_articles` step тягнув `А_РазделCFS` але не `А_ИсключатьИзОтчетаCashflow` (поле, додане Stage v3.6 у DDL `_Fld56081`, лишалось без ETL-маршруту).
- `dim_departments` не мав `Direction_ID` мапінгу; колонка частково заповнювалась з minim попередніх ран.

### Що зробили

Усі 7 ключових кроків + 1 НОВИЙ переписали на `"type": "raw_sql"` з SQL Server recursive CTE, що сам обчислює:

- `Hierarchy_Path` (nvarchar(500), сегменти `_IDRRef` через `|`)
- `Hierarchy_Depth` (int, 1 → корінь, 2..5 — діти)
- `Level1..Level5` (nvarchar(150), назва предка на N-му рівні; NULL якщо глибина менша)
- Для Departments: `Direction_ID` через `COALESCE(NULLIF(self.Direction_ID, 0x..00), r.Direction_ID, 0x..0001)` — самостійно або від найближчого предка, інакше unknown.
- Unknown-member row `0x00000000000000000000000000000001 "(Пусто)"` через `UNION ALL` — додається у кожен з 8 ключових Dim.

| Step | SQL Table | Depth | Особливості |
|---|---|---|---|
| `dim_dds_articles` | `_Reference529` | 5 | + `_Fld55969RRef → CFS_Section` (enum_resolver) + `COALESCE(_Fld56081, 0x00) → Is_Excluded_From_Cashflow` |
| `dim_departments` | `_Reference540` | 3 | + `_Fld54614RRef` Direction_ID COALESCE cascade; `has_folder=False` (немає `_Folder`) |
| `dim_items` | `_Reference306` | 5 | 41 137 рядків (41 136 + 1 unknown); `OPTION (MAXRECURSION 0)` |
| `dim_expense_articles` | `_Chrc1772` | 4 | План видів характеристик; CTE дає 340 рядків (виключено 5 orphan з некоректним Parent ref) |
| `dim_partners` (**НОВИЙ**) | `_Reference360` | 5 | `has_folder=False`; вставлено між `dim_counterparties` (#4) і `dim_contracts` (#6) |
| `dim_counterparties` | `_Reference263` | flat | + `_Fld31169RRef → Partner_ID`, `_Fld31175 → Code_EDRPOU`, `_Fld31173 → Tax_Code` |
| `dim_directions` | `_Reference292` | 2 | hierarchy + unknown |
| `dim_income_articles` | `_Chrc1771` | flat | flat + unknown (немає ієрархії у каталозі) |

**Загалом:** 17 → 18 steps. `raw_sql` контракт у `ai_olap/extractors/sql_backend.py:41-46` приймає `cfg.get("raw_sql")` або `cfg.get("sql")`. Поточний код factory (`extractors/factory.py:12`) направляє і `sql_backend` і `raw_sql` на той самий `SqlBackendExtractor` клас.

### Helper

Виконано через `scripts/update_pipeline_hierarchies.py` (290+ рядків). Дві базові SQL-функції:

- `hierarchy_cte_sql(table, code_col='_Code', has_folder=True, extra_select_root='', extra_select_recurse='', extra_select_unknown='', extra_columns_in_alias='')` — параметризований CTE з 3-х частин (root anchor + recursive child + unknown UNION).
- `flat_with_unknown_sql(table, code_col='_Code', has_parent=True, has_folder=True, extra_select_real='', extra_select_unknown='')` — для не-ієрархічних з UNION unknown.

`main()` патчить pipeline JSON через `replace_step()` + `insert_step_after()`. Запуск: `python scripts/update_pipeline_hierarchies.py`.

### Результати верифікації (2026-05-11)

| Метрика | Очікувано | Реально |
|---|---|---|
| Hierarchy_Path NULL у 5 ієрархічних Dim | 0 | **0** ✓ |
| Unknown member у 8 Dim | 1 кожен | **1 кожен** ✓ |
| Direction_ID NULL у Departments | 0 з 387 | **0 з 387** ✓ |
| DDS Is_Excluded_From_Cashflow True | ≥7 (baseline) | **7** ✓ |
| DDS CFS_Section заповнено | за 1С даними | **0 з 426** (баг даних: фінкоманда не заповнила реквізит `А_РазделCFS` у 1С — не наш баг) |
| Power BI hierarchies drill-down | 14+ Level1 значень | **14+** ✓ |
| **Acceptance gate** (Globyno-2 Feb 2026 Income, PL_ЕРП) | **38 432 968.66 ₴** | **38 432 968,66 ₴** ✓ |

ETL run: 61 338 рядків, status=Success, ~3 хв.

---

## §2 — SQL OlapBASERP нові колонки і таблиці (контекст для Stage v3.7)

DDL у SQL Server для Stage v3.7 уже був готовий (Stage v3.6 commits `06e7f380b`, `80426c26e`); ETL у Stage v3.7 просто **заповнив** колонки що раніше лишались NULL.

### Колонки додані у 5 ієрархічних Dim

```sql
ALTER TABLE Dim_DDS_Articles      ADD Hierarchy_Path nvarchar(500), Hierarchy_Depth int, Level1..Level5 nvarchar(150);
ALTER TABLE Dim_Departments       ADD Hierarchy_Path nvarchar(500), Hierarchy_Depth int, Level1..Level5 nvarchar(150), Direction_ID char(32);
ALTER TABLE Dim_Items             ADD Hierarchy_Path nvarchar(500), Hierarchy_Depth int, Level1..Level5 nvarchar(150);
ALTER TABLE Dim_Expense_Articles  ADD Hierarchy_Path nvarchar(500), Hierarchy_Depth int, Level1..Level5 nvarchar(150);
ALTER TABLE Dim_Partners          ADD Hierarchy_Path nvarchar(500), Hierarchy_Depth int, Level1..Level5 nvarchar(150);  -- сама таблиця нова
```

### Dim_DDS_Articles нові поля

```sql
ALTER TABLE Dim_DDS_Articles ADD
    CFS_Section                 varchar(15) NULL,         -- денормалізовано з А_РазделCFS enum
    Is_Excluded_From_Cashflow   bit NOT NULL DEFAULT 0;   -- = А_ИсключатьИзОтчетаCashflow
```

### Нова таблиця Dim_Partners (4919 рядків)

```sql
CREATE TABLE Dim_Partners (
    Partner_ID            char(32) PRIMARY KEY,
    Partner_Code          varchar(50) NULL,
    Partner_Name          nvarchar(150) NOT NULL,
    Parent_ID             char(32) NULL,
    Is_Group              bit NOT NULL DEFAULT 0,
    Marked_For_Deletion   bit NOT NULL DEFAULT 0,
    Hierarchy_Path        nvarchar(500) NULL,
    Hierarchy_Depth       int NULL,
    Level1..Level5        nvarchar(150) NULL,
    Loaded_At             datetime2 NOT NULL DEFAULT SYSDATETIME()
);
```

Джерело: `Справочник.Партнеры` (SQL `_Reference360`). FK з Dim_Counterparties.Partner_ID на Dim_Partners.Partner_ID.

### Dim_Counterparties додано Partner_ID

```sql
ALTER TABLE Dim_Counterparties ADD Partner_ID char(32) NULL;  -- FK на Dim_Partners
```

Джерело: `Контрагенты.Партнер` (SQL `_Reference263._Fld31169RRef`).

---

## §3 — Stage v3.7 entry

| Параметр | Значення |
|---|---|
| Дата | 2026-05-11 |
| Branch | `claude/infallible-rosalind-d59cc5` |
| Commit | (TBD після `git commit`) |
| ETL run | run_id=217, status=Success, rows=61338, duration≈3 хв |
| Acceptance | Globyno-2 Feb 2026 Income (PL_ЕРП) = **38 432 968,66 ₴** ✓ |
| Critical files | `pipelines/dim_catalogs.json`, `scripts/update_pipeline_hierarchies.py`, `olap_changelog_2026_05.md` (NEW), `olap_powerbi_pl_pbix.md` (STATUS v3.7), `olap_sql_schema.md` (19 Dim) |

### Power BI column renames (2026-05-11)

Виконано через MCP `column_operations.Rename` у моделі PL.pbix (потребує Save .pbix для збереження у файл):

| Таблиця PBI | OlapBASERP column | Стара PBI назва → Нова |
|---|---|---|
| `СтатьиДвиженияДенежныхСредств` | `Is_Excluded_From_Cashflow` | `А_ИсключатьИзОтчетаCashflow` → **`Исключить из Cashflow`** |

Збереження alias на рівні моделі (не SQL DDL) — для природної кирилічної назви у візуалах і фільтрах.

### Знайдені під час реалізації проблеми

1. **DDL Dim_DDS_Articles.Is_Excluded_From_Cashflow визначений як NOT NULL**, але SQL Server `_Reference529._Fld56081` має 71 з 425 NULL (рядки де реквізит не заповнений у 1С). Без COALESCE ETL падає на bulk_insert.
   - **Fix:** у CTE root + recursive — `COALESCE(d._Fld56081, 0x00)` (FALSE).

2. **CFS_Section усі NULL** (`А_РазделCFS` реквізит порожній у 1С на 425 рядках з 425). Баг даних, не ETL — потрібне рішення від фінкоманди.

3. **Expense Articles CTE повертає 340 замість 344+1=345** (5 orphan-рядків з некоректним Parent reference виключені recursive walk). Старий sql_backend брав їх; CTE — ні. Acceptance gate не зачеплений (orphan rows не з'являються у Fact_PnL), тому залишено як є.

4. **Power BI Refresh "Automatic" не оновлює Import-mode дані** — потрібен `Refresh "Full"`. У Stage v3.7 успішно через `mcp__powerbi-modeling-mcp__model_operations.Refresh refreshType=Full`.

5. **Dim_Counterparties не має `_Code` колонки у поточному mapping** — `code_col=""` у `flat_with_unknown_sql()` повертає `N'' AS Код`.

---

## §5 — Перенос А_ОтчетБаланс_Свод → Fact_Balance + фікси (2026-05-17)

**Контекст:** перший реальний обмін регістра `А_ОтчетБаланс_Свод`
(проведений `Свод_СебестоимостьТоваров`, січень/ТОВ, 4775 рядків) у
`OlapBASERP.Fact_Balance`.

**Зроблено:**

1. **DDL застосовано:** `apply_07_balance.py` → `Fact_Balance` +
   `Dim_PAP_Articles` створені (ідемпотентно).
2. **`dim_pap_articles`** ETL — 54 рядки (full_reload, з `_Chrc1770`).
3. **`fact_balance --period 2026-01`** — **4775 рядків** перенесено;
   acceptance: Fact_Balance **1:1 == регістр** `А_ОтчетБаланс_Свод`
   (rows/Σ 4 ресурси) == `ПАП.ОстаткиИОбороты` до копійки; «Товары на
   оптовых складах» Sum_Close=**83 627 719,44** (== еталон ПАП/Отчёт).
   Тест: `_Rarzrabotki/Python/test/acceptance_fact_balance_seb.py`.

**Фікс A — Source enum (spec↔реальність):** `А_ОтчетБаланс_Свод.Source`
ФАКТИЧНО = типове `Перечисление.ИсточникиУправленческогоБаланса` (31
значення), а НЕ кастомне `А_ИсточникБаланса` (7) зі spec v3. ETL падав
`Cannot insert NULL into Source`. Виправлено (Python ETL, 1С не чіпали):
`enum_resolver.py` FROZEN_ENUMS += `ИсточникиУправленческогоБаланса` (31,
_EnumOrder 0..30; FROZEN бо в `_Enum1234` немає `_Description`);
`pipelines/fact_balance.json` Source→правильний enum;
`mapping/refresh_mapping.py` WHITELIST += перечислення; regenered
baserp_storage.json (`_Enum1234`).

**Фікс B — `dim_documents` вилучено з default:** джерело
`РегистрСведений.А_ДокументРасшифровка` ВІДСУТНЄ у поточній BaseERP
(немає `_InfoRg56031`) → `python main.py` падав на 3-му кроці.
`main.py` `ALL_DEFAULT_PIPELINES` -= `dim_documents` (коментар з умовою
повернення). Повний `python main.py` тепер проходить: dim_catalogs
61 338, fact_pnl 12 033, fact_cashflow 62 045 — Success.

### Знахідки

1. **`test_etl_acceptance_balance.py` (штатний) FAIL — очікувано:** він
   вимагає `Σ Sum_Close≈0` (Актив=Пасив, ПОВНИЙ баланс з усіма Свод_*).
   Зараз у `Документ.А_ФинРез_Баланс` активна ТІЛЬКИ
   `Свод_СебестоимостьТоваров` → Σ Sum_Close=83 627 719,44 (одна стаття).
   Це НЕ дефект обміну. Релевантна перевірка часткового свода —
   `acceptance_fact_balance_seb.py`. Повний штатний тест стане
   застосовним коли всі `Свод_*` активують (див.
   `knowledge_Balanse/balanse_pattern_and_roadmap.md` §6 roadmap).
2. ETL `raw_sql` extractor НЕ читає `baserp_storage.json` (mapping
   потрібен лише `sql_backend`/enum_resolver). Зміна модуля документа
   `А_ФинРез_Баланс` (не структури регістра) refresh_mapping НЕ вимагає.

### Status
✅ DONE — обмін балансу працює. Повний баланс в OLAP зʼявиться
автоматично коли в 1С активують решту `Свод_*` (ETL лише копіює регістр).
