# OLAP — Python ETL Pipeline (Stage 3, ✅ DONE)

> **STATUS:** ✅ Stage 3 реалізовано 2026-05-03. Архітектура SQL-first, не COM-first як попередній чорновик плану. Acceptance gate (Глобино-2 / ERP_Income / Feb 2026 = 38 432 968.66 ₴ exact) проходить — 10/10 pytest.
>
> **Live код:** [`_Rarzrabotki/Olap/Ai_Olap/`](../../Olap/Ai_Olap/)
> **Final commit:** `952c46db7` (Stage 10: README); merged into main.

---

## Призначення

Python-оркестратор переливає дані з 1С BAS ERP 2.5 (`Ref="BaseERP"` MSSQL backend) у BI-базу `OlapBASERP` (24 таблиці) за нічним розкладом.

**Принцип:** ETL — чистий копіювач без бізнес-логіки. Уся свёртка вже зроблена в 1С при проведенні документів `А_ФинРез_PL/DDS`. ETL переносить готові регістри сведень + 16 справочників-вимірювань 1:1 у SQL та додає drill-down URL.

---

## Архітектурна зміна: SQL-first замість COM-first

**Попередній чорновик** цього файлу описував COM-first — кожен extractor читав 1С через `V83.COMConnector` + 1С Query Language. **Реалізовано** SQL-first:

```
┌──────────────────┐    pyodbc     ┌─────────────┐    pyodbc    ┌──────────────┐
│  BaseERP MSSQL   │ ────────────> │ sql_backend │ ──────────>  │  OlapBASERP  │
│  (_Reference*,   │   (primary)   │  Extractor  │              │  (24 tables) │
│   _InfoRg*, ...) │               └─────────────┘              └──────────────┘
└──────────────────┘
        │
        │  V83.COMConnector       ┌─────────────┐
        └──────────────────────>  │ ComExtractor │  (only for virtual tables:
                  (fallback only) │              │   .Остатки/.Обороты/.ОстаткиИОбороты)
                                  └─────────────┘
```

**Чому SQL-first**: BAS ERP 2.5 на cluster-mode зберігає всі дані у MSSQL з автогенерованими іменами таблиць (`_Reference329` для Справочник.Организації, `_InfoRg55970` для А_ОтчетPL_Свод). Прямий pyodbc-зчитування **у рази швидше** ніж COM (на 3937 рядках Fact_PnL за лютий — 58 ms vs ~2 s через COM). COM лишається для віртуальних таблиць де платформа додає обчислювані колонки які raw-таблиці не дають.

**Mapping `1С_object → SQL_table + {1С_field: SQL_column}`** — генерується одним викликом `mapping/refresh_mapping.py` (Python COM-обгортка над платформенним методом `ПолучитьСтруктуруХраненияБазыДанных()`) і зберігається у `mapping/baserp_storage.json` (commited у git, 56 об'єктів, 150KB).

---

## Структура каталогу (live)

```
_Rarzrabotki/Olap/Ai_Olap/
├── ai_olap/                              ← Python пакет
│   ├── __init__.py
│   ├── core/                             — connections, logging, exceptions, decorators
│   │   ├── connections.py                — get_baserp_sql / get_olap_sql / get_com_connection
│   │   ├── logging_setup.py              — structlog → logs/etl_YYYY-MM-DD.log
│   │   ├── exceptions.py                 — ETLException + 7 subclasses
│   │   └── decorators.py                 — @retry / @measure_time
│   ├── extractors/
│   │   ├── base.py                       — abstract Extractor
│   │   ├── sql_backend.py                — PRIMARY: pyodbc → BaseERP backend
│   │   ├── com.py                        — FALLBACK: V83 для віртуальних таблиць
│   │   └── factory.py                    — make_extractor(config)
│   ├── transformers/
│   │   ├── varbinary_to_uuid.py          — _IDRRef → char(32) hex; null UUID → None
│   │   ├── onec_date.py                  — datetime → 1st-of-month / .date()
│   │   ├── enum_resolver.py              — Enum UUID → frozen string via _EnumOrder
│   │   ├── drill_down.py                 — recorder + meta → e1cib/data/... URL
│   │   ├── column_mapper.py              — rename + drop + defaults
│   │   └── pipeline.py                   — REGISTRY + apply(rows, steps, options)
│   ├── loaders/
│   │   ├── base.py                       — get_table_columns + bulk_insert
│   │   ├── dim.py                        — TRUNCATE + bulk insert
│   │   ├── fact.py                       — DELETE WHERE Period_Month=? + bulk insert
│   │   ├── bridge.py                     — TRUNCATE + bulk insert
│   │   ├── etl_runs.py                   — open_run / close_run lifecycle
│   │   └── factory.py                    — make_loader(config)
│   ├── orchestrator/
│   │   ├── pipeline.py                   — Pipeline (E→T→L) + auto_period_params
│   │   ├── runner.py                     — load + validate + ETL_Runs wrap
│   │   └── scheduler.py                  — APScheduler BlockingScheduler
│   ├── config/
│   │   ├── schema.py                     — JSON Schema for pipelines/*.json
│   │   ├── loader.py                     — load_pipeline_config / discover_pipelines
│   │   └── validator.py                  — jsonschema-based validation
│   └── utils/
│       └── mapping_resolver.py           — resolve(meta) → (sql_table, {field: column})
├── mapping/
│   ├── refresh_mapping.py                — COM-генератор JSON-mapping
│   └── baserp_storage.json               — 56 об'єктів, ~150KB (gitcommited)
├── pipelines/                            — 5 JSON-конфігів
│   ├── dim_catalogs.json                 — 16 Dim + 1 Bridge multistep, schedule '0 1 * * *'
│   ├── fact_pnl.json                     — Fact_PnL idempotent, '0 2 1 * *'
│   ├── fact_cashflow.json                — Fact_Cashflow, '0 2 1 * *'
│   ├── fact_cf_balance.json              — COM extractor для .ОстаткиИОбороты
│   └── etl_runs_keepalive.json           — daily heartbeat '0 6 * * *'
├── tests/                                — pytest 10/10 PASS
│   ├── test_etl_acceptance_globyno2.py   — 🎯 Глобино-2 = 38 432 968.66 ₴
│   ├── test_etl_row_counts.py            — Σ Fact_PnL=3937, Σ Fact_Cashflow=4652
│   ├── test_mapping_resolver.py
│   └── test_sql_backend_extractor.py
├── logs/                                 — JSON structured logs (gitignored)
├── .env / .env.example                   — 3 connection strings
├── requirements.txt                      — pyodbc, pywin32, APScheduler, jsonschema, structlog, dotenv, pytest
├── pyproject.toml                        — ruff + pytest config
├── main.py                               — CLI: --validate / --run-once / --scheduled / --refresh-mapping
├── README.md                             — quickstart + troubleshooting
├── IMPLEMENTATION_PROMPT.md              — superseded original COM-first plan
└── RESEARCH.md                           — framework choice (APScheduler+custom > Airflow/Prefect/Dagster)
```

**Реальний обсяг:** 11 commits, ~1900 рядків Python, 56 об'єктів у mapping JSON, 17 ETL pipeline steps.

---

## Connection Layer

### .env (gitignored)

```ini
BASERP_SQL_DSN=DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=BaseERP;UID=sa;PWD=Brw739182465!;TrustServerCertificate=yes
OLAP_SQL_DSN=DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;TrustServerCertificate=yes
BASERP_COM_CONN=Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"
LOG_LEVEL=INFO
MAPPING_PATH=mapping/baserp_storage.json
```

### ai_olap/core/connections.py — три точки входу

```python
def get_baserp_sql() -> pyodbc.Connection:
    """Read-only pyodbc до BaseERP backend (primary read path)."""
    return pyodbc.connect(_require_env("BASERP_SQL_DSN"), timeout=15, readonly=True)

def get_olap_sql() -> pyodbc.Connection:
    """Write pyodbc до OlapBASERP. autocommit=False, fast_executemany=True у курсорі."""
    conn = pyodbc.connect(_require_env("OLAP_SQL_DSN"), timeout=15, autocommit=False)
    conn.cursor().fast_executemany = True
    return conn

def get_com_connection() -> Any:
    """V83.COMConnector — lazy import; тільки для віртуальних таблиць/Свод-логіки."""
    import win32com.client
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(_require_env("BASERP_COM_CONN"))
```

---

## Mapping Resolver — серце SQL-first архітектури

### mapping/baserp_storage.json — як генерується

Платформа 1С експонує метод `ПолучитьСтруктуруХраненияБазыДанных()` (англ. `GetDBStorageStructureInfo()`), що для кожного об'єкта повертає `ОписаниеТаблицы` з полями `ИмяТаблицыХранения`, `Метаданные`, `Назначение`, `Поля` (внутрішня таблиця `ИмяПоля → ИмяПоляХранения`). Скрипт `mapping/refresh_mapping.py` викликає його через COM один раз і дампить структуру у JSON:

```jsonc
{
  "version": "8.3.20.x",
  "generated_at": "2026-05-03T13:15:42",
  "source_db": "BaseERP",
  "object_count": 56,
  "objects": {
    "Справочник.Организации": {
      "primary_table": "_Reference329",
      "tables": [{"name": "_Reference329", "purpose": "Основная", "fields": {...}}],
      "fields": {
        "Ссылка": "_IDRRef",
        "Наименование": "_Description",
        "ПометкаУдаления": "_Marked",
        "КодПоЕДРПОУ": "_Fld32799",
        ...
      }
    },
    "РегистрСведений.А_ОтчетPL_Свод": {
      "primary_table": "_InfoRg55970",
      "fields": {
        "Регистратор": "_RecorderRRef",
        "Source": "_Fld55984RRef",
        "Организация": "_Fld55971RRef",
        "СуммаЕРПГрн": "_Fld55982",
        ...
      }
    }
  }
}
```

Файл commited у git; рефреш — `python mapping/refresh_mapping.py` після кожної суттєвої зміни конфігурації 1С (нові реквізити, нові регістри). При тому ж стані конфіги SQL-таблиці stable.

### ai_olap/utils/mapping_resolver.py

```python
@functools.lru_cache(maxsize=1)
def _load_mapping(path: str) -> dict: ...

def resolve(meta_full: str) -> tuple[str, dict[str, str]]:
    """
    resolve("Справочник.Организации")
        -> ("_Reference329", {"Ссылка": "_IDRRef", "Наименование": "_Description", ...})
    """
```

Один `lru_cache(1)` вистачає для всього процесу. Скидати кеш — `reload_cache()` (викликається `main.py --refresh-mapping`).

---

## Pipeline JSON-конфіги (декларативно)

Кожен pipeline — JSON-файл у `pipelines/` зі схемою з `ai_olap/config/schema.py`. Pipeline містить `steps[]`; кожен крок — `extractor` + `transformer` + `loader`.

### Приклад: один Dim step з dim_catalogs.json

```json
{
  "step_id": "dim_organizations",
  "extractor": {
    "type": "sql_backend",
    "object": "Справочник.Организации",
    "fields": ["Ссылка", "Наименование", "ПометкаУдаления"],
    "where": "_Fld32799 = '40645273'"
  },
  "transformer": {
    "steps": ["varbinary_to_uuid", "column_mapper"],
    "options": {
      "column_mapper": {
        "column_map": {
          "Ссылка": "Organization_ID",
          "Наименование": "Organization_Name",
          "ПометкаУдаления": "Marked_For_Deletion"
        },
        "defaults": {"Is_Group": false, "Marked_For_Deletion": false}
      }
    }
  },
  "loader": {"target_table": "Dim_Organizations", "mode": "full_reload"}
}
```

### Приклад: Fact_PnL step (raw_sql + JOIN до документа)

```json
{
  "step_id": "load",
  "extractor": {
    "type": "raw_sql",
    "auto_period_params": true,
    "raw_sql": "SELECT r._RecorderRRef AS Recorder_PL_ID, r._Fld55984RRef AS Source, d._Date_Time AS Period, r._Fld55971RRef AS Organization_ID, ..., r._Fld55982 AS Sum_ERP_Grn FROM _InfoRg55970 r INNER JOIN _Document56011 d ON d._IDRRef = r._RecorderRRef WHERE r._Active = 0x01 AND r._Fld55971RRef = 0x80D3000C29BBAC2311E653F06BEE36B2 AND d._Date_Time >= ? AND d._Date_Time < ?"
  },
  "transformer": {
    "steps": ["varbinary_to_uuid", "enum_resolver", "onec_date", "column_mapper"],
    "options": {
      "enum_resolver": {"column_to_enum": {"Source": "Перечисление.А_ИсточникPL"}},
      "onec_date": {"columns": ["Period_Month"], "mode": "month"},
      "column_mapper": {"column_map": { ... }}
    }
  },
  "loader": {
    "target_table": "Fact_PnL",
    "mode": "idempotent_period",
    "period_column": "Period_Month"
  }
}
```

`auto_period_params=true` каже Pipeline.run() підставити `[period_start, period_end]` як SQL-параметри (з +2000 рік-офсетом — див. lessons learned).

### 5 поточних pipeline'ів

| файл | що робить | schedule |
|------|-----------|----------|
| `dim_catalogs.json` | full reload 16 Dim_* + 1 Bridge у одному multi-step pipeline | `0 1 * * *` |
| `fact_pnl.json` | idempotent reload Fact_PnL для періоду | `0 2 1 * *` |
| `fact_cashflow.json` | idempotent reload Fact_Cashflow | `0 2 1 * *` |
| `fact_cf_balance.json` | balance через COM `.ОстаткиИОбороты` (placeholder) | `0 3 1 * *` |
| `etl_runs_keepalive.json` | daily heartbeat | `0 6 * * *` |

---

## Extractor types

| `extractor.type` | клас | коли |
|------------------|------|------|
| `sql_backend`    | `SqlBackendExtractor` | усі Dim, Bridge, прості SELECT-и |
| `raw_sql`        | `SqlBackendExtractor` | Fact-таблиці (треба JOIN до документа) |
| `com`            | `ComExtractor`        | віртуальні таблиці (`.Остатки`, `.Обороты`, `.ОстаткиИОбороты`) |

`SqlBackendExtractor` через `mapping_resolver.resolve()` будує `SELECT _IDRRef AS "Ссылка", _Description AS "Наименование" FROM _Reference329 [WHERE ...]` — повертає рядки із 1С-іменами полів. Якщо поле відсутнє у мапінгу (системні `_Folder` для неієрархічних довідників, тощо) — extractor warn'ить і пропускає, transformer-defaults заповнюють.

---

## Transformer chain

Кожен transformer — функція `(rows: list[dict], **opts) -> list[dict]`. REGISTRY у `ai_olap/transformers/pipeline.py`:

| name | що робить |
|------|-----------|
| `varbinary_to_uuid` | `bytes(16) → hex(32 chars)`. 1-byte → bool. Null UUID (b"\\x00"\*16) → None. |
| `onec_date`         | datetime → `dt.datetime(y, m, 1)` (Period_Month) або `.date()`. |
| `enum_resolver`     | UUID → frozen string. Hardcoded списки `FROZEN_ENUMS` (8/3/4 значення); SQL-lookup `_Enum*._EnumOrder`, lru_cache. |
| `drill_down`        | `(meta, uuid_hex) → "e1cib/data/<meta>?ref=<hex>"`. |
| `column_mapper`     | `{src: dst}` rename + drop unmapped + `defaults={dst: value}` (працює і коли src відсутнє у row, і коли value у row = None). |

Pipeline крок викликає їх через `tpipe.apply(rows, steps_list, options)`.

---

## Loader modes

| mode | semantics | таблиці |
|------|-----------|---------|
| `full_reload`         | `TRUNCATE TABLE` + `INSERT` | усі Dim_* + Bridge_* |
| `idempotent_period`   | `DELETE WHERE Period_Month = ?` + `INSERT` | Fact_PnL, Fact_Cashflow, Fact_CF_Balance |
| `append`              | `INSERT` без DELETE | ETL_Runs (через окремі helpers `open_run`/`close_run`) |

`get_table_columns(table)` introspect-ить `sys.tables`/`sys.columns` і виключає IDENTITY (`Fact_ID`, `Run_ID`) та DEFAULT-заповнені (`Loaded_At = sysdatetime()`). `bulk_insert(table, rows)` використовує `fast_executemany=True` + batches по 1000.

---

## Orchestrator

### main.py — 4 режими

```bash
python main.py --validate                                # перевіряє всі pipelines/*.json по schema
python main.py --run-once dim_catalogs                   # один pipeline (Dim — без --period)
python main.py --run-once fact_pnl --period 2026-02      # Fact — з --period YYYY-MM
python main.py --scheduled                               # APScheduler BlockingScheduler daemon
python main.py --refresh-mapping                         # скинути кеш mapping_resolver
```

### ai_olap/orchestrator/pipeline.py

`Pipeline(cfg, period=date).run()` ітерує `cfg["steps"]`:
1. `extractor_cfg = dict(step["extractor"])`. Якщо `auto_period_params=true` і `period` заданий — додає `[period_start, period_end]` (+2000 рік-офсет!) до `params`.
2. `extractor = make_extractor(extractor_cfg); rows = extractor.extract()` — повертає list[dict].
3. Якщо `transformer.steps` є — `tpipe.apply(rows, steps, options)`.
4. `loader_cfg = dict(step["loader"])`. Для `idempotent_period` додає `period`. `make_loader(loader_cfg).load(rows)`.
5. Логує `step done` через structlog.

Excepton у будь-якому кроці бульбашить вгору; runner (`runner.py`) обгортає Pipeline у `open_run/close_run` і пише traceback у `ETL_Runs.Error`.

### Scheduler

`EtlScheduler.register_jobs()` для кожного `pipelines/*.json` парсить `schedule` як cron і реєструє `run_pipeline(cfg, script=path.stem)` у `BlockingScheduler` з `coalesce=True, max_instances=1, misfire_grace_time=600`. `start()` блокує процес — підходить для Windows Task Scheduler як daemon.

---

## Lessons learned (під час реалізації)

Реальні відкриття, що відрізняли реалізацію від чорновика плану:

### 1. BaseERP cluster backend зберігає `_Date_Time` з +2000 річним офсетом

Документ за 2026-02-15 у MSSQL зберігається як `4026-02-15`. Це специфіка cluster-mode 1С (file-mode без офсету). `Pipeline.auto_period_params` додає `period_offset_years: 2000` (default) до Period перед SQL `?`-bind. Якщо пишете кастомний raw_sql — додавайте 2000 самостійно або запозичте логіку.

### 2. `ПолучитьСтруктуруХраненияБазыДанных()` не повертає системні поля

`_Folder`, `_ParentIDRRef`, `_PredefinedID`, `_Period` методом не виводяться. Якщо просите `ЭтоГруппа`/`Родитель` від неієрархічного довідника — extractor warn'ить і пропускає; `column_mapper.defaults` заповнює.

### 3. Не всі довідники мають `Код`

`Справочник.Организации` у цій ERP-конфіги не має реквізиту `Код` (тільки `Наименование`); pipeline для Dim_Organizations не передає `Код` у extract. Це ERP-config-specific — інші конфіги можуть мати.

### 4. `Справочник.Подразделения` ≠ `Справочник.СтруктураПредприятия`

В ERP стандартний довідник — `СтруктураПредприятия`. `Подразделения` нема. Whitelist у `mapping/refresh_mapping.py` стрінює `Подразделения` мовчки (warning у консолі).

### 5. NULL UUID — окремий випадок

В 1С backend всі-нулі (16×0x00) — це "пустая ссылка". `varbinary_to_uuid` повертає `None` (не "00"\*16). Це корисно для Fact-таблиць де FK може бути порожній (наприклад, Касса не має `BankAccount_ID` → потрібен `defaults: {"BankAccount_ID": "00000000000000000000000000000000"}` або NULLable FK у схемі).

### 6. Filtering на власну організацію через КодПоЕДРПОУ

Усі Fact-pipeline'и фільтрують по `Organization_ID = 0x80D3000C29BBAC2311E653F06BEE36B2` (UUID ТОВ ІНДАСТРІАЛБУД). У `Dim_Organizations` фільтр через `WHERE _Fld32799 = '40645273'` (`КодПоЕДРПОУ` SQL колонка). Якщо група компаній отримає другу юр.особу — додавайте explicit (а не пасивно).

### 7. `Fact_Cashflow.BankAccount_ID NOT NULL` — обмеження схеми

Stage 2 DDL заклав це жорстко, але в реальних даних Касса має `BankAccount_ID = NULL` (це не банк). Поки workaround — placeholder zero-UUID; правильніше — alter schema до NULLable. TBD у Stage 4 ревізії.

### 8. Bridge_PLArticle_DDS поки не наповнюється

Реквізит `СтатьяДвиженияДенежныхСредств` на `Справочник.А_Статьи_PL` у поточній конфігурації не заповнений (`_Fld55816RRef = 0x00...00` для всіх статей). Bridge step повертає 0 рядків. Це не блокує acceptance gate (PnL і Cashflow незалежні), але до Stage 4 фінансист має заповнити мапу або переглянути джерело.

---

## Acceptance verification

### Перевірка через pytest

```bash
cd _Rarzrabotki/Olap/Ai_Olap
.venv/Scripts/python.exe -m pytest tests/ -v
# 10 passed in 0.22s
```

### Acceptance numbers (Feb 2026)

| metric | expected | actual | source |
|--------|----------|--------|--------|
| Fact_PnL Σ rows                            | 3937           | 3937           | `test_fact_pnl_february_2026_row_count` |
| Fact_Cashflow Σ rows                       | 4652           | 4652           | `test_fact_cashflow_february_2026_row_count` |
| Fact_PnL distinct Source                   | 7              | 7              | `test_fact_pnl_distinct_sources` |
| Dim_Organizations rows                     | 1 (own only)   | 1 ІНДАСТРІАЛБУД | `test_dim_organizations_only_own_company` |
| 🎯 **Глобино-2 / ERP_Income / 2026-02-01** | **38 432 968.66 ₴** | **38 432 968.66 ₴** | `test_globyno_erp_income_february_2026` |

Tolerance acceptance gate: ± 0.01 ₴.

### Перевірка через SQL

```sql
SELECT SUM(F.Sum_ERP_Grn)
FROM Fact_PnL F
JOIN Dim_Departments D ON F.Department_ID = D.Department_ID
WHERE D.Department_Name = N'Глобино-2'
  AND F.Source = 'ERP_Income'
  AND F.Period_Month = '2026-02-01';
-- 38432968.66
```

---

## Cross-references

- Як працює свёртка в 1С: [olap_obrabotka_provedeniya.md](olap_obrabotka_provedeniya.md)
- 8 нових об'єктів метаданих: [olap_1c_objects.md](olap_1c_objects.md)
- Куди вантажиться (DDL OlapBASERP): [olap_sql_schema.md](olap_sql_schema.md)
- Що читається з ERP: [olap_data_sources_erp.md](olap_data_sources_erp.md)
- Acceptance criteria: [olap_acceptance_etalons.md](olap_acceptance_etalons.md)
- Архітектурні принципи: [olap_architecture_overview.md](olap_architecture_overview.md)
- Live код: [`_Rarzrabotki/Olap/Ai_Olap/`](../../Olap/Ai_Olap/)
- README з quickstart: [`Ai_Olap/README.md`](../../Olap/Ai_Olap/README.md)
