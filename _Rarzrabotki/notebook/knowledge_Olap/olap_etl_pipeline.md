# OLAP — Python ETL Pipeline (Stage 3, ✅ DONE)

> **STATUS:** ✅ Stage 3 реалізовано 2026-05-03. Архітектура SQL-first, не COM-first як попередній чорновик плану. Acceptance gate (Глобино-2 / ERP_Income / Feb 2026 = 38 432 968.66 ₴ exact) проходить — 10/10 pytest.
>
> **Live код:** [`_Rarzrabotki/Olap/Ai_Olap/`](../../Olap/Ai_Olap/)
> **Останній commit:** `8d5ebf3a1` (default mode: full reload без прапорців); merged into main.

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
| `onec_date`         | datetime → `dt.datetime(y, m, 1)` (Period_Month) або `.date()` — in-place на listed columns. |
| `period_offset_fix` | **NEW (commit 8d5ebf3a1).** Знімає +2000 рік-офсет з `_Date_Time` (4026 → 2026) та деривує `Period_Month = date(year, month, 1)`. Threshold-based — безпечно для дат що вже без офсету. Опції: `source_column`, `period_month_column`, `offset_years` (default 2000), `epoch_threshold` (default 3000). |
| `enum_resolver`     | UUID → frozen string. Hardcoded списки `FROZEN_ENUMS` (8/3/4 значення); SQL-lookup `_Enum*._EnumOrder`, lru_cache. |
| `drill_down`        | `(meta, uuid_hex) → "e1cib/data/<meta>?ref=<hex>"`. |
| `column_mapper`     | `{src: dst}` rename + drop unmapped + `defaults={dst: value}` (працює і коли src відсутнє у row, і коли value у row = None). |

Pipeline крок викликає їх через `tpipe.apply(rows, steps_list, options)`.

---

## Loader modes

| mode | semantics | таблиці |
|------|-----------|---------|
| `full_reload`         | `TRUNCATE TABLE` + `INSERT` | усі Dim_* + Bridge_* |
| `idempotent_period`   | `DELETE WHERE Period_Month = ?` + `INSERT` | Fact_PnL, Fact_Cashflow, Fact_CF_Balance (коли передано `--period`) |
| `append`              | `INSERT` без DELETE | ETL_Runs (через окремі helpers `open_run`/`close_run`) |

**Auto-degrade (commit 8d5ebf3a1):** якщо у `Pipeline.run()` `self.period is None` і loader.mode = `idempotent_period`, режим **автоматично перетворюється на `full_reload`**. Тобто `python main.py` без `--period` робить TRUNCATE+INSERT для Fact-таблиць замість DELETE WHERE. Це необхідно бо raw_sql тоді витягує всі дати, і per-period DELETE не має сенсу. У логах буде `dim full reload component=dim_loader rows=N table=Fact_PnL` — це правильно (factory повертає `DimLoader` для `full_reload` mode незалежно від таблиці).

`get_table_columns(table)` introspect-ить `sys.tables`/`sys.columns` і виключає IDENTITY (`Fact_ID`, `Run_ID`) та DEFAULT-заповнені (`Loaded_At = sysdatetime()`). `bulk_insert(table, rows)` використовує `fast_executemany=True` + batches по 1000.

---

## Orchestrator

### main.py — режими CLI

```bash
python main.py                                           # DEFAULT: повний прогон (validate + 3 Dim + 3 Fact, усі робочі), TRUNCATE+INSERT всіх дат
python main.py --period 2026-02                          # Той самий ланцюжок, але Fact-pipelines обмежуються одним місяцем (idempotent DELETE WHERE + INSERT)
python main.py --validate                                # перевіряє всі pipelines/*.json по schema
python main.py --run-once dim_catalogs                   # один pipeline (Dim — без --period)
python main.py --run-once fact_pnl --period 2026-02      # один Fact — з --period YYYY-MM (idempotent)
python main.py --run-once fact_pnl                       # один Fact без --period (full reload Fact_PnL)
python main.py --scheduled                               # APScheduler BlockingScheduler daemon (поки НЕ використовуємо)
python main.py --refresh-mapping                         # скинути кеш mapping_resolver (не регенерує JSON!)
```

### Default mode (без прапорців) — `cmd_all`

Коли не передано жодного режим-flag, виконуються **6 pipeline'ів підряд** (3 Dim + 3 Fact) з `ALL_DEFAULT_PIPELINES`:

1. `validate` — перевірка JSON-конфігів
2. `dim_catalogs` — full reload 16 Dim + 1 Bridge
3. `dim_pap_articles` — full reload Dim_PAP_Articles (FK Fact_Balance, канон OD-9)
4. `dim_denezhnye_sredstva` — full reload Dim_DenezhnyeSredstva (FK Fact_Cashflow/Balance)
5. `fact_pnl` — full reload `_InfoRg55970`
6. `fact_cashflow` — full reload `_InfoRg55992`
7. `fact_balance` — full reload `_InfoRg56091` (idempotent_period без `--period` → full_reload усіх дат)

Зупиняється на першій помилці. **Виключені з дефолту** (зламані/неповні — падали б весь ланцюг; тримати поза `ALL_DEFAULT_PIPELINES` до фіксу, запускати/чинити окремо через `--run-once`):
- `dim_documents` — джерело `РегистрСведений.А_ДокументРасшифровка` ВІДСУТНЄ у поточній BaseERP (немає `_InfoRg56031`). Повернути коли об'єкт створять у 1С.
- `fact_cf_balance` — COM-запит `.ОстаткиИОбороти` падає синтаксичною помилкою 1С (перевірено 2026-05-19, run_id=296 Failed); ще й фреймворк не передає `period` у COM-екстрактор. Повернути після фіксу пайплайна.
- `etl_runs_keepalive` — heartbeat, не data-pipeline.

> **Оновлення 2026-05-19 (на вимогу користувача):** дефолт розширено —
> `python main.py` (без прапорців) тепер оновлює **усі робочі Dim і Fact за
> весь період**. `ALL_DEFAULT_PIPELINES = ["dim_catalogs", "dim_pap_articles",
> "dim_denezhnye_sredstva", "fact_pnl", "fact_cashflow", "fact_balance"]`.
> Раніше (2026-05-17) Balance/нові Dim були поза дефолтом → `Fact_Balance` не
> оновлювався звичайним прогоном (першопричина бага «контрагенти Dec2025 не
> в розшифровці»). Verified row-counts повного прогону 2026-05-19
> (run_id 297–302, усі Success): dim_catalogs **75 975**, dim_pap_articles
> **54**, dim_denezhnye_sredstva **5 907**, fact_pnl **12 033**,
> fact_cashflow **62 045**, fact_balance **23 799** (усі періоди).
> `--period YYYY-MM` лишається для точкового idempotent-перезавантаження місяця.

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

Документ за 2026-02-15 у MSSQL зберігається як `4026-02-15`. Це специфіка cluster-mode 1С (file-mode без офсету). `Pipeline.auto_period_params` додає `period_offset_years: 2000` (default) до Period перед SQL `?`-bind. Якщо `--period` не передано — підставляє широкий діапазон `[date(2001,1,1), date(9999,1,1)]` що покриває всі дати. Якщо пишете кастомний raw_sql — додавайте 2000 самостійно або запозичте логіку.

Для **колонок у row** використовуй transformer **`period_offset_fix`** — він знімає 2000 років з `Period` та деривує `Period_Month` (без офсету). Конфіг в JSON:
```json
"period_offset_fix": {"source_column": "Period", "period_month_column": "Period_Month"}
```
Це дозволяє Power BI бачити правильні роки 2026 у Fact-таблицях, а не 4026.

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

---

## Balance Stage (2026-05-16) — pipelines fact_balance + dim_pap_articles

Два нові декларативні pipeline (дзеркало fact_cashflow / dim_income_articles):

| pipeline | джерело SQL backend | loader | rows (січ.2026/ТОВ) |
|---|---|---|---|
| `pipelines/dim_pap_articles.json` | `_Chrc1770` (ПВХ.СтатьиАктивовПассивов) + enum_resolver АктивПассив→`_Enum1172` | full_reload | 54 |
| `pipelines/fact_balance.json` | `_InfoRg56091` r INNER JOIN `_Document56084` d ON d._IDRRef=r._RecorderRRef | idempotent_period (Period_Month) | 11 561 |

**Rule #-1 (discovery-first):** перед raw_sql ОБОВ'ЯЗКОВО resolve `_InfoRg/_Fld`
через `mapping/refresh_mapping.py`. WHITELIST +5: `РегистрСведений.А_ОтчетБаланс_Свод`,
`Документ.А_ФинРез_Баланс`, `ПланВидовХарактеристик.СтатьиАктивовПассивов`,
`Перечисление.А_ИсточникБаланса`, `Перечисление.ВидыСтатейУправленческогоБаланса`.
`enum_resolver.FROZEN_ENUMS` +2: `ВидыСтатейУправленческогоБаланса` (3,
ASCII Aktiv/Passiv/AktivPassiv — DAX PL.pbix + varchar(15)) та
**`ИсточникиУправленческогоБаланса`** (див. виправлення нижче).
baserp_storage.json: additive — cashflow/PnL не зачеплені.

> **⚠️ Виправлення 2026-05-17 (spec↔реальність):** регістр
> `А_ОтчетБаланс_Свод.Source` ФАКТИЧНО зберігає **типове**
> `Перечисление.ИсточникиУправленческогоБаланса` (31 значення, _EnumOrder
> 0..30; `Свод_СебестоимостьТоваров` пише `СебестоимостьТоваров`), а **НЕ**
> кастомне `А_ИсточникБаланса` (7) зі spec v3. Перший прогон ETL падав
> `Cannot insert NULL into Source` (enum_resolver не знаходив UUID).
> **Фікс (Python ETL, 1С не чіпали):** `enum_resolver.py` FROZEN_ENUMS +=
> `ИсточникиУправленческогоБаланса` (31 значення, НЕ DYNAMIC — у `_Enum1234`
> немає `_Description`, лише `_IDRRef`/`_EnumOrder`); `pipelines/
> fact_balance.json` `Source`→`Перечисление.ИсточникиУправленческогоБаланса`;
> `mapping/refresh_mapping.py` WHITELIST += це перечислення (→ `_Enum1234`
> у baserp_storage.json). Це узгоджено з `knowledge_Balanse` (там регістр
> вже описаний з `ИсточникиУправленческогоБаланса`).
`scripts/introspect_balance_fields.py` — дамп resolved _Fld + additive-guard.

resolved (джерело істини raw_sql): Source `_Fld56104RRef`, Орг `_Fld56092RRef`,
Стаття `_Fld56095RRef`, ДенСр `_Fld56124_RRRef` (composite), НачОст `_Fld56122`,
Приход `_Fld56119`, Расход `_Fld56120`, КонОст `_Fld56106`, Регистратор
`_RecorderRRef`; Period з `_Document56084._Date_Time` (реєстратор-підпорядкований
регістр — власного Періоду немає). Фільтр власна орг `0x80D3...36B2`.

**Запуск:** `python main.py --run-once dim_pap_articles` →
`python main.py --run-once fact_balance --period 2026-01` (venv `.venv/Scripts/python.exe`).

> **⚙️ 2026-05-18 — `Свод_ПрочиеАктивыПассивы_Прямой` LIVE (+колонка TaxType):**
> прямі рухи ПАП `Source=ПустаяСсылка` (0 JOIN, ВЫРАЗИТЬ Аналітика).
> Зміни Python ETL (1С не чіпали понад додане користувачем вимір ТипНалога):
> - `mapping/refresh_mapping.py` WHITELIST += `Перечисление.ТипыНалогов`
>   (→ `_Enum1651` у baserp_storage.json); `refresh_mapping.py` перезапущено
>   (А_ОтчетБаланс_Свод тепер resolve **`ТипНалога`→`_Fld56130RRef`**).
> - `enum_resolver.FROZEN_ENUMS += "Перечисление.ТипыНалогов"` — 14 значень
>   `_EnumOrder` 0..13 (НДС/НДФЛ/…/НачисленныйЕСВ/ВоенныйСбор/ЕдиныйНалог/
>   НалогНаПрибыль/ДругиеНалоги/Акциз; метаімена, не синоніми).
> - **`enum_resolver.transform`: пуста ссилка enum → `"ПустаяСсылка"`.**
>   Корінь: `varbinary_to_uuid._convert` мапить 16 нулів (`ALL_ZERO`)→`None`
>   ДО `enum_resolver` (рядок 20-21). Прямой пише `Source=ПустаяСсылка`
>   (16 нулів) → `None` → `Fact_Balance.Source NOT NULL` падав
>   `Cannot insert NULL`. Фікс: у `transform` для enum-колонки `val is None`
>   → `"ПустаяСсылка"` (1С-ім'я пустої ссилки; PnL/Cashflow Source ніколи не
>   порожній → не зачеплені; Глобино-2 PASS).
> - `pipelines/fact_balance.json`: raw_sql += `r._Fld56130RRef AS TaxType`;
>   `enum_resolver.column_to_enum += "TaxType":"Перечисление.ТипыНалогов"`;
>   `column_mapper.column_map += "TaxType":"TaxType"`.
> - SQL DDL OlapBASERP: `ALTER TABLE Fact_Balance ADD TaxType varchar(50)
>   COLLATE Cyrillic_General_CI_AS NULL` (як Source, NULLABLE — заповнено
>   лише у статті «Налоги», решта `"ПустаяСсылка"`).
> ETL `--run-once fact_balance --period 2025-12|2026-01|2026-02` Success
> (7682/7643/8208). Verify `scripts/verify_olap_balance_papdirect.py` PASS.
>
> **⚙️ 2026-05-18 — `Свод_ОплатаТруда` LIVE (БЕЗ змін ETL-конфігів):**
> статья «Оплата труда» зводиться БЕЗ аналітики під наявним
> `Source=ПустаяСсылка` (ПАП Источник=пусто, Статья=&ОТ; зеркало канон-
> Прямого, комплементарно `Т.Статья<>&ОТ`). **`fact_balance.json` /
> `enum_resolver` / `Dim_TaxTypes` / `refresh_mapping` НЕ міняються** (нових
> Source/вимірів немає; ОТ без субконто → TaxType="ПустаяСсылка").
> `refresh_mapping.py` НЕ потрібен (структура РС не змінилась). Потрібен лише
> `--run-once fact_balance --period 2025-12|2026-01|2026-02` (перечитає
> регістр; рядки «Оплата труда» з'являться під `Source=ПустаяСсылка`,
> `TaxType="ПустаяСсылка"`). Новий ИТОГ `Source=ПустаяСсылка` КО 2026-01 =
> **−108 631 177,36** (Прямой −101 434 478,92 + ОТ −7 196 698,44, серверно
> COM). `verify_olap_balance_papdirect.py` еталони оновлені (+«Оплата труда»).
> ✅ ETL виконано (3 періоди); `verify_olap_balance_papdirect.py` PASS.
>
> **⚙️ 2026-05-18 — `Свод_ПрочиеРасходыДоходы` LIVE → ПОВНИЙ БАЛАНС
> (БЕЗ змін ETL-конфігів):** Источник∈{ПрочиеДоходы/ПрочиеРасходы/
> ПартииПрочихРасходов} (FROZEN ИсточникиУправленческогоБаланса 31 —
> вже є). `fact_balance.json`/`enum_resolver`/Dim НЕ міняються; лише
> `--run-once fact_balance --period 2025-12|2026-01|2026-02` (дек 7757 /
> янв 7721 / лют 8319; +Source=ПрочиеРасходы). Регістр Σ Sum_Close (OD-3,
> ВСІ Source) = **0,00** ⇒ Актив=Пассив == штатний звіт; `verify_olap_
> balance_papdirect.py` доповнено блоком «ПОЛНЫЙ БАЛАНС» — PASS.
>
> **⚙️ 2026-05-18 — вимір `ТипПоказателя` → `Fact_Balance.TipPokazatelya`
> + `Dim_TipPokazatelya`:** фінансист додав у `А_ОтчетБаланс_Свод` вимір
> `ТипПоказателя` (`Перечисление.ВидыСтатейУправленческогоБаланса`),
> заповнюється централізовано в `ПровестиБалансСвод` формулою штатного
> УпрБаланс (АктивПассив→Пассив; «Налоги»→Пассив). Зміни ETL:
> - `refresh_mapping.py` перезапущено (ВидыСтатейУправленческогоБаланса +
>   А_ОтчетБаланс_Свод вже у WHITELIST; `ТипПоказателя`→**`_Fld56131RRef`**
>   у `baserp_storage.json`);
> - `enum_resolver.FROZEN_ENUMS["Перечисление.ВидыСтатейУправленческогоБаланса"]`
>   латиниця→**кирилиця** `["Актив","Пассив","АктивПассив"]` (_EnumOrder
>   0..2) — конвенція як Source/ТипыНалогов (FK Fact==Dim, seed з 1С COM
>   Имя); пуста ссилка enum→"ПустаяСсылка" (як TaxType);
> - `fact_balance.json`: raw_sql += `r._Fld56131RRef AS TipPokazatelya`;
>   `column_to_enum += "TipPokazatelya":"...ВидыСтатейУправленческогоБаланса"`;
>   `column_map += "TipPokazatelya":"TipPokazatelya"`;
> - SQL DDL OlapBASERP: `ALTER TABLE Fact_Balance ADD TipPokazatelya
>   varchar(50)` + **`Dim_TipPokazatelya`** (4: Актив/Пассив/АктивПассив/
>   ПустаяСсылка; seed `seed_dim_tip_pokazatelya.py` з 1С метаданих, як
>   Dim_TaxTypes). ETL 3 періоди. `verify_olap_balance_tippokazatelya.py`
>   **PASS**: TipPokazatelya 0 NULL, FK→Dim 100%, «Налоги»→Пассив,
>   ПОЛНЫЙ БАЛАНС по TipPokazatelya дек 278 093 267,32 / янв
>   288 787 750,11 == штатний звіт.
> PL.pbix: користувач додав таблицю `ТипПоказателя`; зв'язок
> `Fact_Balance[TipPokazatelya]→ТипПоказателя[TipPokazatelya]`
> (Many→One, OneDirection, active) налаштовано через MCP — **видалено
> ложну авто-зв'язку** `ТипПоказателя[EnumOrder]→ТипыНалогов[EnumOrder]`
> (Power BI хибно сматчив `EnumOrder`=0,1,2 двох різних Dim). Потрібен
> Ctrl+S.

**Acceptance** `tests/test_etl_acceptance_balance.py` (PASS): Σ Fact_Balance.Sum_Close
по PAP_Article == ПАП `ОстаткиИОбороты` (січень/ТОВ, виключення 3 груп як Етап4)
до копійки (ОТ tol 1.0); Σ Sum_Close ≈ 0 (Актив=Пасив). Канон-регрес
`tests/test_etl_acceptance_globyno2.py` (PnL) PASS.

### Оновлення 2026-05-17 — `Свод_ДенежныеСредства` LIVE + `dim_warehouses`

**Свод_ДенежныеСредства активна в 1С** (`Документ.А_ФинРез_Баланс`): регістр
`А_ОтчетБаланс_Свод` тепер має +4 ден. Source. ETL **без змін конфігу** —
`pipelines/fact_balance.json` raw_sql фільтрує лише власну орг + період
(Source НЕ фільтрує), тож `--run-once fact_balance --period 2026-01` сам тягне
нові ден. рядки. Заповнюються `Cash_ID` (безнал→`_Fld56124_RRRef`→
Dim_DenezhnyeSredstva), `Individual_ID` (підзвіт→`_Fld56123RRef`→Dim_Individuals).
Січень/ТОВ Fact_Balance: Себест 4775 + ден. ~213; Σ ден. групи Sum_Close=
**75 265 344,95**; «безнал» КО=50 435 887,99==ПАП==УпрБаланс.

**Новий крок `dim_warehouses`** у `pipelines/dim_catalogs.json` (Dim_Warehouses
для Fact_Balance.Warehouse_ID, Себест). `Справочник.Склады` ієрархічний але
**без Кода** → raw_sql recursive-CTE по `_Reference502` БЕЗ `_Code` (паттерн
dim_organizations). DDL `ddl/08_dim_warehouses.sql` + `apply_08_dim_warehouses.py`.
`mapping/refresh_mapping.py` WHITELIST += `Справочник.Склады` (baserp_storage.json
84 об'єкти). Live: 347 рядків, 303 групи, Level1..5. dim_catalogs повний
прогон run_id=261, **61 685** рядків (19 Dim-кроків).

**Запуск:** `apply_08_dim_warehouses.py` → `main.py --run-once dim_catalogs`
→ `main.py --run-once fact_balance --period 2026-01`. Діагностика:
`scripts/probe_olap_balance_state.py`, `scripts/verify_olap_balance_densr.py`,
`scripts/probe_ref502_columns.py`.

> **Урок:** ETL `sql_backend` НЕ дає `_ParentIDRRef`/`_Folder`/`_Code`
> (GetDBStorageStructureInfo пропускає системні) — для ієрархічних Dim
> завжди raw_sql recursive-CTE. Відсутність `_Code` (Длина кода=0) ≠
> неієрархічний — перевіряти Конфігуратор / `_ParentIDRRef`+`_Folder`.

### Оновлення 2026-05-18 — розшифровка `Свод_РасчетыСПартнерами` по Контрагент/Договор/Партнёр

`Свод_РасчетыСПартнерами` (1С) тепер заповнює виміри РС
`Контрагент`/`Партнер`/`Договор` з `РегистрСведений.АналитикаУчетаПоПартнерам`
(на деталях; плуги — ПустаяСсылка). **OLAP конфіг/Dim/mapping НЕ
змінювались** — `pipelines/fact_balance.json` raw_sql вже тягнув
`r._Fld56098RRef AS Counterparty_ID, r._Fld56099RRef AS Partner_ID,
r._Fld56102RRef AS Contract_ID` і `column_map` вже мав ці колонки
(зарезервовано раніше). Раніше колонки були порожні (регістр не
наповнювався) → у Fact_Balance NULL; після доробки 1С + перепроведення
ETL просто перечитує регістр.

`mapping/refresh_mapping.py` **НЕ потрібен** (структура РС не змінювалась —
виміри `Контрагент`/`Партнер`/`Договор` існували в `А_ОтчетБаланс_Свод`
завжди, додано лише дані). **Лекція:** наповнення вже-наявного виміру РС
(не новий реквізит) = `--run-once fact_balance --period <re-posted>` і
все; OLAP нічого не показує, доки ETL не перезапущено ПІСЛЯ
перепроведення документа.

**Запуск (виконано):** `main.py --run-once fact_balance --period 2026-01`
(run_id=290, 7721 рядків; користувач перепровів лише янв2026 — для
дек2025/лют2026 потрібні перепроведення+ETL тих періодів). Dim
(`Dim_Counterparties`/`Dim_Contracts`/`Dim_Partners`) вже засіяні
кроком `dim_catalogs` — перезапуск НЕ потрібен (довідники не змінились).

### Оновлення 2026-05-19 — dim_contracts/ObjektyRaschetov → raw_sql + Dim_TipyDogovorov/FinAgents

**WHITELIST `mapping/refresh_mapping.py` +5 об'єктів:**
`Справочник.А_ФинАгенты`, `Справочник.ИдентификаторыОбъектовМетаданных`,
`Перечисление.ТипыДоговоров`, `Перечисление.ТипыРасчетовСПартнерами`,
`Перечисление.ТипыОбъектовРасчетов`. `baserp_storage.json` перегенерований
(refresh_mapping — строгий WHITELIST; об'єкт не зі списку → нема у mapping).

**FROZEN_ENUMS +3 enum** (`enum_resolver.py`; джерело `scripts/gen_frozen_enums_contracts.py`):
- `ТипыДоговоров` — 11 значень (порядок = \_EnumOrder; фіз. `_Enum1626`)
- `ТипыРасчетовСПартнерами` — 5 значень (`_Enum1684`)
- `ТипыОбъектовРасчетов` — 4 значення (`_Enum1657`)

**`dim_contracts` у `pipelines/dim_catalogs.json` переписано `sql_backend` → `raw_sql`:**
Джерело `_Reference171` (ДоговорыКонтрагентов); LEFT JOIN для денорм. назв:
`_Reference540` (підрозділи), `_Reference360` (підрозділ послуг), `_Reference263` (партнери),
`_Reference529` (контрагенти), `_Reference329` (статті ДДС); `enum_resolver` для TipDogovora;
нові колонки `Is_FinAgent_Contract/TipDogovora/FinAgent_ID/Department_Name/Partner_Name/…`

**`dim_objekty_raschetov` у `pipelines/dim_catalogs.json` переписано `sql_backend` → `raw_sql`:**
Джерело `_Reference319` (ОбъектыРасчетов); `enum_resolver` для `TipRaschetov`/`TipObjektaRaschetov`;
composite UUID (`_Fld...RRef`); `Object_Type_Name` через `ТипСсылки` JOIN `_Reference211`.

**Новий шаг `dim_fin_agents`** у `pipelines/dim_catalogs.json` (`raw_sql`; джерело `_Reference54722`
+ unknown-member; 13 рядків; входить у **default** dim_catalogs — запускається з `python main.py`).

**`Dim_TipyDogovorov` — сид окремо** (не у default pipelines):
`scripts/seed_dim_tipy_dogovorov.py` (1С COM; паттерн Dim_TaxTypes; 12 рядків);
запускається вручну після DDL `scripts/ddl_dim_tipy_dogovorov.sql`.

**Verify** (`scripts/verify_olap_contracts_dims.py`) **PASS:**
Dim_Contracts=8248, Dim_ObjektyRaschetov=14109, Dim_TipyDogovorov=12, Dim_FinAgents=13;
FK 0 orphans TipDogovora→Dim_TipyDogovorov / FinAgent_ID→Dim_FinAgents; enum=кирилиця.
Регрес: Fact_Balance незмінний (`verify_olap_balance_tippokazatelya.py` PASS,
ПОЛНИЙ БАЛАНС дек 278 093 267,32 / янв 288 787 750,11 == штатний звіт).
Повний `python main.py` зелений (dim_catalogs 75990 рядків, fact_* без змін).

**Verify** `scripts/verify_olap_balance_raschety_kontragent.py` **PASS**
(янв2026): розрахункові деталі (SettlementObj_ID NOT NULL, 2287) —
`Counterparty_ID/Contract_ID/Partner_ID` **100% NOT NULL**; плуги
(SettlementObj_ID NULL) — субконто порожні; FK
`Counterparty_ID→Dim_Counterparties` / `Contract_ID→Dim_Contracts` /
`Partner_ID→Dim_Partners` **0 orphans**; повний баланс Σ Sum_Close
(всі Source) = **0,00**, Σ Актив = **288 787 750,11** == штатний звіт
(не змінився); статті КО ЗадолженностьКлиентов 60 888 300,36 /
ПередПоставщиками −131 478 882,01 == регістр/ПАП. Старий
`verify_olap_balance_raschety.py` має застарілі (до-дрейфові) еталони
статей (61 165 524,68 …) — для перевірки субконто/FK використовувати
новий `_kontragent`-скрипт.
