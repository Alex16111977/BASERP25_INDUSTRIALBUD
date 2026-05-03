# 🎯 БОЕВОЙ ПРОМПТ: Реализация Ai_Olap Python ETL Orchestrator

> Скопіюй цей промпт цілком на початок нової сесії Claude Code. Усі знання спроектовані; залишилось реалізувати код.

---

## ROOT CONTEXT — що це за проект

Ти будуєш **Ai_Olap** — Python-оркестратор що щодня переливає таблиці з 1С BAS ERP 2.5 (`BaseERP`) у SQL Server `OlapBASERP` (24 таблиці) для подальшого Power BI dashboard. Це **Шар 2** BI-конвеєра (між Шаром 1 — регістрами сведень в 1С — і Шаром 3 — SQL `OlapBASERP`). Шари 1 і 3 вже готові (Stage 1 і Stage 2 у проекті).

**Робоча директорія:** `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\`

---

## 📚 ОБОВ'ЯЗКОВО ПРОЧИТАЙ ПЕРЕД СТАРТОМ

### База знань про OLAP-проект
**Шлях:** `C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_Olap\`

| Файл | Що там | Коли потрібен |
|---|---|---|
| `KNOWLEDGE_MAP_OLAP.md` | Manifest, потік даних, статус stages | ПЕРШИМ — як вступний огляд |
| `olap_architecture_overview.md` | 4-шарова архітектура, принципи (свёртка в 1С, Source-маркер, drill-down) | Для розуміння WHY |
| `olap_1c_objects.md` | 8 нових об'єктів метаданих 1С: 3 Перечислення (А_ИсточникPL/DDS/А_РазделыCFS), 1 реквізит, 2 РегСв, 2 Документи + frozen Source identifiers | Для extract logic та enum_resolver |
| `olap_obrabotka_provedeniya.md` | BSL логіка А_ФинРез_PL/DDS — 8 CTE + 7 UNION з Source, GROUP BY dedup для DDS | Не потрібно для ETL (свёртка робиться в 1С), але корисно знати що відбувається до того як ми вантажимо |
| `olap_data_sources_erp.md` | Джерельні регістри/документи/справочники 1С — 4 регістри для PnL, 3 для Cashflow, 16 справочників-вимірювань | **КРИТИЧНО** — список того що тягнемо |
| `olap_sql_schema.md` | 24 таблиці OlapBASERP, повний DDL, connection details, схема індексів | **КРИТИЧНО** — куди вантажимо |
| `olap_etl_pipeline.md` | Planned-state Python ETL pipeline (це наш проект) | **КРИТИЧНО** — основа для архітектури |
| `olap_powerbi_model.md` | DAX-міри, сторінки PBIX (Stage 4) | Як reference — щоб розуміти що Power BI чекає на виході |
| `olap_acceptance_etalons.md` | Точні цифри для verification: 🎯 Глобино-2 / ERP_Income = 38 432 968.66 ₴, 3937 рядків PnL, 4652 DDS | **КРИТИЧНО** — acceptance criteria |
| `olap_deviations_from_spec.md` | 7 deviations vs spec v3 final (Source як Dimension, ВидДвижения тип fix, тощо) | Щоб не повторити помилки |

### Existing knowledge (для контексту)
- `_Rarzrabotki/notebook/knowledge/baserp25_knowledge.md` — Golden Rules архітектури BAS ERP 2.5
- `_Rarzrabotki/notebook/knowledge/perenos_dvizheniy_iz_kazny.md` — паттерн `Движения.<Регистр>.Загрузить()` (для розуміння як 1С пише регістри)
- `_Rarzrabotki/notebook/knowledge_PL/pl_methodology.md` — економічна семантика 68 PL-статей

### Research що зроблено заздалегідь
**Файл:** `RESEARCH.md` у поточній папці. Містить:
- Порівняння Apache Airflow / Prefect / Dagster / Kedro / dlt / Bonobo
- Чому **APScheduler + custom modular framework** — найкраще для нашого випадку
- Готові патерни з ETL_Tool (AlisaMitchikov), Tiago Valverde OOP, Aliakbar Hosseinzadeh structure
- Запропонована архітектура (директорія, JSON формат, main.py)

---

## 🎯 ЦІЛЬ ПРОЕКТУ

Створити **production-ready Python проект** який:

1. **Запускається через `main.py`** з 3 режимами:
   - `python main.py --scheduled` — daemon з APScheduler за cron-розкладом
   - `python main.py --run-once <pipeline_id>` — ad-hoc запуск одного pipeline
   - `python main.py --validate` — тільки validate JSON конфіги і exit

2. **Налаштовується через `config.py`** (статичні константи) + `.env` (secrets) + `pipelines/*.json` (динамічна конфігурація що/звідки/куди)

3. **Виконує конкретні задачі для нашого OLAP проекту:**
   - **16 Dim таблиць** з 1С справочників → OlapBASERP.Dim_* (full reload)
   - **1 Bridge** з ТЧ `А_Статьи_PL.Статьи` → OlapBASERP.PLArticle_DDS (full reload)
   - **3 Fact таблиці**:
     - `А_ОтчетPL_Свод` (за вказаний місяць) → Fact_PnL з drill-down URL
     - `А_ОтчетDDS_Свод` → Fact_Cashflow з drill-down URL
     - `.ОстаткиИОбороты` ДенежныеСредства{Безналичные,Наличные} → Fact_CF_Balance
   - **ETL_Runs логування** для кожного запуску

4. **Має повну модульну архітектуру** (extractors/transformers/loaders/orchestrator/config/utils) з типізацією, тестами, structured logging.

5. **Підтримує scheduling** через APScheduler — cron-вирази у JSON `schedule.expression`.

---

## 🚨 КРИТИЧНІ ПРАВИЛА (обов'язково дотримуватися)

### Rule #-1: Python COM тест ПЕРЕД кодом BSL
**Не застосовується тут** — ми не пишемо BSL у цьому проекті. Тільки читаємо вже існуючі регістри 1С через COM.

### Rule #0: Golden Sequence — `list_metadata_objects` → `get_metadata_structure` → `execute_query`
Якщо потрібно перевірити структуру 1С об'єкта (наприклад, чи правильно іменована колонка):
```javascript
mcp__1c-workerp__get_metadata_structure({metaType: "InformationRegisters", name: "А_ОтчетPL_Свод"})
```
Не вигадуй імена полів!

### Rule #4: cp у основну
Якщо працюєш у worktree — після створення файлів `cp` у `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\` після кожного meaningful коміту.

### Frozen identifiers (КРИТИЧНО!)
8 значень `А_ИсточникPL` і 3 значення `А_ИсточникDDS` — це **рядкові литерали** які використовуються у DAX мірах Power BI:
```dax
CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_Income")
```
**Будь-яка зміна імени = breakage Power BI.** При завантаженні Source у Fact-таблицю — використовуй точні імена з `Перечисление.А_ИсточникPL.Метаданные().Имя` (через `enum_resolver`).

### Memory whitelist
Цей проект — **тільки Python код** + **читання даних з 1С**. **Жодних змін у метаданих 1С.** Якщо виникає бажання щось додати/виправити в 1С — STOP and ask. Whitelist у `feedback_no_1c_changes.md` дозволяє ТІЛЬКИ 8 об'єктів які вже створено в Stage 1.

---

## 🏗️ АРХІТЕКТУРА (з RESEARCH.md)

### Stack
- **Python 3.11+**
- **pywin32** (win32com) — COM до 1С
- **pyodbc** + ODBC Driver 17 — SQL Server
- **APScheduler 3.x** — scheduling
- **jsonschema** — JSON config validation
- **structlog** — structured logging
- **pytest + pytest-mock** — тестування
- **python-dotenv** — secrets

### Структура (повна)
```
Ai_Olap/
├── main.py                          ← entry point
├── config.py                        ← static + .env
├── requirements.txt                 ← pinned deps
├── pyproject.toml
├── .env.example                     ← secrets template
├── README.md                        ← quick start + architecture diagram
├── RESEARCH.md                      ← вже є — research summary
├── IMPLEMENTATION_PROMPT.md         ← цей файл
│
├── ai_olap/                         ← core package
│   ├── __init__.py
│   ├── core/
│   │   ├── connections.py
│   │   ├── logging.py
│   │   └── exceptions.py
│   ├── extractors/
│   │   ├── base.py                  ← BaseExtractor (ABC)
│   │   ├── catalog.py
│   │   ├── accumulation_register.py
│   │   ├── information_register.py
│   │   ├── document_tabular.py
│   │   └── factory.py
│   ├── transformers/
│   │   ├── uuid_to_hex.py
│   │   ├── date_normalizer.py
│   │   ├── enum_resolver.py
│   │   ├── drill_down.py
│   │   └── column_mapper.py
│   ├── loaders/
│   │   ├── base.py
│   │   ├── dim_loader.py
│   │   ├── fact_loader.py
│   │   ├── bridge_loader.py
│   │   └── etl_runs.py
│   ├── orchestrator/
│   │   ├── pipeline.py
│   │   ├── runner.py
│   │   └── scheduler.py
│   ├── config/
│   │   ├── loader.py
│   │   ├── schema.py
│   │   └── validator.py
│   └── utils/
│       ├── retries.py
│       └── perf.py
│
├── pipelines/
│   ├── _schema.json                 ← JSON Schema validation
│   ├── dim_catalogs.json            ← 16 Dim + Bridge
│   ├── fact_pnl.json                ← Fact_PnL
│   ├── fact_cashflow.json           ← Fact_Cashflow
│   ├── fact_cf_balance.json         ← Fact_CF_Balance
│   └── all.json                     ← composite
│
├── sql/
│   └── extract_1c/                  ← 1С запити як reference
│
├── tests/
│   ├── conftest.py
│   ├── test_config_loader.py
│   ├── test_extractors/
│   ├── test_transformers/
│   ├── test_loaders/
│   ├── test_pipelines/
│   └── integration/
│       └── test_e2e_dim.py
│
├── logs/                            ← gitignored
│   └── .gitkeep
│
└── docs/
    ├── architecture.md
    └── adding_new_table.md
```

---

## 🛠️ ЕТАПИ РЕАЛІЗАЦІЇ

Виконувати послідовно. Після кожного етапу — тести і **окремий git commit** на гілці `claude/jolly-shtern-233fbd`.

### Етап 0 — Bootstrap (15 хв)
- [ ] `cd C:/Configuration_downloads/BASERP25/.claude/worktrees/jolly-shtern-233fbd/_Rarzrabotki/Olap/Ai_Olap`
- [ ] `pyproject.toml` (для editable install: `pip install -e .`)
- [ ] `requirements.txt`:
  ```
  pywin32==306
  pyodbc==5.0.1
  APScheduler==3.10.4
  jsonschema==4.21.0
  structlog==24.1.0
  python-dotenv==1.0.0
  pytest==8.0.0
  pytest-mock==3.12.0
  ```
- [ ] `.env.example` з placeholder-ами (НЕ commit реальний `.env`)
- [ ] `.gitignore`: `__pycache__/`, `*.pyc`, `logs/*`, `.env`, `*.egg-info/`, `.pytest_cache/`
- [ ] `README.md` (1-page quick start)
- [ ] `config.py` (per RESEARCH.md приклад)
- [ ] Створити пусті `__init__.py` у всіх папках `ai_olap/*/`

**Verify:** `python -c "import ai_olap"` без помилок.

**Commit:** `Ai_Olap: bootstrap project structure (config.py, requirements, package init)`

### Етап 1 — Core layer (1 година)
- [ ] `ai_olap/core/exceptions.py` — `AiOlapError`, `ExtractError`, `TransformError`, `LoadError`, `ConfigError`
- [ ] `ai_olap/core/logging.py` — structlog setup з JSON output до `logs/<date>.log`
- [ ] `ai_olap/core/connections.py`:
  - `get_erp_connection()` — V83.COMConnector singleton
  - `get_olap_connection()` — pyodbc з context manager
  - `bulk_insert(conn, table, columns, rows)` — fast_executemany
  - `execute_sql(conn, sql, params=None)` — single-statement execute
- [ ] `ai_olap/utils/retries.py` — `@retry(attempts=3, backoff=5)` decorator
- [ ] `ai_olap/utils/perf.py` — `@timed("step_name")` decorator → log duration

**Tests:** `tests/test_core/`:
- `test_connections.py` — mock COM, mock pyodbc, перевірити що connections відкриваються/закриваються
- `test_logging.py` — перевірити що log JSON містить очікувані поля

**Commit:** `Ai_Olap: core layer (connections, logging, exceptions, utils)`

### Етап 2 — Config layer (45 хв)
- [ ] `ai_olap/config/schema.py` — JSON Schema definitions для:
  - `pipeline.json` schema (with $ref до `extractor`, `transformer`, `loader` schemas)
  - Окремі schemas для кожного extractor type (catalog/accumulation_register/information_register/document_tabular)
- [ ] `pipelines/_schema.json` — full JSON Schema файл (для IDE auto-complete і validation)
- [ ] `ai_olap/config/loader.py`:
  - `load_pipelines() -> list[PipelineConfig]` — читає `pipelines/*.json`, парсить, повертає dataclass-list
  - `PipelineConfig` dataclass з полями name/schedule/steps/transactional/parameters
- [ ] `ai_olap/config/validator.py`:
  - `validate_pipeline(config: dict) -> bool` — через jsonschema
  - `validate_all(pipelines: list) -> None` — викликає validate_pipeline для кожного, raises ConfigError якщо хоч один зламаний
  - **Cross-checks:** перевірити що `target_table` існує в OlapBASERP (опційно, через INFORMATION_SCHEMA)

**Tests:** `tests/test_config/`:
- `test_loader.py` — load valid/invalid configs, перевірити структуру
- `test_validator.py` — invalid schemas → ConfigError; valid → silent

**Commit:** `Ai_Olap: config layer (loader, validator, JSON Schema)`

### Етап 3 — Extractors layer (2 години)
- [ ] `ai_olap/extractors/base.py`:
  ```python
  class BaseExtractor(ABC):
      def __init__(self, erp_conn, config: dict):
          self.erp = erp_conn
          self.config = config
      
      @abstractmethod
      def extract(self, params: dict = None) -> Iterator[dict]:
          """Yield rows as dicts."""
  ```
- [ ] `ai_olap/extractors/catalog.py` — `CatalogExtractor` для справочників (виконує запит, ітерує `Запрос.Выполнить().Выгрузить()`)
- [ ] `ai_olap/extractors/accumulation_register.py` — для регістрів накопичення (.Обороты, .ОстаткиИОбороты)
- [ ] `ai_olap/extractors/information_register.py` — для регістрів сведень (читає за фільтром `Регистратор.Месяц`)
- [ ] `ai_olap/extractors/document_tabular.py` — для ТЧ документів (наприклад `А_Статьи_PL.Статьи` для Bridge)
- [ ] `ai_olap/extractors/factory.py`:
  ```python
  def get_extractor(extractor_type: str, erp_conn, config: dict) -> BaseExtractor:
      EXTRACTOR_MAP = {
          "catalog": CatalogExtractor,
          "accumulation_register": AccumulationRegisterExtractor,
          "information_register": InformationRegisterExtractor,
          "document_tabular": DocumentTabularExtractor,
      }
      cls = EXTRACTOR_MAP.get(extractor_type)
      if not cls:
          raise ConfigError(f"Unknown extractor type: {extractor_type}")
      return cls(erp_conn, config)
  ```

**Tests:** `tests/test_extractors/`:
- Mock COM connection (`mock.MagicMock`), повернути fake `Запрос.Выполнить().Выгрузить()` ТЗ
- Перевірити що CatalogExtractor для `Справочник.Организации` повертає очікувані dicts

**Commit:** `Ai_Olap: extractors layer (catalog, registers, tabular, factory)`

### Етап 4 — Transformers layer (1 година)
- [ ] `ai_olap/transformers/uuid_to_hex.py` — `transform(rows, fields=[...])` — конвертує COM UUID → 32-char hex
- [ ] `ai_olap/transformers/date_normalizer.py` — `transform(rows, fields=[...])` — 1С Date → datetime для SQL
- [ ] `ai_olap/transformers/enum_resolver.py` — `transform(rows, field, enum_metadata)` — EnumRef → name string (для frozen identifiers!)
- [ ] `ai_olap/transformers/drill_down.py` — `transform(rows, field, outputs=[...])` — формує:
  - `Source_Recorder_ID` (UUID hex)
  - `Source_Recorder_Type` (1C metadata name)
  - `Source_Recorder_Url` (`e1cib/data/Документ.X?ref=UUID`)
  - `Source_Recorder_Presentation` (str(recorder))
- [ ] `ai_olap/transformers/column_mapper.py` — `transform(rows, mapping={1c_name: sql_name})`

**Tests:** для кожного transformer окремий test з мок-даними.

**Commit:** `Ai_Olap: transformers layer (uuid, date, enum, drill_down, column_mapper)`

### Етап 5 — Loaders layer (1.5 години)
- [ ] `ai_olap/loaders/base.py`:
  ```python
  class BaseLoader(ABC):
      def __init__(self, olap_conn, config: dict):
          self.conn = olap_conn
          self.config = config
      
      @abstractmethod
      def load(self, rows: Iterable[dict], params: dict = None) -> int:
          """Insert rows, return count."""
  ```
- [ ] `ai_olap/loaders/dim_loader.py` — `DimLoader`: DELETE FROM table → bulk_insert (full reload)
- [ ] `ai_olap/loaders/fact_loader.py` — `FactLoader`: DELETE WHERE Period_Month=? → bulk_insert (idempotent per period)
- [ ] `ai_olap/loaders/bridge_loader.py` — `BridgeLoader`: same as DimLoader але без Marked_For_Deletion
- [ ] `ai_olap/loaders/etl_runs.py`:
  - `log_run_start(conn, script, period_month=None) -> int`
  - `log_run_success(conn, run_id, rows_loaded)`
  - `log_run_failure(conn, run_id, error_text)`

**Tests:** mock pyodbc, перевірити SQL queries (DELETE+INSERT pattern), перевірити ETL_Runs row insertion.

**Commit:** `Ai_Olap: loaders layer (dim, fact, bridge, etl_runs)`

### Етап 6 — Orchestrator layer (2 години)
- [ ] `ai_olap/orchestrator/pipeline.py`:
  ```python
  class Pipeline:
      def __init__(self, config: PipelineConfig):
          self.config = config
      
      def run(self, params: dict = None) -> PipelineResult:
          """Виконує всі steps послідовно. Логує у ETL_Runs."""
  ```
- [ ] `ai_olap/orchestrator/runner.py`:
  ```python
  class PipelineRunner:
      def __init__(self, pipelines: list[PipelineConfig]):
          self.pipelines = {p.pipeline_id: p for p in pipelines}
      
      def run(self, pipeline_id: str, params: dict = None):
          pipeline = self.pipelines[pipeline_id]
          return Pipeline(pipeline).run(params)
  ```
- [ ] `ai_olap/orchestrator/scheduler.py`:
  ```python
  def build_scheduler(runner: PipelineRunner, pipelines: list) -> BlockingScheduler:
      scheduler = BlockingScheduler(timezone=SCHEDULER_TIMEZONE,
                                     job_defaults=SCHEDULER_JOB_DEFAULTS)
      for p in pipelines:
          if p.schedule.type == "cron":
              cron = CronTrigger.from_crontab(p.schedule.expression)
              scheduler.add_job(runner.run, cron, args=[p.pipeline_id],
                                 id=p.pipeline_id, name=p.name)
      return scheduler
  ```

**Tests:** `tests/test_orchestrator/`:
- Перевірити що Pipeline виконує steps у правильному порядку
- Перевірити що PipelineRunner кидає помилку якщо невідомий pipeline_id
- Перевірити що build_scheduler регіструє правильну кількість jobs

**Commit:** `Ai_Olap: orchestrator layer (pipeline, runner, scheduler)`

### Етап 7 — Pipeline JSON-конфіги (1 година)
Створи 5 JSON файлів у `pipelines/`:

- [ ] `dim_catalogs.json` — 17 steps (16 Dim + 1 Bridge):
  - Dim_Organizations, Dim_Departments, Dim_Directions, Dim_Counterparties, Dim_Contracts, Dim_Items, Dim_ItemGroups, Dim_Individuals, Dim_Users, Dim_BankAccounts, Dim_Currencies, Dim_DDS_Articles, Dim_Expense_Articles, Dim_Income_Articles, Dim_PL_Articles, Dim_PL_ArticleGroups, PLArticle_DDS (Bridge)
  - Schedule: `0 1 * * *` (01:00 щодня)
  - Кожен step з extractor + transformers + loader (per RESEARCH.md приклад)
  - **Спецвипадки:**
    - Dim_DDS_Articles: + enum_resolver для `А_РазделCFS` → CFS_Section
    - Dim_BankAccounts: extractor з UNION 2 справочників (БанковскиеСчетаОрганизаций + Кассы) + Account_Type literal
    - PLArticle_DDS: document_tabular extractor по `А_Статьи_PL.Статьи`
- [ ] `fact_pnl.json` — 1 step:
  - Schedule: `0 2 * * *` (02:00 щодня — після Dim)
  - Parameters: `period_month` (default `last_completed`)
  - Extractor: information_register `А_ОтчетPL_Свод`, фільтр `Регистратор.Месяц = &Месяц И Регистратор.Проведен`
  - Transformers: uuid_to_hex (для всіх UUID полів), drill_down (`Документ_Источник` → 4 колонки), enum_resolver (Source), column_mapper
  - Loader: fact_period (Fact_PnL, Period_Month)
- [ ] `fact_cashflow.json` — аналогічно fact_pnl.json але для `А_ОтчетDDS_Свод` → Fact_Cashflow
- [ ] `fact_cf_balance.json` — extractor accumulation_register з .ОстаткиИОбороты, окремо для ДенежныеСредстваБезналичные та ДенежныеСредстваНаличные → Fact_CF_Balance
- [ ] `all.json` — composite: запускає dim_catalogs → fact_pnl → fact_cashflow → fact_cf_balance послідовно (для manual full refresh)
- [ ] `_schema.json` — JSON Schema для всіх вище

**Verify:** `python main.py --validate` → всі configs valid.

**Commit:** `Ai_Olap: 5 pipeline JSON configs (dim, fact_pnl, fact_cashflow, balance, all) + schema`

### Етап 8 — main.py (30 хв)
- [ ] `main.py` per RESEARCH.md приклад:
  - argparse з 3 mutually-exclusive modes (`--scheduled` / `--run-once` / `--validate`)
  - Optional `--period YYYY-MM` для fact pipelines

**Verify:**
```bash
python main.py --validate
# Очікується: "All N pipelines valid."

python main.py --run-once dim_catalogs
# Очікується: запуск всіх 17 steps, всі Dim_* tables заповнені
```

**Commit:** `Ai_Olap: main.py entry point with --scheduled/--run-once/--validate`

### Етап 9 — End-to-end test з реальними даними (1 година)
- [ ] `tests/integration/test_e2e_dim.py`:
  - Запустити `dim_catalogs` pipeline
  - Перевірити: всі 17 таблиць Dim_* + PLArticle_DDS заповнені (count > 0)
  - Перевірити: ETL_Runs має 17 successful runs
- [ ] `tests/integration/test_e2e_fact_pnl.py`:
  - Запустити `fact_pnl` для `period_month='2026-02'`
  - Acceptance criteria (з `olap_acceptance_etalons.md`):
    ```sql
    -- 1. Σ рядків
    SELECT COUNT(*) FROM Fact_PnL WHERE Period_Month='2026-02-01'
    -- Очікується: ~3937
    
    -- 2. Σ Plan
    SELECT SUM(Sum_Plan_Grn) FROM Fact_PnL 
    WHERE Source='PL_Excel' AND Period_Month='2026-02-01'
    -- Очікується: 115 728 517.43 ₴
    
    -- 3. 🎯 Глобино-2 / ERP_Income
    SELECT SUM(F.Sum_ERP_Grn)
    FROM Fact_PnL F
    JOIN Dim_Departments D ON D.Department_ID = F.Department_ID
    WHERE F.Source='ERP_Income' 
      AND F.Period_Month='2026-02-01'
      AND D.Department_Name='Глобино-2'
    -- Очікується: 38 432 968.66 ₴ ± 0.01
    
    -- 4. Drill-down URL валідні
    SELECT TOP 100 Source_Recorder_Url FROM Fact_PnL
    -- Усі починаються з 'e1cib/data/Документ.'
    ```
- [ ] `tests/integration/test_e2e_fact_cashflow.py`:
  ```sql
  SELECT Source, COUNT(*), SUM(Sum_Grn) FROM Fact_Cashflow 
  WHERE Period_Month='2026-02-01' GROUP BY Source
  -- Очікується: ERP_Безнал ~1828 / 5 009 938 998.60 ₴
  --             ERP_Нал ~351 / 40 090 065.48 ₴
  --             Казна ~2473 / 5 261 822 347.34 ₴
  ```

**Run:** `pytest tests/integration/ -v`

**Commit:** `Ai_Olap: integration tests + acceptance verified for лютий 2026`

### Етап 10 — README + docs (30 хв)
- [ ] `README.md`:
  - Що це за проект (1 параграф)
  - Quick start (install + .env + python main.py --scheduled)
  - Архітектурна діаграма (ASCII art з 4 шарів)
  - Cross-link на `knowledge_Olap/`
- [ ] `docs/architecture.md` — детальніший опис, link на RESEARCH.md
- [ ] `docs/adding_new_table.md` — how-to для майбутніх розробників (як додати Dim_NewCatalog без правки коду — тільки JSON)

**Commit:** `Ai_Olap: README + docs/`

---

## ✅ ACCEPTANCE CRITERIA (фінал)

Після завершення всіх 10 етапів — звіт користувачу з підтвердженням:

### Структурні
- [ ] `Ai_Olap/` фолдер створено у двох локаціях (worktree + main config) і ідентичні (Rule #4)
- [ ] Структура відповідає RESEARCH.md (всі папки + файли наявні)
- [ ] `pip install -e .` встановлюється без помилок
- [ ] `python -c "import ai_olap"` без помилок

### Functional
- [ ] `python main.py --validate` → "All 5 pipelines valid"
- [ ] `python main.py --run-once dim_catalogs` → 17 таблиць заповнено, ETL_Runs має 17 OK runs
- [ ] `python main.py --run-once fact_pnl --period 2026-02` → Fact_PnL за лютий заповнено
- [ ] `python main.py --scheduled` → APScheduler стартує, показує наступні запуски (verify через `Ctrl+C` після того як побачив log)

### Acceptance numerical (з knowledge_Olap/olap_acceptance_etalons.md)
- [ ] 🎯 **Глобино-2 / Source=ERP_Income / Sum_ERP_Grn = 38 432 968.66 ₴** (точно ± 0.01)
- [ ] **Σ Sum_Plan_Grn (Source=PL_Excel) = 115 728 517.43 ₴** (точно)
- [ ] **Fact_PnL за лютий = ~3937 рядків**
- [ ] **Fact_Cashflow за лютий = ~4652 рядки** (3 Source значення)
- [ ] **ERP_Безнал у Fact_Cashflow ≈ Σ ДенежныеСредстваБезналичные.Обороты** (Δ < 0.5%)

### Tests
- [ ] `pytest tests/` — всі unit tests проходять (≥80% coverage)
- [ ] `pytest tests/integration/` — всі e2e tests проходять

### Logs
- [ ] `logs/<date>.json` містить structured log entries з усіх pipelines
- [ ] ETL_Runs табличка в OlapBASERP має повну історію запусків

### Git
- [ ] 11 commits на гілці `claude/jolly-shtern-233fbd` (по одному на етап + final)
- [ ] Кожен commit message має формат `Ai_Olap: <зміна>`

---

## 🛠️ КОРИСНІ КОМАНДИ

### Setup
```bash
cd C:/Configuration_downloads/BASERP25/.claude/worktrees/jolly-shtern-233fbd/_Rarzrabotki/Olap/Ai_Olap
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
pip install -e .
cp .env.example .env
# Заповни .env реальними credentials
```

### Daily run
```bash
python main.py --scheduled
# Або для cron-job у Windows Task Scheduler:
python C:/path/to/Ai_Olap/main.py --run-once all
```

### Тестування з реальними даними (read-only)
```bash
python main.py --run-once dim_catalogs       # завантажує всі Dim
python main.py --run-once fact_pnl --period 2026-02  # ad-hoc PnL
```

### Verification queries (через sqlcmd або pyodbc)
```sql
-- Глобино-2 еталон
SELECT SUM(F.Sum_ERP_Grn)
FROM Fact_PnL F
JOIN Dim_Departments D ON D.Department_ID = F.Department_ID
WHERE F.Source='ERP_Income' 
  AND F.Period_Month='2026-02-01'
  AND D.Department_Name='Глобино-2';
-- Має бути 38 432 968.66
```

---

## 📋 КРИТИЧНІ ФАЙЛИ ДЛЯ REFERENCE

### Code patterns
- `_Rarzrabotki/Python/Olap/ddl/*.sql` — DDL OlapBASERP таблиць (Stage 2 готовий)
- `_Rarzrabotki/Python/Olap/tests/test_olap_connectivity.py` — pyodbc connection pattern
- `_Rarzrabotki/Python/test/test_finrez_pl_query_2026-05-01.py` — Python COM read pattern
- `_Rarzrabotki/Python/PnL/scripts/utils/com_connect.py` — `connect_erp()` helper приклад
- `_Rarzrabotki/Python/PnL/scripts/test/test_vyruchka_parity.py` — еталон Глобино-2 = 38 432 968.66

### Knowledge base
- `_Rarzrabotki/notebook/knowledge_Olap/` — повний опис проекту (10 файлів)
- `_Rarzrabotki/notebook/knowledge/baserp25_knowledge.md` — Golden Rules

### Memory
- `feedback_no_1c_changes.md` — whitelist 8 1С-об'єктів (тут не дотримуємося — ми тільки read)

---

## 🚀 СТАРТОВА ПОСЛІДОВНІСТЬ

**Шаг 1.** Прочитати `RESEARCH.md` (вже у поточній папці).

**Шаг 2.** Прочитати `_Rarzrabotki/notebook/knowledge_Olap/KNOWLEDGE_MAP_OLAP.md` + `olap_etl_pipeline.md` + `olap_sql_schema.md` + `olap_acceptance_etalons.md`.

**Шаг 3.** Створити TodoWrite зі списком 10 етапів.

**Шаг 4.** Почати з Етапу 0 (bootstrap).

**Шаг 5.** Після кожного етапу — git commit + перехід до наступного.

**Шаг 6.** Етап 9 — головний (acceptance з реальними даними) — підтверджує що pipeline працює.

**Шаг 7.** Фінальний звіт користувачу з усіма acceptance criteria.

---

## ⚠️ ОСТАНОВИТИСЬ І СПРОСИТИ якщо

- Знайдено архітектурне обмеження що вимагає додавання нового extractor типу (наприклад, ChartOfAccounts для майбутньої інтеграції з регістром Хозрасчетный) — додавати треба, але через RESEARCH.md, не імпровізувати.
- Acceptance тест Глобино-2 повертає число яке відхиляється > 0.01 від еталона — означає є баг у transformer logic. STOP і дебаг.
- pyodbc fast_executemany падає на Cyrillic — треба перейти на pandas DataFrame + to_sql, обговорити.
- APScheduler у --scheduled mode не може запуститись через права на Windows — обговорити Task Scheduler як альтернативу.
- JSON config має невідомий extractor type — STOP і фактично додати його, не silently fail.

---

## 🎁 ФІНАЛ

Після завершення всіх 10 етапів і acceptance criteria — звіт користувачу:

```
🎯 Ai_Olap PYTHON ETL ORCHESTRATOR — DONE

[Структура]
✓ 60+ файлів у Ai_Olap/ (worktree + main config, ідентичні)
✓ Modular architecture: core/extractors/transformers/loaders/orchestrator/config/utils
✓ 5 pipelines у JSON конфігах
✓ pip install -e . OK

[Functional]
✓ python main.py --scheduled — APScheduler daemon працює
✓ python main.py --run-once dim_catalogs — 17 Dim+Bridge таблиць заповнено
✓ python main.py --run-once fact_pnl --period 2026-02 — Fact_PnL заповнено
✓ python main.py --validate — всі 5 pipelines valid

[Acceptance — лютий 2026]
✓ 🎯 Глобино-2 / ERP_Income / Sum_ERP_Grn = 38 432 968.66 ₴ (точно)
✓ Σ Plan = 115 728 517.43 ₴ (точно)
✓ Fact_PnL = 3937 рядків
✓ Fact_Cashflow = 4652 рядки (3 Source)
✓ ERP_Безнал ≈ Обороты безналу (Δ < 0.5%)

[Tests]
✓ pytest tests/ — N tests passed (≥80% coverage)
✓ pytest tests/integration/ — все e2e OK

[Git]
✓ 11 commits на claude/jolly-shtern-233fbd

[Готово до Stage 4 — Power BI PBIX]
Чекаю OK Alex.
```

---

**Готовий?** Виконуй Шаг 1 → запам'ятай acceptance numbers → почни з Етапу 0.
