# Research: Python ETL Orchestrator для 1С BaseERP → OlapBASERP

**Дата дослідження:** 2026-05-03
**Мета:** знайти best-of-breed архітектуру для production-ready Python проекту що оркеструє вивантаження таблиць з 1С BAS ERP у SQL OlapBASERP за розкладом, керується JSON-конфігом, модульний.

---

## Розглянуті варіанти (з порівнянням)

### A. Heavy frameworks

#### Apache Airflow
- **Pro:** standard de facto, web UI, retries, monitoring, DAG visualization
- **Con:** overhead — потребує scheduler+webserver+metadb (Postgres), DAG-as-code separate from logic, складна установка на Windows Server, надлишковий для 1 інстансу
- **Verdict:** ❌ overkill для нашого випадку

#### Prefect
- **Pro:** modern Python-native, no DAG ceremony, type-safe, dynamic flows
- **Con:** SaaS-orientated, потребує Prefect server / cloud agent, складно повністю on-premise
- **Verdict:** ❌ якщо хочемо стояти на своїй машині — overkill

#### Dagster
- **Pro:** asset-aware, software engineering mindset, type-checking
- **Con:** вимагає Dagit web UI server, learning curve
- **Verdict:** ❌ overkill

#### Kedro
- **Pro:** project template, YAML-config, modular pipelines, nodes/catalog/runner abstractions, separation of concerns
- **Con:** потребує Kedro CLI, специфічна термінологія (nodes/catalog), DataCatalog pattern не підтримує custom 1С COM connector нативно
- **Verdict:** ⚠️ Можна, але треба писати custom DataSet adapter для COM. Більше ceremonia ніж потрібно.

### B. Modern data tools

#### dlt-hub/dlt (data load tool)
- **Pro:** Python-native, schema inference, incremental loading, MS SQL destination support, лінива конфігурація
- **Con:** оптимізований для REST APIs/SQL/cloud sources; **немає built-in 1С COM connector**, треба писати custom source. Schema inference може ускладнити мапінг до існуючих 24 OlapBASERP таблиць (вони вже визначені, не infer)
- **Verdict:** ⚠️ Підходить для 1С через custom source, але для нашого fixed-schema випадку — overhead. Краще для greenfield ELT.

#### Bonobo
- **Pro:** простий, pythonic, графовий ETL
- **Con:** низький activity, мало maintenance, не підтримує scheduling нативно
- **Verdict:** ❌ застарілий

#### Singer.io / Meltano
- **Pro:** standard taps + targets, можна писати custom tap для 1С
- **Con:** great для REST API/JSON sources, не для COM. Складна архітектура поверх простої задачі
- **Verdict:** ❌ overhead

### C. Lightweight building blocks (РЕКОМЕНДОВАНО)

#### APScheduler (scheduler) + custom modular framework
- **Pro:** APScheduler — пакет для cron/interval scheduling, мінімальний, працює з будь-яким Python кодом, persistent jobstore (опційно), ніяких зайвих процесів
- **Architecture pattern:** combine APScheduler + custom JSON-driven extractors/loaders + APScheduler у `main.py`
- **Verdict:** ✅ **РЕКОМЕНДОВАНО** — найкращий fit для нашого case (1 машина, fixed schema, custom 1С COM, повний контроль)

---

## Найближчі знайдені референси на GitHub

### 1. AlisaMitchikov/ETL_Tool — найближчий patterns
**URL:** https://github.com/AlisaMitchikov/ETL_Tool

**Що цінне:**
- Python ETL для multiple sources (SQL Server, MySQL, CSV, API) → SQL Server data warehouse
- Selective load з queries (delta vs full)
- ODS staging pattern
- `app.py` orchestration + `functions.py` shared utilities
- Per-table ETL files (зручно для конкретних tables як у нас 16 Dim + 3 Fact)
- Tests з success/failure scenarios

**Adaptable for our case:** Додати COM extractor module поряд з SQL/MySQL/CSV — той самий patterns.

### 2. Tiago Valverde "Building a Basic ETL Pipeline in Python with OOP"
**URL:** https://www.tiagovalverde.com/posts/building-a-basic-etl-pipeline-in-python-with-oop

**Що цінне:**
- BaseExtractor / BaseTransformer / BaseLoader abstract classes
- OOP separation of concerns
- Type hints throughout
- Class-based pipeline composition

**Adaptable:** наша архітектура успадкує OOP базовий patterns.

### 3. Aliakbar Hosseinzadeh "ETL/ELT Pipeline Project Structure"
**URL:** https://medium.com/@aliakbarhosseinzadeh/structuring-an-etl-pipeline-project-best-practices-5ed1e4d5a601

**Що цінне:**
- Structure: `src/`, `tests/`, `data/`, `docs/`, `scripts/`, `sql/`, `Docker/`, `requirements/`
- `src/main.py` як orchestration entry point
- `src/config/settings.py` з .env
- `src/config/logging_config.py` для централізованого логування
- Окремі extract/transform/load/quality модулі
- Factory pattern у `db_connection/builder.py`
- Mirror tests/

**Adaptable:** наша структура — це адаптація цього patterns.

### 4. dlt-hub/dlt
**URL:** https://github.com/dlt-hub/dlt

**Що цінне:**
- Schema inference, incremental loading patterns
- Pipeline + Source + Destination абстракції
- Code patterns для bulk_insert через pyodbc

**Adaptable:** поки overhead, але якщо у майбутньому треба буде cloud destinations — можна migrate.

### 5. agronholm/apscheduler
**URL:** https://github.com/agronholm/apscheduler

**Що цінне:**
- `BlockingScheduler` для standalone скриптів (наш case)
- `BackgroundScheduler` якщо інтегруємо в Flask/web
- Cron + Interval triggers
- Persistent jobstore через SQLAlchemy (на майбутнє)

**Use directly.**

---

## Рекомендована архітектура для Ai_Olap

### Stack
- **Python 3.11+** (вимога Windows + COM)
- **pywin32 (win32com)** — COM connector до 1С BaseERP
- **pyodbc + ODBC Driver 17 for SQL Server** — connector до OlapBASERP
- **APScheduler 3.x** — cron-style scheduling
- **jsonschema** — validation JSON-конфігів pipelines
- **structlog** — structured logging (JSON output для grep/parsing)
- **pytest + pytest-mock** — тестування
- **python-dotenv** — secrets з `.env`

### Розмір залежностей
~30 MB total (pywin32 ~20, pyodbc ~5, інше ~5). Усе pure-Python окрім pywin32.

### Архітектурні принципи
1. **JSON-driven configuration**: `pipelines/*.json` визначають що, звідки, куди тягнути. Жодних hardcoded SQL чи table names у Python коді.
2. **Strict separation Extract/Transform/Load**: 3 окремих папки з абстрактними базовими класами.
3. **Plugin-style extractors**: для нових типів об'єктів (РегістрБухгалтерії, тощо у майбутньому) — додати новий клас, не змінювати orchestrator.
4. **Idempotent loaders**: DELETE+INSERT за `Period_Month` для Fact, full reload для Dim.
5. **Centralized logging**: всі компоненти пишуть у `logs/<date>_<pipeline>.json` через structlog.
6. **ETL_Runs табличка** в OlapBASERP — single source of truth для статусу запусків.
7. **`main.py` як entry point** з 2 режимами: `--scheduled` (APScheduler daemon) і `--run-once <pipeline>` (ad-hoc).
8. **JSON config validation на старті** через jsonschema → fail fast якщо помилка в конфігу.

### Запропонована структура

```
Ai_Olap/
├── main.py                          # entry point — daemon або ad-hoc
├── config.py                        # static: connection strings, paths, defaults
├── requirements.txt                 # pinned deps
├── pyproject.toml                   # [optional] для pip install -e .
├── .env.example                     # template для secrets
├── README.md                        # quick start
├── RESEARCH.md                      # цей файл
│
├── ai_olap/                         # core package
│   ├── __init__.py
│   │
│   ├── core/                        # cross-cutting
│   │   ├── __init__.py
│   │   ├── connections.py           # COMConnectionPool, SQLConnectionPool
│   │   ├── logging.py               # structlog setup
│   │   └── exceptions.py            # AiOlapError, ExtractError, LoadError
│   │
│   ├── extractors/                  # 1С COM readers
│   │   ├── __init__.py
│   │   ├── base.py                  # BaseExtractor (ABC)
│   │   ├── catalog.py               # CatalogExtractor — для справочників
│   │   ├── accumulation_register.py # AccumulationRegisterExtractor
│   │   ├── information_register.py  # InformationRegisterExtractor
│   │   ├── document_tabular.py      # DocumentTabularExtractor — для ТЧ
│   │   └── factory.py               # get_extractor(type) dispatch
│   │
│   ├── transformers/                # data shaping
│   │   ├── __init__.py
│   │   ├── uuid_to_hex.py           # COM UUID → char(32)
│   │   ├── date_normalizer.py       # 1С Date → SQL Server datetime2
│   │   ├── enum_resolver.py         # EnumRef.X → Section_Code string
│   │   ├── drill_down.py            # формує Source_Recorder_Url
│   │   └── column_mapper.py         # 1С field name → SQL column name
│   │
│   ├── loaders/                     # SQL writers
│   │   ├── __init__.py
│   │   ├── base.py                  # BaseLoader (ABC)
│   │   ├── dim_loader.py            # DimLoader — full reload
│   │   ├── fact_loader.py           # FactLoader — DELETE+INSERT за period
│   │   ├── bridge_loader.py         # BridgeLoader — full reload N:M
│   │   └── etl_runs.py              # log_run_start/success/failure
│   │
│   ├── orchestrator/                # координація
│   │   ├── __init__.py
│   │   ├── pipeline.py              # Pipeline class — runs config steps
│   │   ├── runner.py                # PipelineRunner — invoke per pipeline
│   │   └── scheduler.py             # APScheduler config + job registration
│   │
│   ├── config/                      # config layer
│   │   ├── __init__.py
│   │   ├── loader.py                # load_pipelines() → list[PipelineConfig]
│   │   ├── schema.py                # jsonschema definitions
│   │   └── validator.py             # validate_config()
│   │
│   └── utils/                       # допоміжне
│       ├── __init__.py
│       ├── retries.py               # retry decorators (для ETL fails)
│       └── perf.py                  # timing decorators
│
├── pipelines/                       # JSON pipeline definitions
│   ├── _schema.json                 # JSON Schema для validation
│   ├── dim_catalogs.json            # 16 Dim таблиць як один pipeline
│   ├── fact_pnl.json                # Fact_PnL з drill-down
│   ├── fact_cashflow.json           # Fact_Cashflow
│   ├── fact_cf_balance.json         # Fact_CF_Balance
│   └── all.json                     # composite — викликає всі вище
│
├── sql/                             # standalone SQL/BSL queries (як reference)
│   ├── extract_1c/                  # 1С запити
│   │   ├── dim_organizations.bsl
│   │   ├── fact_pnl.bsl
│   │   └── ...
│   └── load/                        # SQL Server допоміжні
│       └── post_load_indexes.sql
│
├── tests/                           # pytest
│   ├── __init__.py
│   ├── conftest.py                  # fixtures (mock COM, mock SQL)
│   ├── test_config_loader.py
│   ├── test_extractors/
│   ├── test_transformers/
│   ├── test_loaders/
│   ├── test_pipelines/
│   └── integration/                 # end-to-end з реальною ERP/Olap
│       └── test_e2e_dim.py
│
├── logs/                            # gitignored — runtime logs
│   └── .gitkeep
│
└── docs/                            # opcional
    ├── architecture.md              # пояснення structure (link to knowledge_Olap)
    └── adding_new_table.md          # how-to для розробника
```

### Pipeline JSON формат

Приклад `pipelines/dim_catalogs.json`:

```json
{
  "$schema": "./_schema.json",
  "pipeline_id": "dim_catalogs",
  "name": "Dim Catalogs Pipeline",
  "description": "Завантажує всі 16 Dim таблиць + 1 Bridge з 1С BaseERP в OlapBASERP",
  "schedule": {
    "type": "cron",
    "expression": "0 1 * * *",
    "description": "Щодня о 01:00 (перед Fact pipelines о 02:00)"
  },
  "transactional": false,
  "steps": [
    {
      "step_id": "dim_organizations",
      "extractor": {
        "type": "catalog",
        "object": "Справочник.Организации",
        "query": "ВЫБРАТЬ Ссылка, Код, Наименование, ПометкаУдаления ИЗ Справочник.Организации",
        "params": {}
      },
      "transformers": [
        {"type": "uuid_to_hex", "fields": ["Ссылка"]},
        {"type": "column_mapper", "mapping": {
          "Ссылка": "Organization_ID",
          "Код": "Organization_Code",
          "Наименование": "Organization_Name",
          "ПометкаУдаления": "Marked_For_Deletion"
        }}
      ],
      "loader": {
        "type": "dim_full_reload",
        "target_table": "Dim_Organizations",
        "columns": ["Organization_ID", "Organization_Code", "Organization_Name", "Marked_For_Deletion"]
      }
    },
    {
      "step_id": "dim_dds_articles",
      "extractor": {
        "type": "catalog",
        "object": "Справочник.СтатьиДвиженияДенежныхСредств",
        "query": "ВЫБРАТЬ Ссылка, Код, Наименование, А_РазделCFS, ПометкаУдаления, ЭтоГруппа, Родитель ИЗ Справочник.СтатьиДвиженияДенежныхСредств"
      },
      "transformers": [
        {"type": "uuid_to_hex", "fields": ["Ссылка", "Родитель"]},
        {"type": "enum_resolver", "field": "А_РазделCFS", "enum": "Перечисление.А_РазделыCFS"},
        {"type": "column_mapper", "mapping": {
          "Ссылка": "DDS_Article_ID",
          "Код": "DDS_Article_Code",
          "Наименование": "DDS_Article_Name",
          "А_РазделCFS": "CFS_Section",
          "Родитель": "Parent_ID",
          "ЭтоГруппа": "Is_Group",
          "ПометкаУдаления": "Marked_For_Deletion"
        }}
      ],
      "loader": {
        "type": "dim_full_reload",
        "target_table": "Dim_DDS_Articles"
      }
    },
    "// ... ще 15 steps для решти Dim ..."
  ]
}
```

Приклад `pipelines/fact_pnl.json`:

```json
{
  "pipeline_id": "fact_pnl",
  "name": "Fact PnL Pipeline",
  "description": "А_ОтчетPL_Свод → Fact_PnL з drill-down URL",
  "schedule": {
    "type": "cron",
    "expression": "0 2 * * *",
    "description": "Щодня о 02:00 (після Dim о 01:00)"
  },
  "parameters": {
    "period_month": {
      "type": "string",
      "format": "YYYY-MM",
      "default": "last_completed",
      "description": "Місяць для завантаження. 'last_completed' = попередній місяць"
    }
  },
  "transactional": true,
  "steps": [
    {
      "step_id": "fact_pnl",
      "extractor": {
        "type": "information_register",
        "object": "РегистрСведений.А_ОтчетPL_Свод",
        "filter": "Регистратор.Месяц = &Месяц И Регистратор.Проведен"
      },
      "transformers": [
        {"type": "uuid_to_hex", "fields": ["Регистратор", "Организация", "Подразделение", "*_ID"]},
        {"type": "drill_down", "field": "Документ_Источник",
         "outputs": ["Source_Recorder_ID", "Source_Recorder_Type", "Source_Recorder_Url", "Source_Recorder_Presentation"]},
        {"type": "enum_resolver", "field": "Source", "enum": "Перечисление.А_ИсточникPL"},
        {"type": "column_mapper", "mapping": {...}}
      ],
      "loader": {
        "type": "fact_period",
        "target_table": "Fact_PnL",
        "period_column": "Period_Month"
      }
    }
  ]
}
```

### Запуск з main.py

```python
"""
main.py — entry point для Ai_Olap orchestrator.

Modes:
    python main.py --scheduled                    # запустити daemon з APScheduler
    python main.py --run-once dim_catalogs        # ad-hoc запустити один pipeline
    python main.py --run-once fact_pnl --period 2026-02
    python main.py --validate                     # тільки валідувати JSON конфіги
"""
import argparse
from ai_olap.config.loader import load_pipelines
from ai_olap.config.validator import validate_all
from ai_olap.orchestrator.runner import PipelineRunner
from ai_olap.orchestrator.scheduler import build_scheduler
from ai_olap.core.logging import setup_logging


def main():
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--scheduled", action="store_true",
                       help="Run as daemon with APScheduler")
    group.add_argument("--run-once", metavar="PIPELINE_ID",
                       help="Run single pipeline ad-hoc")
    group.add_argument("--validate", action="store_true",
                       help="Validate all JSON configs and exit")
    parser.add_argument("--period", metavar="YYYY-MM",
                       help="Period for fact pipelines")
    args = parser.parse_args()

    setup_logging()
    pipelines = load_pipelines()
    validate_all(pipelines)

    if args.validate:
        print(f"All {len(pipelines)} pipelines valid.")
        return

    runner = PipelineRunner(pipelines)

    if args.run_once:
        runner.run(args.run_once, params={"period_month": args.period})
        return

    if args.scheduled:
        scheduler = build_scheduler(runner, pipelines)
        scheduler.start()
        print("Scheduler started. Press Ctrl+C to stop.")
        try:
            import time
            while True:
                time.sleep(60)
        except (KeyboardInterrupt, SystemExit):
            scheduler.shutdown()


if __name__ == "__main__":
    main()
```

### config.py приклад

```python
"""config.py — статичні налаштування Ai_Olap."""
import os
from pathlib import Path
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent
PIPELINES_DIR = PROJECT_ROOT / "pipelines"
LOGS_DIR = PROJECT_ROOT / "logs"

load_dotenv(PROJECT_ROOT / ".env")

# 1С BaseERP COM
CONN_ERP = os.getenv("CONN_ERP",
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# OlapBASERP SQL Server
CONN_OLAP = os.getenv("CONN_OLAP", (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;DATABASE=OlapBASERP;"
    "UID=sa;PWD=Brw739182465!;"
))

# APScheduler
SCHEDULER_TIMEZONE = "Europe/Kiev"
SCHEDULER_JOB_DEFAULTS = {
    "coalesce": True,           # collapse missed runs
    "max_instances": 1,         # no parallel runs of same job
    "misfire_grace_time": 3600  # 1 hour grace
}

# Logging
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
LOG_FORMAT = "json"             # json | text

# ETL behavior
DEFAULT_BATCH_SIZE = 1000       # bulk insert chunk
RETRY_ATTEMPTS = 3
RETRY_BACKOFF = 5               # seconds, exponential
```

---

## Чому така архітектура

| Вимога користувача | Як вирішено |
|---|---|
| `main.py` | ✅ Entry point з argparse, 3 режими |
| `config.py` | ✅ Static config + .env |
| JSON налаштування "що, звідки, яким запитом" | ✅ `pipelines/*.json` з jsonschema validation |
| Модульність "як інженер Python" | ✅ ai_olap/ package з core/extractors/transformers/loaders/orchestrator/config/utils |
| Scheduler з заданням часу виконання | ✅ APScheduler cron у `pipelines/<id>.json.schedule` |
| Перенос таблиць 1С → OlapBASERP | ✅ Extractors (COM) → Transformers → Loaders (pyodbc) |
| Наглядно і керовано | ✅ JSON конфіги читабельні; logs JSON; ETL_Runs табличка |

---

## Sources

- [GitHub: dlt-hub/dlt](https://github.com/dlt-hub/dlt) — modern Python data load tool
- [GitHub: AlisaMitchikov/ETL_Tool](https://github.com/AlisaMitchikov/ETL_Tool) — multi-source ETL → SQL Server data warehouse
- [GitHub: agronholm/apscheduler](https://github.com/agronholm/apscheduler) — Python task scheduling
- [GitHub: qkhan07/etl-integration-python-sqlserver](https://github.com/qkhan07/etl-integration-python-sqlserver) — Python ETL з Task Scheduler
- [GitHub: 11AbiRam11/ETL_Py-to-Sql](https://github.com/11AbiRam11/ETL_Py-to-Sql) — production-inspired ETL
- [Tiago Valverde: Building a Basic ETL Pipeline in Python with OOP](https://www.tiagovalverde.com/posts/building-a-basic-etl-pipeline-in-python-with-oop) — OOP patterns
- [Aliakbar Hosseinzadeh: ETL/ELT Pipeline Project Structure](https://medium.com/@aliakbarhosseinzadeh/structuring-an-etl-pipeline-project-best-practices-5ed1e4d5a601) — directory structure
- [GitHub Gist: AndrewV6/1C COMConnector with Python](https://gist.github.com/AndrewV6/22cb2d8625f0ba917213b41681aeecbe) — V83.COMConnector basics
- [12 Best Airflow Alternatives 2026 — Airbyte](https://airbyte.com/top-etl-tools-for-sources/airflow-alternatives) — порівняння
- [Python Data Pipeline Tools 2026 — UK Data Services](https://ukdataservices.co.uk/blog/articles/python-data-pipeline-tools-2025) — Airflow vs Prefect vs Dagster
- [Kedro: Modular Pipelines](https://docs.kedro.org/en/stable/nodes_and_pipelines/modular_pipelines.html) — Kedro pattern
- [APScheduler Docs](https://apscheduler.readthedocs.io/) — scheduling library docs
