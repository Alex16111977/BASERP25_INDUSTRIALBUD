# KNOWLEDGE_MAP_OLAP — OLAP BI-конвеєр BASERP25

> **Призначення:** on-disk knowledge base для розробника/інженера. Описує архітектуру і реалізацію BI-конвеєра PnL+Cashflow з BASERP25 у Power BI через проміжну SQL-базу OlapBASERP.
>
> Не для NotebookLM. Не для CFO/фінансиста (для них — `_Rarzrabotki/notebook/knowledge_PL/`).

---

## Project Header

| | |
|---|---|
| **Проект** | BASERP25 → OlapBASERP → Power BI (PnL + Cashflow) |
| **Організація** | ТОВ ІНДАСТРІАЛБУД |
| **1С конфігурація** | BAS ERP 2.5 (v2.13) |
| **SQL Server** | localhost (SQLSERVER instance), порт 1433, Standard Edition 2022 |
| **Гілка** | `main` (Stage 3 merged 2026-05-03 з `claude/kind-ptolemy-729f31`) |
| **Основний repo** | `C:\Configuration_downloads\BASERP25\` |
| **Stage 3 код** | `_Rarzrabotki/Olap/Ai_Olap/` (SQL-first Python ETL) |
| **База знань створена** | 2026-05-03 |
| **Останнє оновлення** | 2026-05-03 (Stage 3 done) |

## Поточний стан проекту

| Stage | Опис | Стан | Hash | Дата |
|---|---|---|---|---|
| **Stage 1** | 8 нових об'єктів метаданих 1С + ОбработкаПроведения | ✅ DONE | `075d0ea08` | 2026-05-01 → 02 |
| **Stage 2** | SQL DDL OlapBASERP — 24 таблиці | ✅ DONE | `df732a73f` | 2026-05-02 → 03 |
| **Stage 3** | Python ETL Ai_Olap (SQL-first; 16 Dim + Bridge + 2 Fact + scheduler) | ✅ DONE | `952c46db7` | 2026-05-03 |
| **Stage 4** | Power BI PBIX × 2 (PnL + Cashflow) + DAX | ⏳ PLANNED | — | — |
| **Stage 5** | Windows Task Scheduler 02:00 щоночі | ⏳ PLANNED | — | — |

**Acceptance verified end-to-end (Stage 1+2+3)**: 🎯 **Глобино-2 / Source=ERP_Income / Period=2026-02-01 / Sum_ERP_Grn = 38 432 968.66 ₴** (точно ± 0.01) — pytest `tests/test_etl_acceptance_globyno2.py` PASS після `python main.py --run-once dim_catalogs && python main.py --run-once fact_pnl --period 2026-02`.

**Stage 3 row counts (Feb 2026):**
- Dim таблиці заповнено через `dim_catalogs`: 55 934 рядки у 17 кроках
- Fact_PnL: 3 937 рядків (7 distinct Source)
- Fact_Cashflow: 4 652 рядки

---

## Потік даних (4 шари)

```
┌─ ШАР 0: ДЖЕРЕЛА (read-only, існуючі) ─────────────────────────┐
│ Excel → Документ.А_ОтчетPL                                    │
│ Реалізації, ОтражениеЗП → РегНакоп.ПрочиеРасходы              │
│                          → РегНакоп.ПрочиеДоходы              │
│                          → РегНакоп.ВыручкаИСебестоимостьПродаж│
│ ПКО/РКО/ПлатПоруч → РегНакоп.ДенежныеСредства{Безналичные,Наличные}│
│ Казна (через А_ПереносДвижений) → РегНакоп.А_ДвиженияДенегИзКазны│
│ Справочники: А_Статьи_PL, А_ГруппаСтатей_PL, СтатьиДДС, ...   │
└────────────┬──────────────────────────────────────────────────┘
             │ свёртка при проведенні документа (BSL ОбработкаПроведения)
             ▼
┌─ ШАР 1: НОВЕ В 1С (Stage 1, ✅) ──────────────────────────────┐
│ Документ.А_ФинРез_PL  → РегСв.А_ОтчетPL_Свод (8 dim, 5 res)  │
│ Документ.А_ФинРез_DDS → РегСв.А_ОтчетDDS_Свод (8 dim, 2 res) │
│ + 3 Перечисления (А_ИсточникPL/DDS/А_РазделыCFS)              │
│ + Реквізит СтатьиДДС.А_РазделCFS                              │
└────────────┬──────────────────────────────────────────────────┘
             │ pyodbc (primary, BaseERP MSSQL backend) +
             │ V83.COMConnector (fallback only — virtual tables)
             ▼
┌─ ШАР 2: PYTHON ETL Ai_Olap (Stage 3, ✅) ────────────────────┐
│ _Rarzrabotki/Olap/Ai_Olap/                                    │
│ ├── ai_olap/                                                  │
│ │   ├── core/         — connections, logging, exceptions      │
│ │   ├── extractors/   — sql_backend (primary) + com (fallback)│
│ │   ├── transformers/ — varbinary→uuid, dates, enums, drill   │
│ │   ├── loaders/      — dim/fact/bridge/etl_runs              │
│ │   ├── orchestrator/ — Pipeline + Runner + APScheduler       │
│ │   ├── config/       — JSON-schema validator                 │
│ │   └── utils/mapping_resolver.py                             │
│ ├── mapping/baserp_storage.json (56 obj 1С→SQL mapping)       │
│ ├── pipelines/        — 5 JSON-конфіги                        │
│ ├── tests/            — pytest 10/10 PASS                     │
│ └── main.py           — CLI: --validate/run-once/scheduled    │
└────────────┬──────────────────────────────────────────────────┘
             │ INSERT через pyodbc (bulk_insert, fast_executemany)
             ▼
┌─ ШАР 3: SQL OLAPBASERP (Stage 2, ✅) ────────────────────────┐
│ 24 таблиці на SQL Server localhost:                           │
│ - 3 Fact: Fact_PnL, Fact_Cashflow, Fact_CF_Balance            │
│ - 16 Dim з UUID-ключами + variants                            │
│ - 1 Bridge: PLArticle_DDS                                     │
│ - 4 Util: Calendar (2191 днів), CFS_Sections, Table_Measures, │
│           ETL_Runs                                            │
└────────────┬──────────────────────────────────────────────────┘
             │ Sql.Database("localhost","OlapBASERP") — sa auth
             ▼
┌─ ШАР 4: POWER BI (Stage 4, ⏳) ──────────────────────────────┐
│ PnL.pbix      — 14 таблиць, ~70 DAX, 6 сторінок,             │
│                 слайсер Source (8) + 5 рівнів маржі          │
│ Cashflow.pbix —  9 таблиць, ~25 DAX, 5 сторінок,             │
│                  слайсери Source (3) + CFS_Section (4)       │
└──────────────────────────────────────────────────────────────┘
```

---

## Файли в цій базі знань

| # | Файл | Тип | Стан | Тема |
|---|---|---|---|---|
| 1 | [KNOWLEDGE_MAP_OLAP.md](KNOWLEDGE_MAP_OLAP.md) | manifest | DONE | індекс + поточний стан |
| 2 | [olap_architecture_overview.md](olap_architecture_overview.md) | ПОСТІЙНИЙ | DONE | 4-шарова архітектура, принципи |
| 3 | [olap_1c_objects.md](olap_1c_objects.md) | ПОСТІЙНИЙ | DONE | 8 нових об'єктів метаданих 1С |
| 4 | [olap_obrabotka_provedeniya.md](olap_obrabotka_provedeniya.md) | ПОСТІЙНИЙ | DONE | BSL логіка А_ФинРез_PL/DDS |
| 5 | [olap_data_sources_erp.md](olap_data_sources_erp.md) | ПОСТІЙНИЙ | DONE | джерела з ERP (регістри/документи/справочники) |
| 6 | [olap_sql_schema.md](olap_sql_schema.md) | ПОСТІЙНИЙ | DONE | 24 таблиці OlapBASERP, DDL |
| 7 | [olap_etl_pipeline.md](olap_etl_pipeline.md) | ПОСТІЙНИЙ | DONE | Python ETL Ai_Olap (Stage 3, SQL-first архітектура) |
| 8 | [olap_powerbi_model.md](olap_powerbi_model.md) | ПОСТІЙНИЙ | PLANNED | DAX-міри, сторінки PBIX (Stage 4) |
| 9 | [olap_acceptance_etalons.md](olap_acceptance_etalons.md) | ЗМІННИЙ | DONE | еталони verification (оновлюється помісячно) |
| 10 | [olap_deviations_from_spec.md](olap_deviations_from_spec.md) | ЗМІННИЙ | DONE | фактичні зміни vs spec v3 final |

**Тип файлу:**
- **ПОСТІЙНИЙ** — змінюється рідко (тільки якщо зміна архітектури/коду)
- **ЗМІННИЙ** — оновлюється з кожною новою perioдою (acceptance) або фіксом (deviations)

**Стан:**
- **DONE** — описує реалізовану функціональність
- **PLANNED** — описує запланований стан, треба реалізувати

---

## Cross-references

### Документи проекту (`docs/superpowers/`)
- **Spec v3 final** (повна архітектурна специфікація): `docs/superpowers/specs/2026-05-01-olap-baserp-architecture-design-v3-final.md`
- **Implementation plan Stage 1**: `docs/superpowers/plans/2026-05-01-olap-baserp-stage1-1c-metadata.md`
- **Implementation prompt**: `docs/superpowers/specs/2026-05-01-olap-baserp-IMPLEMENTATION_PROMPT.md`

### Сусідні бази знань
- **`knowledge/baserp25_knowledge.md`** — Golden Rules архітектури BAS ERP 2.5 (загальний контекст)
- **`knowledge/perenos_dvizheniy_iz_kazny.md`** — паттерн `Движения.<Регистр>.Загрузить()` (еталон BSL)
- **`knowledge/a_otchet_pl_overview.md`** — оригінальний звіт А_ОтчетPL (звідки портувався SQL)
- **`knowledge/exchange_erp_kazna.md`** — обмін ERP↔Казна (шар 0 джерело)
- **`knowledge_PL/pl_methodology.md`** — економічна семантика 68 PL-статей (для контексту)
- **`knowledge_PL/pl_dds_mapping.md`** — матриця PL↔ДДС

### Memory whitelist
- `feedback_no_1c_changes.md` — whitelist 8 об'єктів 1С які дозволено створювати (Alex 2026-05-01)
- `sql_olap_baserp_credentials.md` — credentials SA для OlapBASERP

### Source files
- `Reports/А_ОтчетPL/Ext/ObjectModule.bsl` — джерельний 8-CTE SQL що портувався
- `Documents/А_ФинРез_PL/Ext/ObjectModule.bsl` — реалізована BSL ОбработкаПроведения
- `Documents/А_ФинРез_DDS/Ext/ObjectModule.bsl` — реалізована BSL
- `Documents/А_ПереносДвиженийИзКазны/Ext/ObjectModule.bsl` — паттерн реєстру через `Загрузить`
- `Catalogs/СтатьиДвиженияДенежныхСредств.xml` — реквізит А_РазделCFS (line ~907)
- `_Rarzrabotki/Python/Olap/ddl/*.sql` — DDL OlapBASERP (4 файли)
- `_Rarzrabotki/Python/test/test_acceptance_finrez_*_lutyi.py` — acceptance tests
- `_Rarzrabotki/Olap/Ai_Olap/` — Stage 3 код (SQL-first ETL, 11 commits)
- `_Rarzrabotki/Olap/Ai_Olap/mapping/baserp_storage.json` — 1С→SQL mapping (56 об'єктів)
- `_Rarzrabotki/Olap/Ai_Olap/pipelines/*.json` — 5 декларативних pipeline-конфіги
- `_Rarzrabotki/Olap/Ai_Olap/tests/` — pytest acceptance + unit tests
- `_Rarzrabotki/Olap/Ai_Olap/README.md` — quickstart + troubleshooting

---

## Інструкції оновлення

### Коли регенерувати кожен файл

| Файл | Тригер регенерації |
|---|---|
| `KNOWLEDGE_MAP_OLAP.md` | Завершення нового Stage; новий commit hash; зміна списку файлів |
| `olap_architecture_overview.md` | Зміна шарової архітектури (рідко) |
| `olap_1c_objects.md` | Зміна метаданих об'єктів (поля, типи); додавання нових Source значень |
| `olap_obrabotka_provedeniya.md` | Правка BSL у А_ФинРез_PL/DDS ОбработкаПроведения |
| `olap_data_sources_erp.md` | Додавання нового джерельного регістра/документа |
| `olap_sql_schema.md` | Зміна DDL OlapBASERP (нові таблиці, індекси) |
| `olap_etl_pipeline.md` | Реалізація Stage 3; правка ETL скриптів |
| `olap_powerbi_model.md` | Реалізація Stage 4; додавання DAX-мір |
| `olap_acceptance_etalons.md` | Кожен новий місяць після проведення документів |
| `olap_deviations_from_spec.md` | Кожне нове відхилення від spec v3 final |

### Як оновлювати

1. **Перевір актуальність даних:**
   - cf-validate / SQL count / acceptance тести
2. **Внеси зміни** в worktree або main config (за Rule #4 — обидва):
   ```bash
   # Edit у worktree
   cp <file>.md C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Olap/
   ```
3. **Cross-references** — перевір що нові згадані файли/шляхи існують
4. **Git commit** — `git add knowledge_Olap/<file>.md && git commit -m "knowledge_Olap: <зміна>"`

---

## TL;DR

**Що це за проект:** BI-конвеєр що бере дані з 1С BAS ERP 2.5 (BASERP25), робить свёртку PnL і Cashflow при проведенні документів `А_ФинРез_PL/DDS`, передає у проміжну SQL-базу `OlapBASERP` (24 таблиці), і відображає у двох Power BI дашбордах для управлінського обліку.

**Ключова ідея:** свёртка робиться **в 1С при проведенні документа** (не в Python ETL), тому фінансист бачить результати свёртки одразу через 1С UI; Python ETL лише копіює готові регістри сведень у SQL.

**Колонка `Source`** — основний інструмент розділення 8 джерел PnL і 3 джерел Cashflow у Power BI слайсерах і DAX мірах.

**Stage 1+2+3 готові** і протестовані — повна parity з оригінальним Reports/А_ОтчетPL до копійки (Глобино-2 / ERP_Income = 38 432 968.66 ₴ exact). **Stage 4** (Power BI) і **Stage 5** (Windows Scheduler) заплановані з повним описом у spec v3 final, чекають реалізації.

**Архітектурна зміна Stage 3 vs початковий план:** Реалізовано **SQL-first** замість COM-first. Прямий pyodbc-доступ до MSSQL backend BaseERP (через mapping `1С_object → _Reference329` тощо) у рази швидший за COM. COM лишений тільки для віртуальних таблиць (`.Остатки`, `.Обороты`). Деталі — у [olap_etl_pipeline.md](olap_etl_pipeline.md) та `_Rarzrabotki/Olap/Ai_Olap/RESEARCH.md`.
