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
| **Останнє оновлення** | 2026-05-18 (Свод_ПрочиеАктивыПассивы_Прямой LIVE: Fact_Balance +колонка **TaxType** (`Перечисление.ТипыНалогов`), Source=ПустаяСсылка прямих рухів ПАП; enum_resolver: пуста ссилка enum→"ПустаяСсылка"; ETL 3 міс. Success; verify PASS — Налоги 9 331 275,92 розріз TaxType==Карточка; PL.pbix колонка «ТипНалога» через MCP — потрібен Desktop Refresh+Ctrl+S) |

## Поточний стан проекту

| Stage | Опис | Стан | Hash | Дата |
|---|---|---|---|---|
| **Stage 1** | 8 нових об'єктів метаданих 1С + ОбработкаПроведения | ✅ DONE | `075d0ea08` | 2026-05-01 → 02 |
| **Stage 2** | SQL DDL OlapBASERP — 24 таблиці | ✅ DONE | `df732a73f` | 2026-05-02 → 03 |
| **Stage 3** | Python ETL Ai_Olap (SQL-first; 16 Dim + Bridge + 2 Fact + scheduler) | ✅ DONE | `8d5ebf3a1` | 2026-05-03 |
| **Stage 4** | Power BI PBIX × 2 (PnL + Cashflow) + DAX | 🔄 IN PROGRESS | — | 2026-05-03 PL.pbix модельна частина зібрана |
| **Balance** | Управлінський баланс OLAP-цикл (Fact_Balance + Dim_PAP_Articles + PL.pbix модель балансу) | ✅ DONE (інфра) | `7ea6d83c5`+ | 2026-05-16 |
| **Balance перенос** | Реальний обмін А_ОтчетБаланс_Свод→Fact_Balance (Себест 4775, січ/ТОВ); фікс Source enum (ИсточникиУправленческогоБаланса 31) | ✅ DONE | — | 2026-05-17 |
| **Balance ДенСр** | `Свод_ДенежныеСредства` LIVE: +4 ден. Source у Fact_Balance (безнал→БанкСчёт, нал→Касса, підзвіт→ФизическоеЛицо, в путі→плуг); Σ КО=75 265 344,95; «безнал» КО=50 435 887,99==УпрБаланс. Себест не регресувала | ✅ DONE | — | 2026-05-17 |
| **Balance OLAP-модель** | Dim_Warehouses (347, ієрарх. без Кода, recursive-CTE `_Reference502`); зв'язки Fact_Balance→{Cash_ID→ДенежныеСредства, Item_ID→Номенклатура, Warehouse_ID→Склады, Individual_ID→ФизическиеЛица}; PL.pbix 1С-нотація RU + сховані тех.колонки + 5 ієрархій | ✅ DONE | — | 2026-05-17 |
| **Balance РасчетыСПартнерами** | `Свод_РасчетыСПартнерами` LIVE: +2 Source (РасчСКлиент/ПоставщПоСрокам, Стаття-по-ресурсу як штатна ДвиженияАктивовПассивов) у Fact_Balance; Σ клієнти КО=12 338 631,09 / постач КО=−62 806 237,00==УпрБаланс; **Dim_ObjektyRaschetov** (плоский `_Reference319`, 13957) + ETL/верифікація PASS; Себест/ДенСр не регресували. PL.pbix таблиця+зв'язок створені через MCP (нативний Query partition), drill-down перевірено DAX | ✅ DONE (ETL+Dim+PL.pbix; потрібен Ctrl+S) | — | 2026-05-17 |
| **Balance ПрочиеАктивыПассивы_Прямой** | `Свод_ПрочиеАктивыПассивы_Прямой` LIVE: **Source=ПустаяСсылка** (прямі рухи ПАП, 0 JOIN, ВЫРАЗИТЬ Аналітика). Fact_Balance +**колонка TaxType** (`_Fld56130RRef`→`Перечисление.ТипыНалогов`); ALTER ADD TaxType varchar(50); fact_balance.json (raw_sql+enum_resolver+column_map); FROZEN_ENUMS+WHITELIST ТипыНалогов(14); **enum_resolver: пуста ссилка enum→"ПустаяСсылка"** (varbinary_to_uuid мапить 16 нулів→None ДО resolver; інакше Source NOT NULL). ETL 2025-12/2026-01/2026-02 Success. `verify_olap_balance_papdirect.py` PASS: Source=ПустаяСсылка по статтях==регістр/ПАП/УпрБаланс (Налоги 9 331 275,92 / ОС −149 202,85 / Прибыли −110 616 551,99 / ИТОГ −101 434 478,92); розріз Налоги по TaxType==Карточка (НДС 9 246 711,36/ДругиеНалоги 72 252,00/НДФЛ 4 925,02/ВоенныйСбор 1 368,07/НачисленныйЕСВ 6 019,47); 3 міс. співіснують; Себест/ДенСр/Расч не регресували; PnL Глобино-2 не регрес. PL.pbix: колонка **«ТипНалога»** (Fact_Balance) + **Dim «ТипыНалогов»** (`Dim_TaxTypes` 15 рядків, native query) + зв'язок `Fact_Balance[ТипНалога]→ТипыНалогов[TaxType]` (Many→One, active) — через MCP; потрібен Desktop Refresh Fact_Balance+«ТипыНалогов»+Ctrl+S (навігатор кешує §13.1) | ✅ DONE (ETL+verify+PL.pbix col+Dim+зв'язок; потрібен Refresh+Ctrl+S) | — | 2026-05-18 |
| **Balance ОплатаТруда** | `Свод_ОплатаТруда` LIVE (BSL+приймання): статья «Оплата труда» **БЕЗ аналітики** під **Source=ПустаяСсылка** (ПАП Источник=пусто, Статья=&ОТ; рішення 2026-05-18 — старий розклад ∝ ВзаиморасчетыССотрудниками відмінено). Σ КО серверно: дек **−3 875 135,00** / січ **−7 196 698,44** / лют **−9 972 924,59** == ПАП.ОстаткиИОбороты(Статья=ОТ) == УпрБаланс. **OLAP-сторона: НОВИХ Source/Dim/колонок PBIX НЕ потрібно** (ОТ під наявним Source=ПустаяСсылка, без субконто/TaxType). verify_olap_balance_papdirect.py еталони оновлені (+«Оплата труда» −7 196 698,44; ИТОГ Source=пусто **−108 631 177,36**). ⏳ ETL fact_balance 3 періоди + прогін verify_olap — **очікує дозволу (запис у OlapBASERP — авто-блок)**; PL.pbix модель НЕ міняється — лише Refresh+Ctrl+S | ⏳ BSL+verify DONE; ETL/verify_olap PENDING (дозвіл) | — | 2026-05-18 |
| **Balance ПрочиеРасходыДоходы — ПОВНИЙ БАЛАНС** | `Свод_ПрочиеРасходыДоходы` LIVE (BSL+приймання 2026-05-18): Источник∈{ПрочиеДоходы/ПрочиеРасходы/ПартииПрочихРасходов}, зеркало канон-Прямого, `Source=Т.Источник`. **Discovery: ніс ВЕСЬ незакритий гэп → повний баланс Актив=Пассив досягнуто; `СверкаСПАП` зайва.** Регістр `А_ОтчетБаланс_Свод` Σ КО (OD-3, ВСІ Source) = 0,00 == штатний Отчет.УправленческийБаланс (дек2025 278 093 267,32 / янв2026 288 787 750,11) + до підрозділів == ПАП.ОстаткиИОбороты. **OLAP: НОВИХ Source/Dim/колонок PBIX НЕ потрібно** (Источник у FROZEN 31, без субконто). verify_olap_balance_papdirect.py доповнено блоком «ПОЛНЫЙ БАЛАНС». ⏳ ETL fact_balance 3 періоди + verify_olap — **очікує дозволу (запис у OlapBASERP — авто-блок)**; PL.pbix модель НЕ міняється — лише Refresh+Ctrl+S | ⏳ BSL+verify DONE; ETL/verify_olap PENDING (дозвіл) | — | 2026-05-18 |
| **Stage 5** | Windows Task Scheduler 02:00 щоночі | ⏳ PLANNED | — | — |

**Acceptance verified end-to-end (Stage 1+2+3)**: 🎯 **Глобино-2 / Source=ERP_Income / Period=2026-02-01 / Sum_ERP_Grn = 38 432 968.66 ₴** (точно ± 0.01) — pytest `tests/test_etl_acceptance_globyno2.py` PASS після `python main.py` (одна команда, без аргументів).

**Acceptance Balance (2026-05-18, фактичний стан):** у `Документ.А_ФинРез_Баланс`
активні **5** `Свод_*`: `Свод_СебестоимостьТоваров` + `Свод_ДенежныеСредства` +
`Свод_РасчетыСПартнерами` + `Свод_ПрочиеАктивыПассивы_Прямой` +
`Свод_ОплатаТруда` (решта 1 `Свод_ПрочиеРасходыДоходы` + `СверкаСПАП`
закоментовані — див. `knowledge_Balanse/balanse_pattern_and_roadmap.md`).
Прямой+ОплатаТруда → Source=ПустаяСсылка ИТОГ КО=**−108 631 177,36** (Прямой
−101 434 478,92 + ОТ −7 196 698,44); ОТ БЕЗ субконто; OLAP ETL ОТ очікує
дозволу (запис у OlapBASERP). Штатний
`tests/test_etl_acceptance_balance.py` (вимагає Σ Close≈0, Актив=Пасив —
ПОВНИЙ баланс) очікувано FAIL (ще не всі `Свод_*`). Релевантні часткові
перевірки (PASS):
- Себест: `Fact_Balance` «Товары на оптовых складах» Sum_Close=**83 627 719,44**
  (==регістр==ПАП до копійки, не регресувала);
- ДенСр: 4 ден. Source у Fact_Balance, Σ КО денежної групи=**75 265 344,95**;
  «Денежные средства (безналичные)» КО=**50 435 887,99**==ПАП==Управлінський
  баланс; підзвіт деталізований по `ФизическоеЛицо` (Individual_ID→
  Dim_Individuals); тести `test_balans_densr_pretest.py`/`_verify.py` PASS.
- РасчетыСПартнерами: 2 Source (`РасчетыСКлиентамиПоСрокам` КО=**12 338 631,09**,
  `РасчетыСПоставщикамиПоСрокам` КО=**−62 806 237,00**) у Fact_Balance;
  статті ЗадолженностьКлиентов **61 165 524,68** / ПолученныеАвансы
  **−48 826 893,59** / ВыданныеАвансы **68 949 869,33** /
  ЗадолженностьПередПоставщиками **−131 756 106,33** == ПАП == УпрБаланс;
  деталізація по `ОбъектРасчетов` (SettlementObj_ID→**Dim_ObjektyRaschetov**
  плоский `_Reference319`, 13957 рядків; 2291 деталь покриття 100% + 57
  плугів); Себест/ДенСр не регресували; тести
  `test_balans_raschety_pretest.py`/`_verify.py`/`verify_olap_balance_raschety.py`
  PASS. PL.pbix: таблиця «ОбъектыРасчетов» (партиція Dim_ObjektyRaschetov,
  **нативний `[Query=…]`** — навігатор SQL кешує нові таблиці в сесії, обхід
  через native query; state=Ready, 13957) + зв'язок
  `Fact_Balance[SettlementObj_ID]→ОбъектыРасчетов[SettlementObj_ID]`
  (Many→One, OneDirection, active) + model Calculate + сховані тех.колонки
  (видима «ОбъектРасчетов» 1С-нотація) — створено через MCP, drill-down
  перевірено DAX. ⚠️ **Зміни in-memory — потрібен Ctrl+S у Power BI Desktop.**
- ПрочиеАктивыПассивы_Прямой (2026-05-18): **Source=ПустаяСсылка** (прямі
  рухи ПАП, 0 JOIN) у Fact_Balance; +**колонка TaxType**
  (`_Fld56130RRef`→`Перечисление.ТипыНалогов`, 14 значень FROZEN). Σ по
  статтях == регістр/ПАП/УпрБаланс: Налоги КО=**9 331 275,92**, ОС
  **−149 202,85**, Прибыли и убытки **−110 616 551,99**, ИТОГ
  **−101 434 478,92**; розріз Налоги по TaxType == Карточка
  (НДС **9 246 711,36** / ДругиеНалоги **72 252,00** / НДФЛ **4 925,02** /
  ВоенныйСбор **1 368,07** / НачисленныйЕСВ **6 019,47**); 2025-12/2026-01/
  2026-02 співіснують; Себест/ДенСр/Расч + PnL(Глобино-2) не регресували.
  Грабля: `varbinary_to_uuid` мапить пусту ссилку enum (16 нулів)→None ДО
  `enum_resolver` → фікс: `enum_resolver.transform` None→"ПустаяСсылка"
  (Source NOT NULL). PL.pbix: колонка **«ТипНалога»** (String, 1С-нотація,
  SourceColumn=TaxType) + **Dim «ТипыНалогов»** (`Dim_TaxTypes` 15 рядків:
  14 enum + ПустаяСсылка; native query; ключ=метаім'я) + зв'язок
  `Fact_Balance[ТипНалога]→ТипыНалогов[TaxType]` Many→One active (FK 100%)
  додані через MCP — навігатор кешує (§13.1) → потрібен Desktop Refresh
  Fact_Balance + «ТипыНалогов» + Ctrl+S. Тести
  `test_balans_papdirect_pretest.py`/`_verify.py`/`verify_olap_balance_papdirect.py`
  (+ `scripts/seed_dim_tax_types.py`, `scripts/ddl_dim_tax_types.sql`) PASS.
Σ Close≈0 та канон 289 064 974,43 стануть застосовні коли активують решту
`Свод_*` (ETL лише копіює регістр).

**Stage 3 default mode** (commit `8d5ebf3a1`): `python main.py` без прапорців виконує повний прогон **всіх Dim + всіх періодів Fact** (TRUNCATE + INSERT). Для пер-місячного режиму — `--period YYYY-MM` (idempotent DELETE WHERE + INSERT). Daemon-режим `--scheduled` поки **не** використовуємо — йде ручне тестування.

**Stage 3 row counts (live, Feb 2026):**

| таблиця | rows | джерело |
|---------|------|---------|
| Dim_Organizations | 1 | тільки ТОВ ІНДАСТРІАЛБУД |
| Dim_Departments | 385 | СтруктураПредприятия |
| Dim_Directions | 9 | НаправленияДеятельности |
| Dim_Counterparties | 4 874 | Контрагенты |
| Dim_Contracts | 8 152 | ДоговорыКонтрагентов |
| Dim_Items | 40 745 | Номенклатура |
| Dim_ItemGroups | 14 | ГруппыФинансовогоУчетаНоменклатуры |
| Dim_Individuals | 725 | ФизическиеЛица |
| Dim_Users | 63 | Пользователи |
| Dim_BankAccounts | 99 | БанковскиеСчетаОрганизаций |
| Dim_Currencies | 4 | Валюты |
| Dim_DDS_Articles | 425 | СтатьиДвиженияДенежныхСредств |
| Dim_Expense_Articles | 344 | СтатьиРасходов (план характ.) |
| Dim_Income_Articles | 15 | СтатьиДоходов (план характ.) |
| Dim_PL_Articles | 71 | А_Статьи_PL |
| Dim_PL_ArticleGroups | 8 | А_ГруппаСтатей_PL |
| Bridge_PLArticle_DDS | **0** | Реквізит не заповнений (pending фінансист) |
| Fact_PnL | 3 937 | А_ОтчетPL_Свод (7 distinct Source) |
| Fact_Cashflow | 4 652 | А_ОтчетDDS_Свод |
| **Σ dim_catalogs** | **55 934** | у 17 кроках |

**Час повного `python main.py`:** ~2 секунди end-to-end (BaseERP backend через pyodbc).

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
┌─ ШАР 4: POWER BI (Stage 4, 🔄 IN PROGRESS) ──────────────────┐
│ PL.pbix       — 18 видимих таблиць (15 Dim 1С-нотація + Fact_PnL│
│                 + Calendar + Table_Measures), 11 зв'язків     │
│                 Fact→Dim, 0 DAX-мір (TODO).                   │
│                 Live: _Rarzrabotki/Olap/PowerBi/PL.pbix       │
│                 Деталі: knowledge_Olap/olap_powerbi_pl_pbix.md │
│ Cashflow.pbix — TODO; не починали                            │
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
| 8 | [olap_powerbi_model.md](olap_powerbi_model.md) | ПОСТІЙНИЙ | PLANNED | DAX-міри, сторінки PBIX (загальна архітектура Stage 4) |
| 9 | [olap_powerbi_pl_pbix.md](olap_powerbi_pl_pbix.md) | ЗМІННИЙ | IN PROGRESS | реалізація PL.pbix — таблиці, зв'язки, partition'и, TODO |
| 10 | [olap_acceptance_etalons.md](olap_acceptance_etalons.md) | ЗМІННИЙ | DONE | еталони verification (оновлюється помісячно) |
| 11 | [olap_deviations_from_spec.md](olap_deviations_from_spec.md) | ЗМІННИЙ | DONE | фактичні зміни vs spec v3 final |

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
- `_Rarzrabotki/Olap/PowerBi/PL.pbix` — PL дашборд (Stage 4 in-progress)

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
| `olap_powerbi_model.md` | Реалізація Stage 4; додавання DAX-мір (загальна архітектура) |
| `olap_powerbi_pl_pbix.md` | Зміна моделі PL.pbix (таблиці, зв'язки, мири, partition'и) |
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
