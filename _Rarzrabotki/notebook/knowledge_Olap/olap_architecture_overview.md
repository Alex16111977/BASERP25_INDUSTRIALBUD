# OLAP Architecture Overview

> 4-шарова архітектура BASERP25 → OlapBASERP → Power BI. Призначення, принципи, чому такий дизайн.
>
> ⚠️ **Stage v2 (2026-05-05) ВАЖЛИВЕ ПРИМІТКА:** DAX-приклади і деякі назви колонок у цьому файлі — **pre-Stage-v2** (історичний контекст принципів). Поточні Stage v2 актуалії:
> - `Sum_ERP_Grn → Sum_Fact`, `Sum_Plan_Grn → Sum_Excel`, `Source_Recorder_* → Document_ID + новий Dim_Documents`.
> - `Source` для PnL: 8 значень → 2 (`PL_Excel`, `PL_ЕРП`). Розрізнення CoGS/Income/OpEx — через виміри `СтатьяДоходов`/`СтатьяРасходов`.
> - Казна повністю прибрана з PnL-моделі.
> - Канонічне джерело актуальної моделі: [olap_changelog_2026_05.md](olap_changelog_2026_05.md) §"Stage v2".

---

## Принципи (rationale)

### 1. Свёртка в 1С, не в ETL

**Що:** документи `А_ФинРез_PL` і `А_ФинРез_DDS` при проведенні виконують єдиний SQL-запит з 7-8 CTE та фінальним `ОБЪЕДИНИТЬ ВСЕ` з колонкою-маркером `Source`. Результат записується у регістр сведень через `Движения.<Регистр>.Загрузить(Запрос.Выполнить().Выгрузить())`.

**Чому не в Python ETL:**
- **Транспарентність для фінансиста.** Якщо щось виглядає підозріло у Power BI, фінансист відкриває документ `А_ФинРез_PL` за лютий 2026 у 1С UI, бачить його записи у регістрі сведень, бачить регістратор → клікає → потрапляє в первинний документ (Реалізацію, ПКО, тощо).
- **Швидкість Q&A.** Свёртка на стороні 1С використовує внутрішні індекси регістрів; Python ETL мав би тягнути сирі дані по мережі COM → це повільніше.
- **Контроль перепроведення.** Якщо план PnL за лютий змінився (фінансист поправив `А_ОтчетPL`), він просто перепроводить `А_ФинРез_PL` і отримує оновлений регістр одним рухом — без перезапуску Python ETL.

**Trade-off:** код свёртки знаходиться у двох документах BSL; зміни архітектури (нові Source, нові колонки регістру) вимагають правки BSL і db-load-xml. Прийнято.

### 2. Колонка `Source` як основний дискримінатор

**Що:** у обох регістрах сведень (А_ОтчетPL_Свод і А_ОтчетDDS_Свод) є вимір `Source` (тип EnumRef.А_ИсточникPL або А_ИсточникDDS). Кожен рядок маркований однозначним джерелом.

**Чому:**
- **DAX simplicity.** `CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_Income")` — однорядковий filter без складного JOIN.
- **Слайсер у PBIX.** Користувач вмикає/вимикає окремі джерела одним кліком (наприклад, поглянути PnL без Каси).
- **Виявлення аномалій.** Source `ERP_БезPL_Расх` і `ERP_БезPL_Доход` — це **аномалії маппінгу** (ERP-дані без прив'язки до PL-статті). Окрема Source-категорія робить їх видимими у дашборді — фінансист одразу бачить, які статті ДДС забув додати у мапу А_Статьи_PL.

**Frozen identifiers:** значення Source — це рядкові константи у DAX і Python ETL. Зміна імені = breakage Power BI dashboard. Список frozen у [olap_1c_objects.md](olap_1c_objects.md).

### 3. Drill-down через `e1cib/data/...` URL

**Що:** Python ETL формує для кожного рядка Fact-таблиці колонки:
- `Source_Recorder_ID` — UUID документа-джерела (наприклад, конкретної Реалізації)
- `Source_Recorder_Type` — тип документа (наприклад, "РеализацияТоваровУслуг")
- `Source_Recorder_Url` — рядок `e1cib/data/Документ.<Type>?ref=<UUID>`
- `Source_Recorder_Presentation` — `Регистратор.Представление()` для текстового відображення

**Як використовується у Power BI:** колонка `Source_Recorder_Url` оголошується як hyperlink. Користувач клікає на рядок таблиці у дашборді → відкривається 1С Web Client / Designer з конкретним документом.

**Чому це важливо:** дашборд PnL показує `Глобино-2 / Виручка ЕРП = 38 432 968.66 ₴` — фінансист хоче розшифровку: "які саме Реалізації?" Drill-down дає список регістраторів. Клік на рядок → бачить документ → бачить акт виконаних робіт чи реалізацію → перевіряє ціни/контрагента.

### 4. Multi-source через UNION ALL

**Що:** замість N окремих регістрів сведень для кожного джерела (один на PL_Excel, один на ERP_Income, один на Казну...) — **один великий регістр з колонкою Source**.

**Альтернативи що відкинуті:**
- ❌ N регістрів сведень (N = 8 для PnL): кожен Power BI запит став би UNION 8 таблиць. Складно для DAX, повільно.
- ❌ Окрема Fact-таблиця у SQL для кожного Source: схожі проблеми + duplication колонок dim FK.

**Прийнято:** єдиний регістр з колонкою Source = архітектурний equivalent **slowly-changing-attribute pattern** для OLAP.

### 5. Python ETL — чистий копіювач, без бізнес-логіки

**Що:** Python ETL читає регістри сведень `А_ОтчетPL_Свод` / `А_ОтчетDDS_Свод` і копіює їх у SQL-таблиці `Fact_PnL` / `Fact_Cashflow` 1:1 (плюс додає drill-down колонки).

**Чого Python ETL **не** робить:**
- ❌ Жодних агрегатів (SUM/GROUP BY) — це робить регістр сведень в 1С.
- ❌ Жодних бізнес-правил (виключення певних рядків, мапінги) — все вже зроблено у BSL `ОбработкаПроведения`.
- ❌ Розрахунку маржі / [Net Cashflow] — це DAX у Power BI.

**Чому:** ізоляція бізнес-логіки в одному місці (BSL модулі). Якщо логіка маржі зміниться — правимо DAX. Якщо логіка свёртки — правимо BSL. ETL не торкається.

### 6. SQL-first ETL: pyodbc до backend BaseERP, COM лише для віртуальних таблиць (Stage 3)

**Що:** Stage 3 (Ai_Olap) реалізований не як COM-first orchestrator (як було у початковому plan'і), а як **SQL-first**: 

- 16 Dim + 1 Bridge + 2 Fact таблиці зчитуються через `pyodbc → BaseERP MSSQL backend` (DSN `localhost/BaseERP/sa`).
- Mapping `1С_object → _Reference329` тощо генерується платформенним викликом `ПолучитьСтруктуруХраненияБазыДанных()` (через Python COM single-shot) і зберігається у `mapping/baserp_storage.json`.
- COM (V83.COMConnector) лишається тільки для `.Остатки/.Обороты/.ОстаткиИОбороты` — там, де платформа додає обчислювані колонки що сирі MSSQL-таблиці не дають (Fact_CF_Balance pipeline).

**Чому**:
- **Швидкість.** 3937 рядків Fact_PnL за лютий — ~58 ms через pyodbc проти ~2 s через COM (×30+ прискорення).
- **Простота транзакцій.** Усе на одному pyodbc connection без міжстекового COM/SQL state.
- **Не блокує 1С Server.** ETL не споживає сесій 1С.
- **Декларативність.** 5 pipeline JSON-конфігів замість Python module per Dim.

**Trade-off**:
- Потрібен refresh `mapping/baserp_storage.json` після значної зміни конфіги 1С (нові реквізити → нові `_Fld<N>`).
- Деякі обчислювані поля недоступні без COM — там лишилися COM extractor'и.

**Where**: [`_Rarzrabotki/Olap/Ai_Olap/RESEARCH.md`](../../Olap/Ai_Olap/RESEARCH.md) — обґрунтування і порівняння з Airflow/Prefect/Dagster (всі відкинуті як overkill для one-instance ETL). [`_Rarzrabotki/Olap/Ai_Olap/ai_olap/extractors/sql_backend.py`](../../Olap/Ai_Olap/ai_olap/extractors/sql_backend.py) — primary impl.

---

## 4 шари — детальний опис

### Шар 0 — Джерела (read-only, існуючі в 1С BAS ERP 2.5)

Не створюємо нічого нового. Читаємо існуючі об'єкти 1С:
- 4 регістри накопичення для PnL (ВыручкаИСебестоимостьПродаж, ПрочиеРасходы, ПрочиеДоходы, А_ДвиженияДенегИзКазны)
- 3 регістри накопичення для Cashflow (ДенежныеСредстваБезналичные/Наличные, А_ДвиженияДенегИзКазны)
- Документ.А_ОтчетPL (план з Excel)
- Справочники А_Статьи_PL (68 статей у 8 групах), СтатьиДДС (~425 статей)
- Інші справочники (Організація, Підрозділ, Контрагент, тощо) — як вимірювання

Деталі: [olap_data_sources_erp.md](olap_data_sources_erp.md).

### Шар 1 — Нове в 1С (Stage 1, ✅)

**8 нових об'єктів метаданих:**
- 3 перерахування (А_ИсточникPL, А_ИсточникDDS, А_РазделыCFS)
- 1 реквізит на існуючому справочнику (СтатьиДДС.А_РазделCFS)
- 2 регістри сведень (А_ОтчетPL_Свод з 9 dim/4 attr/5 res; А_ОтчетDDS_Свод з 8 dim/5 attr/2 res)
- 2 документи (А_ФинРез_PL, А_ФинРез_DDS) з повноцінною ОбработкаПроведения

Memory whitelist (Alex): тільки ці 8 об'єктів дозволено створювати, інше — stop & ask.

Деталі: [olap_1c_objects.md](olap_1c_objects.md), [olap_obrabotka_provedeniya.md](olap_obrabotka_provedeniya.md).

### Шар 2 — Python ETL Ai_Olap (Stage 3, ✅ DONE)

**Структура (live, commit `952c46db7`):**
```
_Rarzrabotki/Olap/Ai_Olap/
├── ai_olap/                          ← Python package
│   ├── core/                         — connections, logging, exceptions, decorators
│   ├── extractors/                   — sql_backend (primary) + com (fallback) + factory
│   ├── transformers/                 — varbinary→uuid, dates, enum_resolver, drill_down,
│   │                                   column_mapper, pipeline (chain runner)
│   ├── loaders/                      — dim, fact (idempotent_period), bridge, etl_runs
│   ├── orchestrator/                 — Pipeline + Runner + APScheduler
│   ├── config/                       — JSON Schema validator
│   └── utils/mapping_resolver.py     — resolve(meta) → (sql_table, fields)
├── mapping/baserp_storage.json       — 56 1С→SQL mappings (gitcommited)
├── pipelines/                        — 5 декларативних JSON-конфіги
├── tests/                            — pytest 10/10 PASS
├── main.py                           — CLI: --validate / --run-once / --scheduled
└── README.md
```

**Архітектурний вибір — SQL-first з COM-fallback:**
- **Primary read** — pyodbc прямо до MSSQL backend BaseERP (`_Reference329`, `_InfoRg55970`, ...)
- **Fallback COM** (V83.COMConnector) — тільки для віртуальних таблиць (`.Остатки`, `.Обороты`)
- **Швидкість**: 3937 рядків Fact_PnL за лютий завантажуються за ~58 ms (через COM було б ~2 s)

**Підключення:**
- BaseERP: pyodbc → `localhost / BaseERP / sa / Brw739182465!` (read-only)
- OlapBASERP: pyodbc → `localhost / OlapBASERP / sa / Brw739182465!` (write, fast_executemany)
- COM (fallback): `Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"`

**Mapping**: платформенний метод `ПолучитьСтруктуруХраненияБазыДанных()` через `mapping/refresh_mapping.py` (Python COM call) генерує `baserp_storage.json`. Refresh — після значної зміни конфіги.

**Acceptance** (Feb 2026): 🎯 Глобино-2 / ERP_Income = 38 432 968.66 ₴ exact (10/10 pytest PASS).

Деталі: [olap_etl_pipeline.md](olap_etl_pipeline.md).

### Шар 3 — SQL OlapBASERP (Stage 2, ✅)

**24 таблиці на SQL Server:**
- **3 Fact**: Fact_PnL, Fact_Cashflow, Fact_CF_Balance
- **16 Dim** з UUID-ключами (char(32)) + варіанти (Account_Type, CFS_Section денормалізація, тощо)
- **1 Bridge**: PLArticle_DDS (мапа N:M PL-статей до ДДС)
- **4 Util**: Calendar (2191 днів 2025-2030), CFS_Sections (4 розділи), Table_Measures (порожня hub-таблиця для DAX), ETL_Runs (логування ETL)

**Звичайні стандарти dimensional modeling:** snowflake schema — Fact-таблиці посилаються на Dim через FK; Dim_DDS_Articles має денормалізований CFS_Section для швидких слайсерів.

Деталі: [olap_sql_schema.md](olap_sql_schema.md).

### Шар 4 — Power BI (Stage 4, ⏳ PLANNED)

**2 PBIX файли:**
- `Управлінський PnL BASERP25.pbix` — 14 таблиць, ~25 зв'язків, ~70 DAX-мір, 6 сторінок
- `Управлінський Cashflow BASERP25.pbix` — 9 таблиць, ~12 зв'язків, ~25 DAX-мір, 5 сторінок

**Підключення:** через PowerQuery `Sql.Database("localhost", "OlapBASERP")`, аутентифікація через вкладку "База данных" у Power BI (НЕ Windows!) з sa/Brw739182465!.

**DAX-міри** (повний список — [olap_powerbi_model.md](olap_powerbi_model.md)):
- 5 рівнів маржі для PnL: Валова маржа → Валова маржа % → EBITDA → EBITDA % → Чистий прибуток
- Direct Method CFS: CFS Operating + CFS Investing + CFS Financing = CFS Total (Internal виключено)
- Reconciliation: Balance Close - Balance Open - CFS Total ≈ 0

**Слайсери:** Source (8 для PL, 3 для Cashflow), CFS_Section (4), Календар, Організація, Підрозділ.

**Drill-down:** колонка Source_Recorder_Url як hyperlink → клік відкриває первинний документ у 1С.

---

## Чому 4 шари, а не 2

**Альтернатива 1 (відкинута):** 1С → Power BI напряму через DirectQuery до 1С (через ODBC adapter або REST API).
- ❌ Повільно (1С не оптимізована для OLAP-агрегатів)
- ❌ Power BI бачить сирі регістри, треба робити свёртку у DAX (складно)
- ❌ Залежність від 1С Server (якщо стане — дашборд не працює)

**Альтернатива 2 (відкинута):** 1С → Python ETL → Power BI без проміжної SQL-бази.
- ❌ Power BI не може робити DirectQuery до Python pickle/parquet
- ❌ Imports з диска повільні для великих обсягів
- ❌ Жодного централізованого місця для звірки/аналізу

**Прийнятий 4-шар архітектури** дає:
- ✅ Кешування свёртки в 1С регістрах сведень (швидкий повторний доступ)
- ✅ Чисту розділеність бізнес-логіки (BSL) і копіювання (Python)
- ✅ SQL-базу як централізоване сховище для майбутніх використань (BI Studio, Excel pivot tables, REST endpoints)
- ✅ Power BI може працювати у режимах Import (швидко) або DirectQuery (актуально)

---

## Cross-references

- Метадані 1С (8 об'єктів): [olap_1c_objects.md](olap_1c_objects.md)
- BSL логіка свёртки: [olap_obrabotka_provedeniya.md](olap_obrabotka_provedeniya.md)
- Джерела ERP (шар 0): [olap_data_sources_erp.md](olap_data_sources_erp.md)
- SQL DDL OlapBASERP: [olap_sql_schema.md](olap_sql_schema.md)
- Python ETL: [olap_etl_pipeline.md](olap_etl_pipeline.md)
- Power BI DAX: [olap_powerbi_model.md](olap_powerbi_model.md)
- Acceptance criteria: [olap_acceptance_etalons.md](olap_acceptance_etalons.md)
- Deviations vs spec v3: [olap_deviations_from_spec.md](olap_deviations_from_spec.md)
