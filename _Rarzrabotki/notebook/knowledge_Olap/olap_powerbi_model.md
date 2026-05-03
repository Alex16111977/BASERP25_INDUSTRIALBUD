# OLAP — Power BI Model (Stage 4, ⏳ PLANNED)

> **STATUS BANNER:** ⏳ Stage 4 ще не реалізовано. Цей файл описує **запланований стан** із spec v3 §6. Реалізація буде окремим планом після ETL Stage 3.

---

## Two PBIX dashboards

| PBIX | Аудиторія | Таблиць | Зв'язків | DAX-мір | Сторінок |
|---|---|---|---|---|---|
| **Управлінський PnL BASERP25.pbix** | CFO, фінансист, керівник напряму | 14 | ~25 | ~70 | 6 |
| **Управлінський Cashflow BASERP25.pbix** | Treasury, фінансовий контролер | 9 | ~12 | ~25 | 5 |

Розташування: `_Rarzrabotki/PowerBi/`

---

## Connection Setup

### PowerQuery M

```m
let
    Source = Sql.Database("localhost", "OlapBASERP", [Query=null, CommandTimeout=#duration(0, 0, 5, 0)]),
    Fact_PnL = Source{[Schema="dbo",Item="Fact_PnL"]}[Data]
    // ... інші таблиці аналогічно
in
    Fact_PnL
```

### Authentication

⚠️ **Вкладка "База данных"** у Power BI credentials dialog (НЕ "Windows"!):
- Имя пользователя: `sa`
- Пароль: `Brw739182465!`
- Уровень: `localhost` (НЕ `localhost;OlapBASERP`)

**Якщо credentials кешуються неправильно:**
- Файл → Параметры → Параметры источника данных → знайти `localhost` → Очистить разрешения → Удалить → перепідключитись

### Modes

- **Import** (рекомендовано на старт): Power BI імпортує всі дані; refreshable через Power BI Service. Швидко працює.
- **DirectQuery**: запити шлються в SQL при кожній взаємодії. Реактивно, але повільніше.

---

## PBIX_PnL — детальна структура

### Таблиці (14)

| # | Тип | Таблиця | Зв'язки |
|---|---|---|---|
| 1 | Fact | Fact_PnL | center (ALL dim FK) |
| 2 | Dim | Dim_Organizations | 1:N → Fact_PnL.Organization_ID |
| 3 | Dim | Dim_Departments | 1:N → Fact_PnL.Department_ID |
| 4 | Dim | Dim_Directions | 1:N → Fact_PnL.Direction_ID |
| 5 | Dim | Dim_PL_Articles | 1:N → Fact_PnL.PL_Article_ID |
| 6 | Dim | Dim_PL_ArticleGroups | 1:N → Dim_PL_Articles.Group_ID (snowflake) |
| 7 | Dim | Dim_DDS_Articles | 1:N → Fact_PnL.DDS_Article_ID |
| 8 | Dim | Dim_Counterparties | 1:N → Fact_PnL.Counterparty_ID |
| 9 | Dim | Dim_Currencies | 1:N → Fact_PnL.Currency_ID |
| 10 | Dim | Dim_Income_Articles | 1:N → Fact_PnL.Income_Article_ID |
| 11 | Dim | Dim_Expense_Articles | 1:N → Fact_PnL.Expense_Article_ID |
| 12 | Util | Calendar | 1:N → Fact_PnL.Period_Month |
| 13 | Util | CFS_Sections | (для cross-PBIX consistency, якщо буде слайсер CFS у PnL) |
| 14 | Util | Table_Measures | hub для всіх ~70 measures |

### DAX міри для PnL (повний список)

#### 1. Базові ресурси через Source filter

```dax
[План Excel] = CALCULATE(SUM(Fact_PnL[Sum_Plan_Grn]), Fact_PnL[Source]="PL_Excel")
[План Ф1]    = CALCULATE(SUM(Fact_PnL[Sum_Plan_F1_Grn]), Fact_PnL[Source]="PL_Excel")
[План Ф2]    = CALCULATE(SUM(Fact_PnL[Sum_Plan_F2_Grn]), Fact_PnL[Source]="PL_Excel")

[Виручка ЕРП]  = CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_Income")
[Собівартість] = CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_CoGS")
[OpEx ЕРП]     = CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_OpEx")
[Каса]         = CALCULATE(SUM(Fact_PnL[Sum_Kazna_Grn]), Fact_PnL[Source] IN {"Казна_PL","Казна_БезPL"})
```

#### 2. Аномалії маппінгу (для діагностики)

```dax
[ЕРП без PL — Расх]  = CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_БезPL_Расх")
[ЕРП без PL — Доход] = CALCULATE(SUM(Fact_PnL[Sum_ERP_Grn]), Fact_PnL[Source]="ERP_БезPL_Доход")
```

#### 3. 5 рівнів маржі (на льоту, для будь-якого filter context)

```dax
[Валова маржа]    = [Виручка ЕРП] - [Собівартість]
[Валова маржа %]  = DIVIDE([Валова маржа], [Виручка ЕРП])

[EBITDA]          = [Валова маржа] - [OpEx ЕРП]
[EBITDA %]        = DIVIDE([EBITDA], [Виручка ЕРП])

[Чистий прибуток] = [EBITDA]   /* мінус Амортизація / Податки / Курсові — окрема логіка */
```

#### 4. План vs Факт

```dax
[План vs Факт]   = [План Excel] - ([Виручка ЕРП] + [OpEx ЕРП])
[План vs Факт %] = DIVIDE([План vs Факт], [План Excel])

[Виконання плану %] = DIVIDE(([Виручка ЕРП] + [OpEx ЕРП]), [План Excel])
```

#### 5. YTD / MTD (для Calendar drill)

```dax
[План Excel YTD]  = TOTALYTD([План Excel], Calendar[Date_Key])
[Виручка ЕРП MTD] = TOTALMTD([Виручка ЕРП], Calendar[Date_Key])
[EBITDA YTD]      = TOTALYTD([EBITDA], Calendar[Date_Key])
```

### Сторінки PnL PBIX (6)

#### 1. Зведений PnL
- Layout: великий matrix з PL-групами в рядках, місяцями у колонках
- Visualizations: matrix [Виручка ЕРП | Собівартість | Валова маржа | OpEx | EBITDA | EBITDA%]
- Слайсери (top): Період (рік/квартал/місяць), Організація, Підрозділ
- Slicer state: всі Source включені

#### 2. План vs Факт
- Layout: 2 stacked bar charts side-by-side (План | Факт), waterfall у центрі для відхилень
- Slicers: Period, PL_Article (multi-select)
- Slicer state: ВСІ Source ON, але візуалі фільтрують через CALCULATE

#### 3. Розшифровка (drill-through)
- Layout: 1 large table з рядками регістру і колонкою Source_Recorder_Url
- Drill-through page: при кліку на будь-який рядок матриці → відкривається ця сторінка з фільтром
- Колонка Source_Recorder_Url оголошена як "Web URL" → клік відкриває документ у 1С

#### 4. Аномалії маппінгу
- Layout: table з фільтром Source IN {ERP_БезPL_Расх, ERP_БезPL_Доход}
- Призначення: фінансист бачить які ДДС не прив'язані до PL-статей → йде в 1С і додає у `А_Статьи_PL.Статьи`
- Slicer: Подразділ (для іпрямлення відповідальності)

#### 5. Каса (довідково)
- Slicer state: Source IN {Казна_PL, Казна_БезPL}
- Призначення: касовий метод PnL для звірки з ДДС-звітом

#### 6. Multi-currency (планується)
- Filter: `Currency_ID <> 'UAH'`
- Призначення: показує валютні рухи у валюті оригіналу

---

## PBIX_Cashflow — детальна структура

### Таблиці (9)

| # | Тип | Таблиця | Зв'язки |
|---|---|---|---|
| 1 | Fact | Fact_Cashflow | center |
| 2 | Fact | Fact_CF_Balance | для [Reconciliation] |
| 3 | Dim | Dim_Organizations | 1:N |
| 4 | Dim | Dim_BankAccounts | 1:N |
| 5 | Dim | Dim_DDS_Articles | 1:N |
| 6 | Dim | Dim_Counterparties | 1:N |
| 7 | Util | Calendar | 1:N |
| 8 | Util | CFS_Sections | 1:N (Section_Code → Fact_Cashflow.CFS_Section) |
| 9 | Util | Table_Measures | hub |

### DAX міри для Cashflow

#### Базові

```dax
[Inflow]       = CALCULATE(SUM(Fact_Cashflow[Sum_Grn]), Fact_Cashflow[Direction]="Inflow")
[Outflow]      = CALCULATE(SUM(Fact_Cashflow[Sum_Grn]), Fact_Cashflow[Direction]="Outflow")
[Net Cashflow] = [Inflow] - [Outflow]
```

#### За джерелом (для звірки ERP vs Казна)

```dax
[Inflow ERP]      = CALCULATE([Inflow], Fact_Cashflow[Source] IN {"ERP_Безнал","ERP_Нал"})
[Inflow Казна]    = CALCULATE([Inflow], Fact_Cashflow[Source]="Казна")
[Розрив ERP-Казна] = [Inflow ERP] - [Inflow Казна]    /* для звіту звірки */
```

#### Direct Method CFS

```dax
[CFS Operating]  = CALCULATE([Net Cashflow], Fact_Cashflow[CFS_Section]="Operating")
[CFS Investing]  = CALCULATE([Net Cashflow], Fact_Cashflow[CFS_Section]="Investing")
[CFS Financing]  = CALCULATE([Net Cashflow], Fact_Cashflow[CFS_Section]="Financing")
[CFS Total]      = [CFS Operating] + [CFS Investing] + [CFS Financing]
                  /* Internal не входить — внутрішні переміщення */
```

#### Контрольна (від балансу)

```dax
[Balance Open]    = SUM(Fact_CF_Balance[Sum_Grn_Open])
[Balance Close]   = SUM(Fact_CF_Balance[Sum_Grn_Close])
[Reconciliation]  = [Balance Close] - [Balance Open] - [CFS Total]   /* має бути ≈ 0 */
```

⚠️ **Якщо [Reconciliation] != 0:** означає що:
- Або деякі рухи ДС не потрапили у регістр сведень (баг свёртки в 1С)
- Або деякі статті ДДС мають CFS_Section='Internal' (і вони не входять в CFS_Total)
- Або баланс містить курсові переоцінки які не відображаються як Cashflow

### Сторінки Cashflow PBIX (5)

#### 1. Direct Method CFS
- Layout: waterfall chart [Operating → Investing → Financing → Net]
- Slicer: Period, Організація, Source (3 значення — для toggle ERP/Казна джерел)
- Висновок: Net Cashflow за період

#### 2. Потоки за джерелами
- Layout: 3 окремих stacked column charts (по місяцях × ВидДвижения):
  - ERP_Безнал: помісячно Inflow vs Outflow
  - ERP_Нал: те саме
  - Казна: те саме
- Призначення: показати порівняння потоків за 3 джерелами

#### 3. Залишки на рахунках
- Source: Fact_CF_Balance
- Layout: matrix БанкСчет/Касса × Місяці, значення = [Balance Close]
- Drill: → клік на рахунок → бачимо Inflow/Outflow за період

#### 4. ERP vs Казна звірка
- Slicer: Period
- Layout: waterfall [Inflow ERP - Inflow Казна = Розрив]
- Призначення: виявити розбіжності між двома потоками (наприклад, документ перенесений у Казну але не закрив парний у ERP)

#### 5. Multi-currency cashflow
- Filter: `Currency_ID <> 'UAH'`
- Призначення: валютні рухи у валюті оригіналу

---

## Слайсери (загальні)

### Для PnL PBIX
- **Source** (8 значень): toggle PL_Excel, ERP_OpEx, ERP_CoGS, ERP_Income, ERP_БезPL_*, Казна_*
- **Період**: Year/Quarter/Month hierarchy з Calendar
- **Організація**: multi-select з Dim_Organizations
- **Підрозділ**: hierarchy (через Dim_Departments.Parent_ID)
- **PL Стаття**: multi-select з Dim_PL_Articles
- **Direction (бізнес-напрям)**: з Dim_Directions

### Для Cashflow PBIX
- **Source** (3): ERP_Безнал / ERP_Нал / Казна
- **CFS_Section** (4): Operating / Investing / Financing / Internal
- **Період**: як для PnL
- **Організація**, **Банківський рахунок/Каса**, **Стаття ДДС**

---

## Drill-down паттерн

### У DAX-таблиці

Колонка `Source_Recorder_Url` оголошена:
- Modeling → Properties → Data Category = **Web URL**
- У візуалі Table → Format pane → Format conditional → URL icon

### Приклад використання

User у дашборді бачить рядок «Глобино-2 / Реалізація / 38 432 968.66 ₴» → клік на іконку URL → відкривається 1С Web Client з конкретною Реалізацією. Час — instant.

---

## Acceptance Criteria

| # | Критерій | Очікуване | Як перевірити |
|---|---|---|---|
| 1 | [Reconciliation] | ≈ 0 | DAX query через `mcp__powerbi-modeling-mcp__dax_query_operations` |
| 2 | Глобино-2 / Виручка ЕРП | 38 432 968.66 ₴ | Сторінка "Зведений PnL" → фільтр Глобино-2 → візуал з [Виручка ЕРП] |
| 3 | План vs Факт за лютий 2026 | [План Excel] = 115.7M, [Виручка ЕРП] = 61.7M | Сторінка "План vs Факт" |
| 4 | Drill-down працює | Клік на рядок → відкривається 1С документ | Manual UI test |
| 5 | Слайсер Source toggling | Вкл/викл значення → візуалі реактивно перераховують | Manual UI test |

---

## Performance considerations

### Import mode (рекомендовано)

- Дані за лютий 2026: ~3937 рядків PnL + ~4652 ДДС = 8589 рядків Fact + 16 Dim з ~10000 рядків справочників. **Загалом < 20MB у Power BI VertiPaq compressed.**
- Refresh: <30 сек (одне читання SQL)
- Запити: instant (in-memory)

### DirectQuery (для real-time)

- Кожна взаємодія = SQL query → 50-500ms latency
- Потрібні правильні індекси (вже є в OlapBASERP)
- Деякі DAX features недоступні (наприклад складні measures)

---

## Cross-references

- SQL-схема джерело даних: [olap_sql_schema.md](olap_sql_schema.md)
- Python ETL що вантажить дані: [olap_etl_pipeline.md](olap_etl_pipeline.md)
- Acceptance etalons: [olap_acceptance_etalons.md](olap_acceptance_etalons.md)
- Source markers (frozen identifiers): [olap_1c_objects.md](olap_1c_objects.md)
- Spec v3 §6: `docs/superpowers/specs/2026-05-01-olap-baserp-architecture-design-v3-final.md`
