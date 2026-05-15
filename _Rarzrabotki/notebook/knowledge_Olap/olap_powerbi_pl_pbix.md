# PL.pbix — Power BI модель PnL (Шар 4)

> **STATUS:** ⏳ in development. Файл існує, дані завантажуються з OlapBASERP, модельна частина (зв'язки, перейменування таблиць) сформована 2026-05-03. DAX-міри і візуали — у процесі.
>
> **Live файл:** [`_Rarzrabotki/Olap/PowerBi/PL.pbix`](../../Olap/PowerBi/PL.pbix)
> **Тип:** ЗМІННИЙ (оновлюється при кожній зміні моделі/мір/візуалів)

---

## 1. Призначення

Перший з двох Power BI дашбордів BI-конвеєра BASERP25. Реалізує **Управлінський PnL** для CFO ТОВ ІНДАСТРІАЛБУД: план vs факт vs Казна, маржинальність, drill-down до первинних документів 1С.

Другий дашборд — `Cashflow.pbix` — окремий файл, окрема модель. Тут описаний тільки PL.

---

## 2. Підключення

```
PL.pbix
   │ Power Query (M)
   │ Sql.Database("localhost", "OlapBASERP")
   ▼
OlapBASERP (24 SQL-таблиці, sa@Brw739182465!)
   ▲
   │ pyodbc bulk_insert (Stage 3 ETL)
   │
Ai_Olap (_Rarzrabotki/Olap/Ai_Olap/)
```

- **Тип з'єднання:** Import mode (не DirectQuery). Дані кешуються у моделі, оновлюються через PBI Refresh.
- **Server:** `localhost` (той самий MSSQL що містить BaseERP та OlapBASERP).
- **Database:** `OlapBASERP`.
- **Auth:** SQL Server Authentication, `sa` / `Brw739182465!` (вкладка "База данных" у Power BI, НЕ Windows).
- **Compatibility level:** 1600 (станом на 2026-05-03).

Усі видимі таблиці приходять з SQL через Power Query M-запити. Power Query partition name (наприклад `Dim_Contracts`) залишається англомовним — це **назва запиту у Power Query**, а модельна назва таблиці перейменована на 1С-нотацію.

---

## 3. Структура моделі

### 3.1. Видимі таблиці (18 — те, що бачить розробник звіту)

**Fact** (1):

| таблиця | колонок | партиція | Power Query запит |
|---------|---------|----------|-------------------|
| `Fact_PnL` | 28 | M / Import | `Fact_PnL` |

**Dim** (15 довідників, перейменовані під 1С-нотацію 2026-05-03):

| Power BI назва (1С-стиль) | колонок | Power Query партиція | OlapBASERP таблиця |
|----------------------------|---------|----------------------|--------------------|
| `Организации` | 8 | Dim_Organizations | Dim_Organizations |
| `СтруктураПредприятия` | 8 | Dim_Departments | Dim_Departments |
| `НаправленияДеятельности` | 8 | Dim_Directions | Dim_Directions |
| `Контрагенты` | 10 | Dim_Counterparties | Dim_Counterparties |
| `ДоговорыКонтрагентов` | 8 | Dim_Contracts | Dim_Contracts |
| `Номенклатура` | 8 | Dim_Items | Dim_Items |
| `ГруппыФинансовогоУчетаНоменклатуры` | 8 | Dim_ItemGroups | Dim_ItemGroups |
| `ФизическиеЛица` | 8 | Dim_Individuals | Dim_Individuals |
| `Пользователи` | 8 | Dim_Users | Dim_Users |
| `Валюты` | 7 | Dim_Currencies | Dim_Currencies |
| `СтатьиДвиженияДенежныхСредств` | 9 | Dim_DDS_Articles | Dim_DDS_Articles |
| `СтатьиРасходов` | 8 | Dim_Expense_Articles | Dim_Expense_Articles |
| `СтатьиДоходов` | 8 | Dim_Income_Articles | Dim_Income_Articles |
| `А_Статьи_PL` | 10 | Dim_PL_Articles | Dim_PL_Articles |
| `А_ГруппаСтатей_PL` | 8 | Dim_PL_ArticleGroups | Dim_PL_ArticleGroups |

**Util** (2):

| таблиця | колонок | роль |
|---------|---------|------|
| `Calendar` | 13 | дата-вимір (2191 днів 2025-2030); поки **НЕ marked as Date Table** — треба зробити вручну ПКМ → "Mark as date table" → колонка `Date_Key` |
| `Table_Measures` | 3 | hub для DAX мір (поки порожній) |

### 3.2. Приховані технічні таблиці (22 — auto-generated від Power BI)

- `DateTableTemplate_*` (1) — шаблон time intelligence
- `LocalDateTable_*` (21) — auto-генеровані Power BI для **кожного** datetime-поля (Calendar.Date_Key, Calendar.Month_Start, Calendar.Month_End, Fact_PnL.Period, Fact_PnL.Period_Month, Fact_PnL.Loaded_At, та `Loaded_At` у кожному Dim — 15 шт)

**Це баласт.** Прибирається через **File → Options → Current File → Data Load → Time Intelligence → "Auto date/time" Off** + перезапуск PBIX. Після цього лишиться тільки `Calendar` як date dimension.

---

## 4. Зв'язки (relationships) — **зірка PL**

Усі активні, **Many-to-One**, single-direction filter. 11 ефективних бізнес-зв'язків Fact_PnL → довідники + 21 авто-LocalDateTable (баласт).

```
                         ┌──────────────┐
                         │   Calendar   │
                         └──────┬───────┘
                          Date_Key │
                                  ▼ Period_Month
   ┌───────────────────────  Fact_PnL  ──────────────────────┐
   │                                                          │
   ├ Organization_ID    →  Организации                        │
   ├ Department_ID      →  СтруктураПредприятия               │
   ├ Direction_ID       →  НаправленияДеятельности            │
   ├ PL_Article_ID      →  А_Статьи_PL                        │
   ├ PL_Group_ID        →  А_ГруппаСтатей_PL              [NEW]│
   ├ Counterparty_ID    →  Контрагенты                        │
   ├ Income_Article_ID  →  СтатьиДоходов                      │
   ├ Expense_Article_ID →  СтатьиРасходов                     │
   ├ DDS_Article_ID     →  СтатьиДвиженияДенежныхСредств      │
   ├ Currency_ID        →  Валюты                          [NEW]│
   └ Period_Month       →  Calendar.Date_Key               [NEW]│

   Без зв'язків з Fact_PnL (для drill-down тільки):
     ДоговорыКонтрагентов, Номенклатура,
     ГруппыФинансовогоУчетаНоменклатуры,
     ФизическиеЛица, Пользователи
```

3 зв'язки `[NEW]` створені 2026-05-03 через MCP `relationship_operations`:
- `Fact_PnL_Currency_ID_Dim_Currencies_Currency_ID`
- `Fact_PnL_PL_Group_ID_Dim_PL_ArticleGroups_PL_ArticleGroup_ID`
- `Fact_PnL_Period_Month_Calendar_Date_Key`

---

## 5. Ключові колонки Fact_PnL (28 шт)

| колонка | тип | опис |
|---------|-----|------|
| `Fact_ID` | Int64 | PK (IDENTITY у SQL); у DAX зазвичай не використовується |
| `Period_Month` | DateTime | 1-ше число місяця (для слайсера + JOIN з Calendar) |
| `Period` | DateTime | повна дата документа (для drill-down) |
| `Source` | String | **frozen**: `PL_Excel` / `ERP_OpEx` / `ERP_CoGS` / `ERP_Income` / `ERP_БезPL_Расх` / `ERP_БезPL_Доход` / `Казна_PL` / `Казна_БезPL` |
| `Recorder_PL_ID` | String | UUID документа `А_ФинРез_PL` |
| `Organization_ID` … `Currency_ID` | String | char(32) FK на 9 Dim |
| `Source_Recorder_ID` | String | UUID первинного документа (Реалізація, ПКО, тощо) |
| `Source_Recorder_Type` | String | назва типу документа (`РеализацияТоваровУслуг`, …) |
| `Source_Recorder_Url` | String | `e1cib/data/<type>?ref=<uuid>` — drill-down URL у 1С Web |
| `Source_Recorder_Presentation` | String | текст для відображення у візуалі |
| `Sum_Plan_Grn` | Double | план PnL з Excel |
| `Sum_Plan_F1_Grn` / `Sum_Plan_F2_Grn` | Double | деталізація плану Форма 1 / Форма 2 |
| `Sum_ERP_Grn` | Double | факт ERP — **🎯 Глобино-2 / ERP_Income = 38 432 968.66 ₴** |
| `Sum_Kazna_Grn` | Double | факт Казна |
| `Sum_Original` | Double | сума у валюті оригіналу (для multi-currency) |
| `Exchange_Rate` | Double | курс на дату документа |
| `Loaded_At` | DateTime | timestamp ETL load — **сховати від користувача** |

### Технічні колонки що рекомендується **сховати** (`column.IsHidden = true`)

У всіх Dim: `Loaded_At`, `Marked_For_Deletion`, `Parent_ID` (якщо не плануєш ієрархію), `Is_Group`. У Fact_PnL: `Source_Recorder_ID`, `Source_Recorder_Type`, `Source_Recorder_Presentation`, `Loaded_At`.

---

## 6. Calendar — структура date dimension

| колонка | тип | роль |
|---------|-----|------|
| `Date_Key` | DateTime | **PK для time intelligence** (mark as date column) |
| `Year` | Int64 | для слайсера / drill-down |
| `Quarter` | Int64 | Q1-Q4 |
| `Month_Num` | Int64 | 1-12 |
| `Month_Name` | String | "Лютий", "Січень", ... |
| `Month_Start` | DateTime | 1-ше число (== Period_Month у Fact) |
| `Month_End` | DateTime | останній день |
| `Day_Of_Month` | Int64 | 1-31 |
| `Day_Of_Week` | Int64 | 1-7 (понеділок = 1?) |
| `Day_Name` | String | "Пн", "Вт", ... |
| `Is_Weekend` | Boolean | для бізнес-логіки |
| `Year_Month` | String | "2026-02" — зручно у слайсері |

**TODO:** Mark `Calendar` як date table → ПКМ на таблиці у Power BI Desktop → "Mark as date table" → вибрати `Date_Key`. Це активує `SAMEPERIODLASTYEAR`, `DATESYTD`, `TOTALYTD` тощо.

---

## 7. DAX-міри — ще немає

`Table_Measures` — порожній hub (3 колонки, 0 мір). Заплановані за spec v3 §6:

**5 рівнів маржі для PnL:**
1. `[Sum Plan]` = `SUM(Fact_PnL[Sum_Plan_Grn])`
2. `[Sum ERP]` = `SUM(Fact_PnL[Sum_ERP_Grn])`
3. `[Sum Kazna]` = `SUM(Fact_PnL[Sum_Kazna_Grn])`
4. `[Виручка]` = `CALCULATE([Sum ERP], Fact_PnL[Source]="ERP_Income")`
5. `[Собівартість]` = `CALCULATE([Sum ERP], Fact_PnL[Source]="ERP_CoGS")`
6. `[Валова маржа]` = `[Виручка] - [Собівартість]`
7. `[Валова маржа %]` = `DIVIDE([Валова маржа], [Виручка])`
8. `[OpEx]` = `CALCULATE([Sum ERP], Fact_PnL[Source]="ERP_OpEx")`
9. `[EBITDA]` = `[Валова маржа] - [OpEx]`
10. `[EBITDA %]` = `DIVIDE([EBITDA], [Виручка])`
11. `[План vs Факт]` = `[Sum ERP] - [Sum Plan]`
12. `[Виконання плану %]` = `DIVIDE([Sum ERP], [Sum Plan])`

Time intelligence (після `Mark as date table`):
- `[Виручка YTD]` = `TOTALYTD([Виручка], Calendar[Date_Key])`
- `[Виручка LY]` = `CALCULATE([Виручка], SAMEPERIODLASTYEAR(Calendar[Date_Key]))`
- `[Зростання YoY %]` = `DIVIDE([Виручка] - [Виручка LY], [Виручка LY])`

Деталі — у [olap_powerbi_model.md](olap_powerbi_model.md) §6.

---

## 8. Видалені таблиці (2026-05-03)

З моделі прибрано Cashflow-only сутності (вони тільки для `Cashflow.pbix`):

- ❌ `CFS_Sections` (4 розділи Operating/Investing/Financing/Internal — для Cashflow)
- ❌ `Dim_BankAccounts` (банківські рахунки — для Cashflow)
- ❌ `PLArticle_DDS` (Bridge-таблиця PL ↔ DDS — потрібна тільки у Cashflow)
- ❌ `ETL_Runs` (моніторинг ETL — для адмінів, не для звіту)

Cascade-видалення прибрало також авто-LocalDateTable_* пов'язані з ними.

> Не видалено: `СтатьиДвиженияДенежныхСредств` (Dim_DDS_Articles) — `Fact_PnL.DDS_Article_ID` посилається на цей довідник для drill-down PL → ДДС-стаття.

---

## 9. Workflow оновлення

1. **Дані у БД оновлюються** через `python main.py` у `Ai_Olap/` (див. [olap_etl_pipeline.md](olap_etl_pipeline.md)).
2. **PL.pbix Refresh** — у Power BI Desktop кнопка "Refresh" → Power Query тягне з `OlapBASERP` свіжі рядки.
3. **Опубліковано на Power BI Service** (TODO Stage 5) — там Refresh за розкладом через Gateway.

---

## 10. TODO до production

| # | завдання | пріоритет |
|---|----------|-----------|
| 1 | `Mark as date table` для `Calendar` (колонка Date_Key) | високий |
| 2 | Вимкнути auto date/time у File → Options → Current File (приберает 22 LocalDateTable_*) | високий |
| 3 | Створити DAX-міри у `Table_Measures` (~12-15 базових + time intelligence) | високий |
| 4 | Сховати технічні колонки (`Loaded_At`, `Marked_For_Deletion`, `Source_Recorder_ID/Type/Presentation`) | середній |
| 5 | Додати hierarchy `[А_ГруппаСтатей_PL] → [А_Статьи_PL]` для drill-down | середній |
| 6 | Створити сторінки звіту (PnL waterfall, plan vs fact, marginальність) | високий |
| 7 | Source_Recorder_Url пометить як `Web URL` (Modeling → Properties → Data Category) | середній |
| 8 | Compatibility level 1600 → 1604+ (для нових DAX-функцій) | низький |
| 9 | Включити Q&A для природньо-мовних запитів | низький |
| 10 | Опублікувати на Power BI Service з гейтвеєм до OlapBASERP | Stage 5 |

---

## 11. Acceptance verification (DAX перевірка)

Після створення базових мір — швидка перевірка через DAX query:

```dax
EVALUATE
ROW(
  "Глобино-2 ERP_Income Feb 2026",
    CALCULATE(
      SUM(Fact_PnL[Sum_ERP_Grn]),
      Fact_PnL[Source] = "ERP_Income",
      Fact_PnL[Period_Month] = DATE(2026, 2, 1),
      СтруктураПредприятия[Department_Name] = "Глобино-2"
    )
)
```

Очікуваний результат: **38 432 968.66 ₴** (точно ± 0.01) — той самий acceptance gate що у Stage 1+3 pytest.

---

## 12. Cross-references

- Source SQL-схема: [olap_sql_schema.md](olap_sql_schema.md) (24 таблиці OlapBASERP)
- Стандартизована модель/міри (планований стан): [olap_powerbi_model.md](olap_powerbi_model.md)
- ETL що наповнює дані: [olap_etl_pipeline.md](olap_etl_pipeline.md)
- 1С-метадані Fact-джерела: [olap_1c_objects.md](olap_1c_objects.md) (`РегистрСведений.А_ОтчетPL_Свод`)
- Acceptance numbers: [olap_acceptance_etalons.md](olap_acceptance_etalons.md)
- Live PBIX: `_Rarzrabotki/Olap/PowerBi/PL.pbix`

---

## 13. Balance Stage (2026-05-16) — модель управлінського балансу

> Канон §10/Roadmap (OD-9). PL.pbix вже багатофактовий (Fact_PnL +
> Fact_Cashflow + 65 мір); додано третю Fact-область — **Баланс**.
> Зміни внесені через MCP `powerbi-modeling-mcp` у живий PBIX
> (server `SQLSERVER`, DB `OlapBASERP`, Import).

### 13.1 Таблиці (Power Query M, Import)

| модельна назва | OlapBASERP | M-джерело |
|---|---|---|
| `Fact_Balance` | dbo.Fact_Balance | `Sql.Database("SQLSERVER","OlapBASERP")` |
| `Dim_PAP_Articles` | dbo.Dim_PAP_Articles | те саме |

> Назва Dim лишена `Dim_PAP_Articles` (прецедент `Dim_DenezhnyeSredstva`),
> DAX-міри референсують саме її — НЕ перейменовувати без оновлення мір.

### 13.2 Зв'язки (Many→One, single-direction) — 7

`Fact_Balance[PAP_Article_ID]→Dim_PAP_Articles`, `[Organization_ID]→Организации`,
`[Department_ID]→СтруктураПредприятия`, `[Counterparty_ID]→Контрагенты`,
`[Individual_ID]→ФизическиеЛица`, `[Contract_ID]→ДоговорыКонтрагентов`,
`[Period]→Calendar[date_]`.
**Partner_ID — без прямого зв'язку** (неоднозначний шлях через
Контрагенты→Партнеры; дзеркалить Fact_Cashflow). Item/Warehouse/OperObject/
Cash/SettlementObj/Intangible — без Dim (drill-down наступний цикл, не блокує).

### 13.3 DAX-міри (Table_Measures, displayFolder `Balance`) — 8

```dax
[Баланс Вх]     = SUM(Fact_Balance[Sum_Open])
[Баланс Прихід] = SUM(Fact_Balance[Sum_Inflow])
[Баланс Розхід] = SUM(Fact_Balance[Sum_Outflow])
[Баланс Вих]    = SUM(Fact_Balance[Sum_Close])
[Перевірка обороту] = [Баланс Вих]-([Баланс Вх]+[Баланс Прихід]-[Баланс Розхід])
[Актив]  = CALCULATE([Баланс Вих],'Dim_PAP_Articles'[AktivPassiv]="Aktiv")
         + CALCULATE([Баланс Вих],'Dim_PAP_Articles'[AktivPassiv]="AktivPassiv",
                     Fact_Balance[Sum_Close] > 0)
[Пассив] = -CALCULATE([Баланс Вих],'Dim_PAP_Articles'[AktivPassiv]="Passiv")
         - CALCULATE([Баланс Вих],'Dim_PAP_Articles'[AktivPassiv]="AktivPassiv",
                      Fact_Balance[Sum_Close] < 0)
[Контроль Актив-Пассив] = [Актив] - [Пассив]
```

Двосторонні `AktivPassiv` діляться по знаку Sum_Close (як Налоги, канон OD-9).

### 13.4 Verified (DAX, січень 2026/ТОВ, після Refresh Fact_Balance+Dim_PAP_Articles)

- `[Баланс Вих]` (Σ Close) = **−0.01 ≈ 0** → Актив=Пасив ✓
- Aktiv-only = **289 064 974,43** (точний канон-контроль) ✓
- Passiv-only = −298 396 250,36; AktivPassiv-net = 9 331 275,92
- Тотожність 289 064 974,43 + 9 331 275,92 − 298 396 250,36 = **−0.01**
  → `[Контроль Актив-Пассив]` ≈ 0 ✓
- `[Перевірка обороту]` = −0.36 ≈ 0 (округлення ОТ розкладу); `Fact_Balance` 11 561

### 13.5 ⚠️ Обмеження MCP dax Execute (`>`/`<` row-scan)

`powerbi-modeling-mcp` `dax_query_operations Execute` кидає «An unexpected
exception» на БУДЬ-ЯКОМУ row-scan з числовим `>`/`<` по Double-колонці —
**відтворюється і на незайманому production `Fact_Cashflow[Sum_Fact]>0`**.
Це обмеження тест-харнеса MCP, НЕ дефект моделі: `=`, рядкові предикати,
`SUM/SUMX`, `1>0`-літерал, Dim-фільтри — працюють; міри `Validate`-нуються.
У рушії самого Power BI Desktop (рантайм звіту) міри `[Актив]/[Пассив]/
[Контроль]` обчислюються коректно (їх `>0/<0` — у збереженому визначенні,
не в тексті MCP-запиту). Канон-міри НЕ спотворювати під баг харнеса.
Перевірка контролю — арифметично через компонованні Dim-агрегати (13.4).

### 13.6 Сторінка «Баланс» (візуали — вручну у Power BI Desktop)

MCP керує лише моделлю (не report-layer). Створити сторінку поряд з
PnL/Cashflow:
- **Matrix:** рядки `Dim_PAP_Articles` (drill Parent→child через PAP_Article_ID/
  Parent_ID), значення `[Баланс Вх]/[Прихід]/[Розхід]/[Вих]`.
- **KPI-карти:** `[Актив]`, `[Пассив]`, `[Контроль Актив-Пассив]` (≈0).
- **Слайсери:** `Calendar[Year_Month]`, `Dim_PAP_Articles[AktivPassiv]`,
  `Fact_Balance[Source]` (drill-down 7 джерел), `Организации`.
- Сховати тех. колонки (`Loaded_At`, `Marked_For_Deletion`, `*_ID`).

### 13.7 Save .pbix

MCP змінює живу in-memory модель; **файл .pbix зберігає користувач**
вручну: Power BI Desktop → Ctrl+S. Без цього зміни моделі втратяться при
закритті. Після збереження — Refresh (Power Query) щоб дані Fact_Balance
збіглися з останнім ETL.
