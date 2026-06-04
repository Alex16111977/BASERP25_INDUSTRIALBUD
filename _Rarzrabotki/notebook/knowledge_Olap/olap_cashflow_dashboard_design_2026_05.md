# Cashflow Dashboard Design (PL.pbix — Stage v3.8)

**Status:** ✅ Model layer (DAX measures + calculated columns) готово 2026-05-11. Visuals/page layout — manual у Power BI Desktop UI (інструкції нижче).

## Implementation summary (2026-05-11)

- ✅ `Calendar` marked as Date Table (`date_` column) — активує SAMEPERIODLASTYEAR, DATESINPERIOD, DATESYTD
- ✅ Calculated column `СтатьиДвиженияДенежныхСредств[CFS_Section_Proxy]` створена — Operating=127, Investing=8, Financing=33, BLANK=258
- ✅ 13 нових DAX measures у Table_Measures (41 existing → 54 total):

| Display Folder | Measure | Use case |
|---|---|---|
| Cashflow\CFS Standard | `CFO Operating`, `CFI Investing`, `CFF Financing`, `Free Cash Flow` | Page 5 — IFRS-style 3-section CFS |
| Cashflow\Cash Position | `Cumulative Cash Position` | Page 4 — running total balance |
| Cashflow\Executive KPI | `Net Cashflow PY`, `Net Cashflow YoY %`, `Avg Daily Outflow 30d`, `Days of Cash Remaining`, `Funding Gap Months` | Page 1 — CFO 5-sec KPI cards |
| Cashflow\Project Tracking | `Days to Depletion`, `Project Status` | Page 3 — project burn rate |
| Cashflow\Variance | `Variance Threshold Flag` | Page 2 — green/amber/red conditional |

**Verified values:**
- CFO Operating = **672 397 148.48 ₴**
- CFI Investing = **8 595 182.48 ₴**
- CFF Financing = **245 362 258.39 ₴**
- Free Cash Flow = **680 992 330.96 ₴**

**Acceptance gate (без зміни):** Globyno-2 Feb 2026 Income (PL_ЕРП) = **38 432 968.66 ₴** ✓ збережено.

---

## Manual UI steps (Power BI Desktop)

Через MCP створено **semantic model** (measures + columns + date table). **Visuals і layout — створити вручну** у Power BI Desktop. Нижче — рецепт для кожної сторінки.

### Page 1 — Executive Overview

1. Створити нову сторінку "Executive Overview"
2. Вгорі (top-left): **Slicer** Calendar[Period_Month] (тип: Dropdown або Tile)
3. **Row 1 — KPI cards** (6 card visuals в ряд):
   - `Net Cashflow` · `% Виконання плану ДДС факт (місяць)` · `Cumulative Cash Position`
   - `Days of Cash Remaining` · `Funding Gap Months` · `Net Cashflow YoY %`
   - На кожному card → Format → Conditional formatting → Background color: Green if positive, Red if negative
4. **Row 2 — Matrix 1** (Department × [Нал/Безнал/Total]):
   - Rows: `СтруктураПредприятия[Подразделение]`
   - Columns: `Dim_DenezhnyeSredstva[Account_Type]` (Bank/Cash) — потрібен relationship з Fact_Cashflow
   - Values: `Net Cashflow`
5. **Row 2 — Matrix 2** (Доходи vs Списання):
   - Rows: `СтруктураПредприятия[Подразделение]`
   - Values: `Поступления ДДС`, `Списания ДДС`, `Net Cashflow`
   - Conditional formatting на `Net Cashflow`: green/amber/red за `Variance Threshold Flag`

### Page 2 — Department Drill (Monthly Variance)

1. **Slicer** `СтруктураПредприятия[Подразделение]` — single select
2. **Slicer** Calendar period (range)
3. **Field Parameter** "Show Metric" (Modeling → New parameter → Fields):
   - Включити: `Сума план ДДС (місяць)`, `Сума факт ДДС`, `Відхилення план-факт (місяць)`, `% Виконання плану ДДС факт (місяць)`
4. **Matrix** (центральна):
   - Rows: `СтатьиДвиженияДенежныхСредств[ИерархияСтатейДДС]` (drill-down hierarchy)
   - Columns: `Calendar[year_month]`
   - Values: [Show Metric] field parameter
   - Filter `Видимість рядка ДДС план-факт (місяць)` is not blank
5. **Decomposition Tree** (Insert → AI visuals → Decomposition tree):
   - Analyze: `Net Cashflow`
   - Explain by: Department, DDS_Article_Name, Counterparty_Name, Direction
6. **Drill-through page setup:**
   - Page settings → Drill-through → add `Source_Recorder_ID` field
   - Кнопка/cell клік → переходить на Detail page з 1С URL

### Page 3 — Project Tracking (Object Plan)

1. **Slicer** `СтруктураПредприятия[Подразделение]` (multi-select)
2. **Filter on page:** `Має план ДДС (об'єкт)` is not blank — лише проекти з активним object planом
3. **KPI cards** (4 шт):
   - `Сума план ДДС (об'єкт)` (заплановано)
   - `Сума факт ДДС` (потрачено)
   - `% Виконання плану ДДС факт (об'єкт)` — burn rate
   - `Days to Depletion`
4. **Bar chart top critical projects:**
   - Y-axis: Department_Name
   - X-axis: `Days to Depletion` (sort ASC)
   - Color: `Project Status` (Critical=red, Over=amber, OK=green)
   - Top N filter: 10
5. **Detail Matrix:**
   - Rows: Department × DDS Hierarchy
   - Values: `Сума план ДДС (об'єкт)`, `Сума факт ДДС`, `Відхилення план-факт (об'єкт)`, `% Виконання плану ДДС факт (об'єкт)`, `Project Status`

### Page 4 — Cash Position

> **Обмеження:** Fact_CF_Balance таблиця не доданий до PBI Power Query (тільки SQL DDL існує). Використовуємо `Cumulative Cash Position` через Net Cashflow як approximation. Для повного opening/closing balance — додати Power Query connector до Fact_CF_Balance в наступній ітерації.

1. **Slicer** Calendar period
2. **Slicer** `Dim_DenezhnyeSredstva` (Bank/Cash filter)
3. **Card cluster:**
   - Net Cashflow · Cumulative Cash Position · Days of Cash Remaining · Funding Gap Months
4. **Waterfall chart** (Insert → Waterfall):
   - Category: Calendar[year_month]
   - Y values: `Net Cashflow`
   - Initial: показує bridging між початком і кінцем періоду
5. **Time series Line chart:**
   - Axis: Calendar[date_]
   - Values: `Cumulative Cash Position`
   - Reference line: 0 (визначає коли позиція стає негативною)
6. **Funding Gap matrix:**
   - Rows: Calendar[year_month]
   - Values: `Net Cashflow`
   - Conditional formatting: red if < 0
   - Filter: `Net Cashflow` < 0

### Page 5 — CFS Standard (IFRS Direct Method)

1. **Slicer** Calendar[year_month] (range)
2. **KPI cards (4):**
   - `CFO Operating` · `CFI Investing` · `CFF Financing` · `Free Cash Flow`
3. **Sankey chart** (Insert → Get more visuals → Sankey by Microsoft):
   - Source: `СтатьиДвиженияДенежныхСредств[CFS_Section_Proxy]`
   - Destination: `СтатьиДвиженияДенежныхСредств[Level3]`
   - Weights: `Net Cashflow` (absolute value)
4. **Matrix CFS by Month:**
   - Rows: `СтатьиДвиженияДенежныхСредств[CFS_Section_Proxy]` → `[Level3]`
   - Columns: Calendar[year_month]
   - Values: `Net Cashflow`
5. **Trend bar chart:**
   - X-axis: Calendar[year_month]
   - Stacked bar: CFO/CFI/CFF
   - Line overlay: Free Cash Flow

### Cleanup

- Видалити дублікати pages "Дублікат Cashflow обєкт (казна)" і "Дублікат Дублікат..." після створення нових Pages 1-5.
- Перейменувати у моделі: переконатись що Fact_Cashflow має активне зв'язок з Calendar по `date_` (single direction Calendar→Fact).

> 5-сторінковий Power BI Cashflow dashboard для CFO ТОВ ІНДАСТРІАЛБУД. Дизайн обґрунтований індустрійним best-practice research (CFO 5-sec rule, variance thresholds, direct method).

---

## 1. Архітектурні рішення

| Рішення | Вибір | Обґрунтування |
|---|---|---|
| Plan horizon | **Окремі сторінки** для `ПланОбъекта` і `План` (місяць) | Кожна сторінка відповідає одному запиту: project tracking vs monthly variance |
| Page count | **5** прицільних | CFO 5-sec rule + 4 спеціалізовані drill |
| CFS Section | **Level1 proxy через DAX** | Реквізит `А_РазділCFS` у 1С NULL для всіх 425 статей. Mapping: `Операционная`→Operating, `Инвестиционная`→Investing, `Финансовая`→Financing |
| Method | **Direct method** | Fact_Cashflow зберігає cash transactions з джерела; IFRS preferred для management reporting |

## 2. 5 Pages overview

### Page 1 — Executive Overview (CFO 5-sec rule)
- 6 KPI cards (no charts, no tables): Net Cashflow MTD, % Plan exec, Closing balance, Days of Cash Remaining, Funding Gap Months, YoY %
- 2 supporting matrix: Department × [Нал/Безнал/Total]; Доходи vs Списання
- Conditional: green<5%, amber 5-10%, red>10% (industry standard threshold)

### Page 2 — Department Monthly Drill
- Slicer: Department, Calendar period
- Matrix: DDS Hierarchy (rows) × Months (cols), field parameter swap [Plan/Fact/Variance/%]
- Decomposition Tree (AI variance breakdown)
- Drill-through до 1С документа через `Source_Recorder_Url`

### Page 3 — Project Tracking (Object Plan Burn Rate)
- No time axis — фокус на lifetime budget
- KPI: % Burned, % Time Elapsed, Burn Rate Variance
- Bar chart top-10 critical projects (sorted by Days to depletion ASC)
- Detail matrix: Department × [Plan/Spent/Remaining/% Done/Status]

### Page 4 — Cash Position (Treasury Reconciliation)
- **Waterfall**: Opening → Inflows → Outflows → Closing (Fact_CF_Balance)
- Time series: balance trend per account, Bank vs Cash split
- Reconciliation check: `[Close]-[Open] = [Inflow]-[Outflow]` (acceptance ±0.01 ₴)
- Funding Gap flag table

### Page 5 — CFS Standard (Direct method, Level1 proxy)
- **Sankey**: Net Inflow → Operating/Investing/Financing → подстатті
- Matrix: 3 CFS sections × Months (period over period)
- KPI: CFO, CFI, CFF, Free Cash Flow

## 3. DAX measures inventory

### Existing (Stage v3.x)
- `Сума факт ДДС`, `Поступління ДДС`, `Списання ДДС`, `Сума казна ДДС`
- `Сума план ДДС (місяць)`, `Сума план ДДС (об'єкт)`
- `Відхилення план-факт/казна (місяць/об'єкт)`
- `% Виконання плану ДДС факт/казна (місяць/об'єкт)`
- `Net Cashflow`

### NEW (Stage v3.8 implementation)

**Page 1 (Executive):**
- `Closing Balance MTD` = LASTNONBLANK(Calendar, [Sum_Grn_Close])
- `Opening Balance MTD` = FIRSTNONBLANK(Calendar, [Sum_Grn_Open])
- `Days of Cash Remaining` = DIVIDE([Closing Balance MTD], [Avg Daily Outflow 30d])
- `Avg Daily Outflow 30d` = AVERAGEX(DATESINPERIOD(Calendar[Date_Key], LASTDATE(Calendar[Date_Key]), -30, DAY), [Списання ДДС] / 30)
- `Funding Gap Months` = COUNTROWS(FILTER(VALUES(Calendar[Year_Month]), [Net Cashflow] < 0))
- `Net Cashflow MTD YoY %` = DIVIDE([Net Cashflow] - [Net Cashflow PY], ABS([Net Cashflow PY]))
- `Net Cashflow PY` = CALCULATE([Net Cashflow], SAMEPERIODLASTYEAR(Calendar[Date_Key]))

**Page 2 (Variance flag):**
- `Variance Threshold Flag` = SWITCH(TRUE(), ABS([% Виконання плану ДДС факт (місяць)] - 1) <= 0.05, "Green", ABS(...) <= 0.10, "Amber", "Red")

**Page 3 (Project Tracking):**
- `% Burned (Object)` = DIVIDE([Сума факт ДДС cumulative], [Сума план ДДС (об'єкт)])
- `Сума факт ДДС cumulative` = CALCULATE([Сума факт ДДС], DATESBETWEEN(Calendar[Date_Key], DATE(2024,1,1), LASTDATE(Calendar[Date_Key])))
- `Burn Rate Variance` = [% Burned (Object)] - [% Time Elapsed (Project)]
- `Days to Depletion` = DIVIDE([Сума план ДДС (об'єкт)] - [Сума факт ДДС cumulative], [Avg Daily Outflow 30d])
- `Project Status` = SWITCH(TRUE(), [% Burned (Object)] > 1.10, "Critical", [% Burned (Object)] > 1.00, "Over", [Burn Rate Variance] > 0.10, "Warning", "OK")

**Page 4 (Cash Position):**
- `Opening Balance` = SUM(Fact_CF_Balance[Sum_Grn_Open])
- `Closing Balance` = SUM(Fact_CF_Balance[Sum_Grn_Close])
- `Total Inflow CFB` = SUM(Fact_CF_Balance[Sum_Grn_Inflow])
- `Total Outflow CFB` = SUM(Fact_CF_Balance[Sum_Grn_Outflow])
- `Reconciliation Check` = ([Closing Balance] - [Opening Balance]) - ([Total Inflow CFB] - [Total Outflow CFB])

**Page 5 (CFS Standard):**
- `CFO Operating` = CALCULATE([Net Cashflow], 'СтатьиДвиженияДенежныхСредств'[CFS_Section_Proxy] = "Operating")
- `CFI Investing` = CALCULATE([Net Cashflow], 'СтатьиДвиженияДенежныхСредств'[CFS_Section_Proxy] = "Investing")
- `CFF Financing` = CALCULATE([Net Cashflow], 'СтатьиДвиженияДенежныхСредств'[CFS_Section_Proxy] = "Financing")
- `Free Cash Flow` = [CFO Operating] + [CFI Investing]

## 4. Calculated columns

**Dim_DDS_Articles (PBI side, не SQL):**
```dax
CFS_Section_Proxy = 
SWITCH(
    TRUE(),
    CONTAINSSTRING([Level2], "Операционная"), "Operating",
    CONTAINSSTRING([Level2], "Инвестиционная"), "Investing",
    CONTAINSSTRING([Level2], "Финансовая"), "Financing",
    CONTAINSSTRING([Level2], "Внутренние"), "Internal",
    BLANK()
)
```
**Чому Level2 а не Level1:** З спостережень структури — Level1 = "Доходы"/"Списание" (тип руху), Level2 = "Операционная"/"Инвестиционная"/"Финансовая" (CFS класифікація).

## 5. Calendar setup

**Mark Calendar as Date Table:** `model_operations.MarkAsDateTable` з `Date_Key` колонкою. Це активує time intelligence DAX (SAMEPERIODLASTYEAR, DATESYTD, тощо).

## 6. Implementation steps

1. ✅ Verify PBI connection (existing session)
2. Create CFS_Section_Proxy calculated column on Dim_DDS_Articles via `column_operations.Create`
3. Mark Calendar as date table
4. Create ~25 new DAX measures у Table_Measures via `measure_operations.Create`
5. **MANUAL у PBI Desktop UI:** створити 5 pages + visuals (MCP не редагує UI layout)
6. Test acceptance: `Globyno-2 Feb 2026 Income PL_ЕРП = 38 432 968.66 ₴` без зміни

## 7. Limitations

- **Page visuals layout** — створюються вручну у Power BI Desktop UI. MCP керує лише semantic model (measures, columns, relationships, hierarchies).
- **Conditional formatting** — налаштовується через UI на кожному visual.
- **Drill-through actions** — налаштовуються через UI.
- **Bookmarks, mobile layout** — UI only.

Через MCP я налаштую всі необхідні DAX measures + calculated columns + relationships. Користувач потім використовує ці measures у візуалах PBI Desktop.

## Sources

- [Power BI Financial Dashboard Best Practices 2026 — EPC Group](https://www.epcgroup.net/power-bi-dashboard-design-best-practices-enterprise-2026)
- [Direct vs Indirect Cash Flow — HBS](https://online.hbs.edu/blog/post/how-to-prepare-a-cash-flow-statement)
- [Power BI Variance Analysis — Zebra BI](https://zebrabi.com/power-bi-variance-reports/)
- [Project Burn Rate — Projectworks](https://www.projectworks.com/blog/the-importance-of-burn-rate-how-to-calculate-it)
- [Cash Flow Spreadsheet Construction — Mastt](https://www.mastt.com/resources/construction-project-cash-flow-spreadsheet)
- [Sankey for income statement — SankeyArt](https://www.sankeyart.com/content/blog/why-a-sankey-diagram-is-the-best-way-to-visualize-an-income-statement/)
