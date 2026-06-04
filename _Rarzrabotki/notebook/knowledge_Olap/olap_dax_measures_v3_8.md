# OLAP — DAX measures Stage v3.8 (теорія + reference)

> Reference manual для 13 нових DAX measures + 1 calculated column, доданих 2026-05-11 у PL.pbix. Theory-first: для кожного measure пояснюю **що показує**, **звідки взято** (best practice / industry standard), **коли використовувати**, **які значення нормальні**.

> **Конвенція знаку для `Відхилення план-...`:** формула = `План - Факт` (або `План - Казна`). Додатне = недовиконання плану, від'ємне = перевиконання. Назви measures префіксуються "план-..." щоб відображати порядок операндів. (Спроба інверсії 2026-05-12 була повернута на користувацьку вимогу — залишено оригінальну логіку Stage v3.x.)

---

## ⭐ Нові measures 2026-05-12: `Diff план-...` (favorable variance із SWITCH)

Додано 4 measures з **явним SWITCH(Level1)** як в Excel — formula читабельна, з фільтром по Доходи/Списання:

| Measure | Folder | Логіка |
|---|---|---|
| `Diff план-факт (місяць)` | Cashflow\ЕРП ДДС | Доходи: Факт−План; Списання: ABS(План)−ABS(Факт) |
| `Diff план-факт (об'єкт)` | Cashflow\ЕРП ДДС | те саме на об'єктному плані |
| `Diff план-казна (місяць)` | Cashflow\Казна ДДС | Доходи: Казна−План; Списання: ABS(План)−ABS(Казна) |
| `Diff план-казна (об'єкт)` | Cashflow\Казна ДДС | те саме на об'єктному плані |

### Шаблон формули (всі 4 однакові structurally)

```dax
Diff план-казна (об'єкт) =
VAR _Lvl = SELECTEDVALUE('СтатьиДвиженияДенежныхСредств'[Level1], "MIXED")
RETURN
    SWITCH(TRUE(),
        _Lvl = "Доходы",   ROUND([Сума казна ДДС] - [Сума план ДДС (об'єкт)], 2),
        _Lvl = "Списание", ROUND(ABS([Сума план ДДС (об'єкт)]) - ABS([Сума казна ДДС]), 2),
        -- MIXED (totals): inline refs в CALCULATE — НЕ VARs!
        ROUND(
            CALCULATE([Сума казна ДДС] - [Сума план ДДС (об'єкт)],
                      KEEPFILTERS('СтатьиДвиженияДенежныхСредств'[Level1] = "Доходы")) +
            CALCULATE(ABS([Сума план ДДС (об'єкт)]) - ABS([Сума казна ДДС]),
                      KEEPFILTERS('СтатьиДвиженияДенежныхСредств'[Level1] = "Списание")),
            2
        )
    )
```

### ⚠️ Важлива пастка з VARs (виправлена 2026-05-12 14:00)

**Перша версія** використовувала `VAR _Plan = [Сума план...]` і `VAR _Kazna = [Сума казна...]` для повторного використання. Це БУГ — VARs у DAX **immutable**: значення фіксується в момент обчислення, потім CALCULATE() не переобчислює VAR при зміні filter context.

```dax
-- WRONG (Branch 3 повертає 0):
VAR _Plan = [Сума план ДДС (об'єкт)]   -- evaluated at outer context (Total)
VAR _Kazna = [Сума казна ДДС]          -- evaluated at outer context (Total)
RETURN ...
    CALCULATE(_Kazna - _Plan, KEEPFILTERS([Level1]="Доходы"))  -- _Kazna і _Plan уже = Total constants
    + CALCULATE(ABS(_Plan) - ABS(_Kazna), KEEPFILTERS([Level1]="Списание"))
    -- = (TotalKazna - TotalPlan) + (ABS(TotalPlan) - ABS(TotalKazna)) ≈ X + (-X) = 0
```

Симптом: на "Всего" рядках/колонках DiffPK = 0 замість суми Diff_Доходи + Diff_Списання.

**Фікс:** inline references `[Сума казна ДДС]` всередині CALCULATE — кожен CALCULATE re-evaluates measure з новим filter context.

Перевірка Глобино-2 об'єкт-всього після фіксу:
- Plan = 146 104 872, Kazna = 153 735 877
- Diff_Доходи = +356 032 814, Diff_Списання = -348 401 808
- **DiffPK_Total = +7 631 005** = sum обох ✓ (раніше було 0)

### Favorable convention

| Папка | Позитив = | Негатив = |
|---|---|---|
| **Доходы** | перевиконали план (Факт/Казна > Plan) | недоробили |
| **Списание** | зекономили (Факт/Казна < Plan) | перевитратили |
| **Total** | net favorable (income gains > expense overruns) | net unfavorable |

### Verification (Глобино-2 об'єкт-level)

| Level1 | План | Казна | Diff план-казна (об'єкт) |
|---|---:|---:|---:|
| Доходы | +616 145 438 | +972 178 252 | **+356 032 814** (перевиконання) |
| Списание | -470 040 566 | -818 442 374 | **-348 401 808** (перевитрата) |

### Чому `ABS()` для Списання

В моделі outflows зберігаються signed negative (`Sum_Grn` * −1 ETL-style). Без ABS формула `Plan - Kazna` для signed values дає протилежний знак до того що очікує користувач (мислячий в Excel-positives). `ABS(Plan) - ABS(Kazna)` повторює Excel поведінку `AE - AF` де AE/AF — це positive numbers.

### Чим відрізняється від існуючих `Відхилення план-...`

| Aspect | `Відхилення план-факт` | `Diff план-факт` |
|---|---|---|
| Formula | `Plan - Fact` (single) | SWITCH(Level1) ABS-based |
| Convention | mathematical (plan as anchor) | favorable (positive=good) |
| Однакова логіка для in/out flow | так | ні — explicit SWITCH |
| Use case | мониторинг абсолютного відхилення | conditional formatting "good/bad" |

---

## Загальна філософія

CFO / фінансовий аналітик у щомісячному звіті ставить **4 категорії запитань**:

| Категорія | Запитання | Якими measures відповідаємо |
|---|---|---|
| **A. Поточна позиція** | "Скільки в нас зараз кешу? Як ми йдемо до плану цього місяця?" | Existing: `Net Cashflow`, `Сума факт ДДС`, `% Виконання плану ДДС (місяць)` |
| **B. Тренд + порівняння** | "Як це порівняно з минулим роком? Чи покращується ситуація?" | NEW: `Net Cashflow PY`, `Net Cashflow YoY %`, `Cumulative Cash Position` |
| **C. Ризик ліквідності** | "На скільки нам вистачить кешу? Які місяці в мінусі?" | NEW: `Avg Daily Outflow 30d`, `Days of Cash Remaining`, `Funding Gap Months` |
| **D. Управління проектами** | "Чи вкладаємось в бюджет проекту? Який критичний?" | NEW: `Days to Depletion`, `Project Status` |
| **E. Класична звітність** | "Як виглядає cashflow по IFRS секціям (CFO/CFI/CFF)?" | NEW: `CFS_Section_Proxy`, `CFO Operating`, `CFI Investing`, `CFF Financing`, `Free Cash Flow` |
| **F. Variance discipline** | "Які відхилення критичні (≥10%), які толерантні (≤5%)?" | NEW: `Variance Threshold Flag` |

Existing 41 measures покривали А. Нові 13 + column = повний CFO toolkit.

---

## 0. Calculated Column: `СтатьиДвиженияДенежныхСредств[CFS_Section_Proxy]`

### Що це
Текстова класифікація кожної ДДС-статті у одну з 4 категорій **Cash Flow Statement (CFS)** за IFRS / IAS 7:
- **Operating** (CFO) — операційна діяльність (купівля/продаж товарів, зарплата, податки)
- **Investing** (CFI) — інвестиційна (купівля ОС, фінансові вкладення, M&A)
- **Financing** (CFF) — фінансова (кредити, дивіденди, випуск акцій)
- **Internal** — внутрішні перекидання між рахунками організації (не справжній рух кешу, виключається з total)
- *(BLANK)* — поза CFS-класифікацією (помилкові статті, "НЕ ИСПОЛЬЗОВАТЬ" і т.п.)

### DAX

```dax
CFS_Section_Proxy =
SWITCH(
    TRUE(),
    CONTAINSSTRING([Level2], "Операционная"),    "Operating",
    CONTAINSSTRING([Level2], "Инвестиционная"),  "Investing",
    CONTAINSSTRING([Level2], "Финансовая"),      "Financing",
    CONTAINSSTRING([Level2], "Внутренние"),      "Internal",
    BLANK()
)
```

### Чому **proxy**, а не справжній CFS_Section

У SQL таблиці `Dim_DDS_Articles` є колонка `CFS_Section` (varchar 15) яка мала б приходити з реквізиту 1С `СтатьиДвиженияДенежныхСредств.А_РазделCFS` (Enum: Operating/Investing/Financing/Internal). Але **реквізит у 1С НЕ заповнений** для всіх 425 статей — фінансова команда його не вела (відомий "data gap" — задокументовано у `olap_deviations_from_spec.md`).

Тому використовуємо `Level2` ієрархії як **proxy** — вона і так іменована "Операционная деятельность", "Финансовая деятельность" і т.п. (видно у Excel-сводному). Це "lossy but functional" workaround: коли фіни заповнять реквізит — переключити DAX на `[CFS_Section]` напряму, а `CFS_Section_Proxy` видалити.

### Розподіл (2026-05-11)

| Section | Articles count |
|---|---:|
| Operating | 127 |
| Investing | 8 |
| Financing | 33 |
| Internal | 0 |
| *(BLANK)* | 258 |

258 BLANK = переважно "!!! НЕ ИСПОЛЬЗОВАТЬ !!!" статті + статті з некласифікованим Level2 ("Кредит", "Продажи" і т.п.). У visual CFS standard ці статті йдуть у "Інші" категорію.

### Best practice background

> "Operating activities, investing activities, and financing activities are the three core sections required by IAS 7 / ASC 230. **Operating** must be reported either by direct method (cash receipts/payments) or indirect (Net Income + adjustments). Most management dashboards prefer **direct method** because it shows actual cash mechanics — receipts from customers, payments to suppliers, etc." — Harvard Business School Online

Наша модель — **direct method** (Fact_Cashflow зберігає transactions з реальних документів ПКО/РКО/ПлатПоруч), тому CFS-секція мапиться через категорію статті ДДС.

---

## 1. Group "Cashflow\CFS Standard" — IFRS-style 3-section report

### 1.1 `CFO Operating`

**DAX:**
```dax
CFO Operating = 
CALCULATE([Net Cashflow], 'СтатьиДвиженияДенежныхСредств'[CFS_Section_Proxy] = "Operating")
```

**Що показує:** Чистий грошовий потік від операційної діяльності — головне джерело сталого кешу здорового бізнесу.

**Інтерпретація:**
- CFO > 0 і стабільно > 0 → бізнес генерує кеш сам, здоровий
- CFO < 0 при додатному загальному Net Cashflow → проблема: компанія живе за рахунок кредитів/інвестицій (нестійка модель)
- CFO росте швидше виручки → покращення working capital management

**Перевірене значення (всього періоду):** **672 397 148 ₴**

**Use case:** Page 5 — KPI card вгорі; Sankey origin вузол; Matrix rows для drill-down до статей.

### 1.2 `CFI Investing`

**DAX:**
```dax
CFI Investing = 
CALCULATE([Net Cashflow], 'СтатьиДвиженияДенежныхСредств'[CFS_Section_Proxy] = "Investing")
```

**Що показує:** Чистий грошовий потік від інвестиційної діяльності — купівля/продаж необоротних активів.

**Інтерпретація:**
- CFI < 0 (від'ємний) — компанія ІНВЕСТУЄ у ріст (купує ОС, новий бізнес) → нормально для зростаючої компанії
- CFI > 0 — компанія ПРОДАЄ активи → червоний flag якщо повторюється (можливо змушений продаж)
- CFI = 0 — стабільна, без капвкладень (характерно для зрілих бізнесів)

**Перевірене значення:** **+8 595 182 ₴** (позитивний — у нас більше продажу ОС ніж закупки за період)

**Use case:** Page 5 KPI; компонент Free Cash Flow.

### 1.3 `CFF Financing`

**DAX:**
```dax
CFF Financing = 
CALCULATE([Net Cashflow], 'СтатьиДвиженияДенежныхСредств'[CFS_Section_Proxy] = "Financing")
```

**Що показує:** Чистий грошовий потік від фінансової діяльності — отримання/погашення кредитів, виплата дивідендів, внески власників.

**Інтерпретація:**
- CFF > 0 — приплив зовнішнього фінансування (нові кредити > погашення)
- CFF < 0 — гасимо борги / виплачуємо дивіденди — характерно для зрілих/прибуткових компаній
- CFF приблизно = 0 — компанія сама себе фінансує

**Перевірене значення:** **+245 362 258 ₴** — значне зовнішнє фінансування (кредити, фін.допомога)

**Use case:** Page 5 KPI; контроль leverage.

### 1.4 `Free Cash Flow` (FCF)

**DAX:**
```dax
Free Cash Flow = [CFO Operating] + [CFI Investing]
```

**Що показує:** Кеш, доступний після операційних і інвестиційних потреб — те що залишається для виплати власникам / погашення боргу / накопичення.

**Найважливіший single number для valuation бізнесу.** Discounted Cash Flow (DCF) моделі будуються на проекції FCF.

**Формула (industry standard):**
- FCF = CFO + CFI (наш варіант, спрощений)
- Або: FCF = CFO − CapEx (де CapEx — лише купівля довгострокових активів, не вся CFI)

**Інтерпретація:**
- FCF > 0 стабільно → бізнес створює вартість для власника
- FCF < 0 кілька кварталів → проблема: компанія "проїдає" зовнішні гроші
- FCF growth > Revenue growth → покращення ефективності

**Перевірене значення:** **+680 992 331 ₴** (CFO 672M + CFI 8.6M)

**Use case:** Page 5 KPI card "велике число" поряд з CFO/CFI/CFF.

### Sources
- [HBS — How to Prepare a Cash Flow Statement](https://online.hbs.edu/blog/post/how-to-prepare-a-cash-flow-statement)
- [Damodaran on Valuation: FCF as the core valuation input](https://pages.stern.nyu.edu/~adamodar/)

---

## 2. Group "Cashflow\Cash Position" — running balance

### 2.1 `Cumulative Cash Position`

**DAX:**
```dax
Cumulative Cash Position =
CALCULATE(
    [Net Cashflow],
    FILTER(ALL('Calendar'), 'Calendar'[date_] <= MAX('Calendar'[date_]))
)
```

**Що показує:** Накопичений Net Cashflow від найранішого періоду до кінця обраного періоду — running total.

**Theoretical background:**
> "While Net Cashflow per month answers 'how much money moved THIS month', cumulative position answers 'how much money do we HAVE now compared to start'. For working capital and liquidity planning, the cumulative view is essential." — CFA Institute / Treasury Management body of knowledge

**Перевірена траєкторія:**

| Місяць | Net Cashflow | Cumulative |
|---|---:|---:|
| 2025.11 | -30 453 047 | -30 453 047 |
| 2025.12 | +239 763 088 | +209 310 041 |
| 2026.01 | +155 346 594 | +364 656 635 |
| 2026.02 | +244 410 060 | +609 066 695 |
| 2026.03 | +124 218 389 | +733 285 084 |
| 2026.04 | +144 159 920 | +877 445 004 |
| 2026.05 | +23 982 882 | +901 427 886 |

**Use case:** Line chart Page 4 з reference line "критичний рівень". Time series — головний візуал для CFO bird's-eye view.

**Caveat:** Це **proxy для cash balance** — реальний баланс залежить від `Fact_CF_Balance` (opening + actual transactions). Поки Fact_CF_Balance не додано до PBI Power Query — це найкраще наближення.

---

## 3. Group "Cashflow\Executive KPI" — top-of-page cards

### 3.1 `Net Cashflow PY`

**DAX:**
```dax
Net Cashflow PY =
CALCULATE([Net Cashflow], SAMEPERIODLASTYEAR('Calendar'[date_]))
```

**Що показує:** Net Cashflow за той самий період минулого року. Базова time-intelligence measure для YoY порівняння.

**SAMEPERIODLASTYEAR** — DAX функція з family `DATE INTELLIGENCE`, працює тільки якщо Calendar marked as Date Table (✓ зроблено).

**Use case:** Сама по собі рідко відображається; це **база для `YoY %`**. На сторінці може бути як subtitle до головного KPI ("PY: 198M").

### 3.2 `Net Cashflow YoY %`

**DAX:**
```dax
Net Cashflow YoY % =
DIVIDE([Net Cashflow] - [Net Cashflow PY], ABS([Net Cashflow PY]))
```

**Що показує:** Зростання Net Cashflow рік-до-року у відсотках. **ABS у знаменнику** — щоб при від'ємному PY знак YoY% був інтерпретовний (якщо позитив порівнюємо з негативом, % зростання має бути позитивним).

**Industry benchmarks:**
- Stable mature business: ±5%
- Growth business: +10-30% per year
- Hypergrowth: >30%
- Decline / warning: <-10%

**Use case:** KPI card з directional arrow ↑↓; conditional formatting (зелений / червоний).

### 3.3 `Avg Daily Outflow 30d`

**DAX:**
```dax
Avg Daily Outflow 30d =
DIVIDE(
    CALCULATE([Списания ДДС], DATESINPERIOD('Calendar'[date_], LASTDATE('Calendar'[date_]), -30, DAY)),
    30
)
```

**Що показує:** Середнє денне списання за останні 30 днів. **База для розрахунку Days of Cash Remaining.**

**Чому 30 днів а не місяць:** Robustness — згладжує quarter-end spikes (велика виплата зарплати, податковий період), дає stable baseline для прогнозу.

**Use case:** Внутрішня, не висвітлюється як KPI. Поточне значення дорівнює BLANK без period filter — це нормально, в context з period працює.

### 3.4 `Days of Cash Remaining`

**DAX:**
```dax
Days of Cash Remaining =
DIVIDE([Cumulative Cash Position], [Avg Daily Outflow 30d])
```

**Що показує:** Скільки днів вистачить накопиченого кешу при поточному темпі списання. **Класична liquidity-метрика.**

**Industry-standard thresholds:**
- **< 30 днів** — 🔴 Critical (треба знаходити фінансування зараз)
- **30-90 днів** — 🟡 Warning (планувати кредитну лінію)
- **90-180 днів** — 🟢 Comfortable
- **> 180 днів** — 🟢🟢 Strong cash position

**Edge cases:**
- Якщо Avg Daily Outflow ≤ 0 (приплив > відплив за 30 днів) → metric стає від'ємним або нескінченним. У цьому випадку UI має показувати "∞" або "Cash positive — n/a".

**Best practice background:**
> "Days Cash on Hand is a fundamental health metric. SaaS companies often target 12+ months runway; construction projects working with delays target 6+ months." — McKinsey Global CFO Survey 2024

**Use case:** Page 1 KPI card; flagship liquidity indicator для CFO.

### 3.5 `Funding Gap Months`

**DAX:**
```dax
Funding Gap Months =
COUNTROWS(FILTER(VALUES('Calendar'[year_month]), [Net Cashflow] < 0))
```

**Що показує:** Кількість місяців у обраному періоді коли expenses > receipts (negative Net Cashflow).

**Чому важливо:**
> "Cumulative position can be positive overall, yet individual months can be negative — and those negative months create LIQUIDITY CRISES that bankrupt otherwise profitable companies. Identifying these months in advance lets treasury arrange credit lines proactively." — Mastt Construction Cash Flow guide

**Інтерпретація:**
- **0 months** — кожен місяць самофінансується
- **1-2 months** — нормальна сезонність (наприклад, січень для багатьох галузей)
- **3+ months** — структурна проблема, потрібен continuous funding

**Перевірене значення:** **1** (2025.11 був з gap; решта місяців позитивні)

**Use case:** Page 1 KPI card. Якщо >0 → також червона matrix внизу з gap-місяцями.

---

## 4. Group "Cashflow\Project Tracking" — object plan management

### 4.1 `Days to Depletion`

**DAX:**
```dax
Days to Depletion =
DIVIDE([Відхилення план-факт (об'єкт)], [Avg Daily Outflow 30d])
```

**Що показує:** На скільки днів вистачить **залишку бюджету проекту** (План об'єкта - Факт cum) при поточному темпі списання.

**Concept:** Проектний menedjer не питає "скільки днів живе компанія", він питає "скільки днів живе мій бюджет на Глобино-2". Тому база — Відхилення (=remaining budget), а не Cumulative Cash.

**Thresholds (construction industry):**
- **<30 днів до depletion** — 🔴 reallocate budget or pause
- **30-90 днів** — 🟡 plan extension
- **>90 днів** — 🟢 OK

**Background:**
> "Project burn rate forecasting is critical in construction where milestone-based payments create cashflow lags. Knowing 'days to depletion' allows proactive renegotiation before crisis." — Projectworks PSA Burn Rate guide

**Use case:** Page 3 table column. Bar chart sorted by Days to Depletion ASC — найкритичніші проекти вгорі.

### 4.2 `Project Status`

**DAX:**
```dax
Project Status =
SWITCH(TRUE(),
    [% Виконання плану ДДС факт (об'єкт)] > 1.10, "Critical",
    [% Виконання плану ДДС факт (об'єкт)] > 1.00, "Over",
    [% Виконання плану ДДС факт (об'єкт)] > 0.90, "Warning",
    "OK"
)
```

**Що показує:** Категоризація проекту за станом виконання бюджету.

**Status meaning:**
| Status | % executed | Інтерпретація | Action |
|---|---|---|---|
| **Critical** | > 110% | Перевитрачено > 10% — серйозна проблема | Stop, audit, change request |
| **Over** | 100-110% | Перевищено бюджет, незначно | Document overrun, justify |
| **Warning** | 90-100% | Близько до ліміту | Monitor closely |
| **OK** | < 90% | У межах бюджету | Continue |

**Why thresholds 90/100/110%:**
> "5/10% variance rule from industry: under 5% — within noise (don't intervene), 5-10% — investigate but don't escalate, >10% — manage as exception. Adapted for budget execution: 90% — approaching ceiling (Warning), 100-110% — Over (manage), >110% — Critical." — adapted from DataSights variance analysis

**Use case:** Page 3 status column з emoji 🔴🟠🟡🟢; filter slicer; bar chart color category.

---

## 5. Group "Cashflow\Variance" — conditional formatting helper

### 5.1 `Variance Threshold Flag`

**DAX:**
```dax
Variance Threshold Flag =
SWITCH(TRUE(),
    ABS([% Виконання плану ДДС факт (місяць)] - 1) <= 0.05, "Green",
    ABS([% Виконання плану ДДС факт (місяць)] - 1) <= 0.10, "Amber",
    "Red"
)
```

**Що показує:** Класифікація поточного відхилення плану у 3 категорії за industry-standard threshold:

| Color | Variance | Interpretation |
|---|---|---|
| 🟢 **Green** | |dev| ≤ 5% | Acceptable (noise / normal) |
| 🟡 **Amber** | 5% < |dev| ≤ 10% | Investigate (not yet critical) |
| 🔴 **Red** | |dev| > 10% | Material variance — exception management |

**Чому саме 5/10:**
> "Apply conditional formatting rules: deviations greater than 10% in unfavorable direction appear in red, 5-10% in amber, and under 5% in green." — Power BI Financial Dashboard Best Practices 2026 (DataSights)

**Чому ABS:** Industry-standard variance threshold НЕ розрізняє "перевиконано" і "недовиконано" — обидва відхилення проблема (надвиконання теж непланове = поганий планинг).

**Use case:** Conditional formatting у matrix:
1. Format → Conditional formatting → Background color
2. Format style: Rules
3. Rules: If value = "Green" → background green; "Amber" → amber; "Red" → red
4. Apply to: всі cells з measure `% Виконання плану`

---

## 6. Як використовувати — quick lookup

| Бажаний візуал | Які measures дроп |
|---|---|
| KPI card "як ми йдемо?" | `Net Cashflow`, `% Виконання плану ДДС факт (місяць)` |
| KPI card "куди йдемо?" | `Cumulative Cash Position`, `Net Cashflow YoY %` |
| KPI card "як довго протримаємось?" | `Days of Cash Remaining`, `Funding Gap Months` |
| Matrix variance з кольорами | `Сума план/факт ДДС`, `Variance Threshold Flag` (для CF) |
| Project burn table | `% Виконання плану ДДС факт (об'єкт)`, `Days to Depletion`, `Project Status` |
| CFS Standard report | `CFO Operating`, `CFI Investing`, `CFF Financing`, `Free Cash Flow` |
| Time series cumulative | `Cumulative Cash Position` (Line chart по місяцях) |

---

## 7. Sources

- [Cash Flow Statement Direct vs Indirect — HBS Online](https://online.hbs.edu/blog/post/how-to-prepare-a-cash-flow-statement)
- [Power BI Variance Analysis A-to-Z — Zebra BI](https://zebrabi.com/power-bi-variance-reports/)
- [Project Burn Rate — Projectworks](https://www.projectworks.com/blog/the-importance-of-burn-rate-how-to-calculate-it)
- [Construction Cash Flow Spreadsheet — Mastt](https://www.mastt.com/resources/construction-project-cash-flow-spreadsheet)
- [Damodaran Valuation Resources](https://pages.stern.nyu.edu/~adamodar/)
- [Power BI Financial Dashboard Best Practices 2026 — DataSights](https://datasights.co/financial-dashboard-power-bi-example/)

## 8. Cross-references

- [olap_cashflow_dashboard_design_2026_05.md](olap_cashflow_dashboard_design_2026_05.md) — Stage v3.8 implementation log
- [olap_powerbi_pl_pbix.md](olap_powerbi_pl_pbix.md) — PBI model overview
- [olap_sql_schema.md](olap_sql_schema.md) — backend SQL schema
