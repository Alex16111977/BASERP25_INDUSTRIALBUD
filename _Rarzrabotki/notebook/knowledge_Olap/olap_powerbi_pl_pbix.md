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
| `ОбъектыРасчetов` (2026-05-17) | dbo.Dim_ObjektyRaschetov | `Sql.Database("SQLSERVER","OlapBASERP")` (партиція `Dim_ObjektyRaschetov`) |

> ⚠️ **Нова таблиця в Import-моделі — навігатор SQL кешується:**
> `Источник{[Schema="dbo",Item="Dim_ObjektyRaschetov"]}[Data]` дає
> `[Expression.Error] Ключу не відповідає жоден рядок` поки PBI Desktop не
> оновить каталог джерела вручну. Обхід (нативний `[Query="SELECT … FROM
> dbo.Dim_ObjektyRaschetov"]`) тригерить prompt дозволу нативного запиту →
> headless MCP Refresh зависає. **Рішення: користувач у Power BI Desktop
> робить Refresh таблиці «ОбъектыРасчetов» (підтвердити будь-який
> credential/native-query prompt) → зв'язок 13.2 → Ctrl+S.**

> Назва Dim лишена `Dim_PAP_Articles` (прецедент `Dim_DenezhnyeSredstva`),
> DAX-міри референсують саме її — НЕ перейменовувати без оновлення мір.

### 13.2 Зв'язки (Many→One, single-direction) — 10 (оновлено 2026-05-17)

`Fact_Balance[PAP_Article_ID]→Dim_PAP_Articles`, `[Organization_ID]→Организации`,
`[Department_ID]→СтруктураПредприятия`, `[Counterparty_ID]→Контрагенты`,
`[Individual_ID]→ФизическиеЛица` (підзвітні особи — `Свод_ДенежныеСредства`
гілка ПодотчетноеЛицо), `[Contract_ID]→ДоговорыКонтрагентов`,
`[Period]→Calendar[date_]`.
**+3 додано 2026-05-17 (drill-down Себестоимость/ДенСр):**
- `[Item_ID]→Номенклатура[Item_ID]` (auto-detected після Refresh; активна)
- `[Warehouse_ID]→Склады[Warehouse_ID]` (нова таблиця `Склады`=Dim_Warehouses,
  4 кол.; зв'язок auto-detected)
- `[Cash_ID]→Dim_DenezhnyeSredstva[Cash_Account_ID]` (створено через MCP
  `relationship_operations Create`, дзеркало `Fact_Cashflow.Cash_Account_ID`;
  після `model Refresh Calculate`). DAX-перевірка січ.2026/ТОВ:
  безнал `Account_Type=BankAccount` Close=50 435 887,99; нал Cashbox
  24 345 737,54 + плуги(без субконто, Cash_ID=NULL) 297 122,93 = 24 642 860,47
  == УпрБаланс ✓.

**+1 додано 2026-05-17 (`Свод_РасчетыСПартнерами`, через MCP):**
- `[SettlementObj_ID]→ОбъектыРасчetов[SettlementObj_ID]` (Many→One,
  single-direction, **активна**) — Dim `Dim_ObjektyRaschetov` (13957, плоский
  `_Reference319`; FK 2291/2291=100%, плуги SettlementObj_ID=NULL). Таблиця
  «ОбъектыРасчetов» — партиція **нативний `[Query="SELECT … FROM
  dbo.Dim_ObjektyRaschetov"]`** (навігатор `Источник{[Item=…]}` кешує нові
  таблиці в сесії PBIX → давав `[Expression.Error] Ключу не відповідає
  жоден рядок`; native query це обходить, state=Ready). Зв'язок створено
  `relationship Create` + `model Refresh Calculate`; сховані
  `SettlementObj_ID/ПометкаУдаления/Loaded_At` (IsHidden), видима колонка
  `ОбъектРасчетов`. Drill-down перевірено DAX:
  `SUMMARIZECOLUMNS('ОбъектыРасчetов'[ОбъектРасчетов], …Sum_Close…)` дає
  розбивку по реальних об'єктах; січ.2026/ТОВ клієнти Σ=12 338 631,09,
  постач Σ=−62 806 237,00. **⚠️ in-memory — обов'язковий Ctrl+S.**

**Partner_ID — без прямого зв'язку** (неоднозначний шлях через
Контрагенты→Партнеры; дзеркалить Fact_Cashflow). OperObject/Intangible —
без Dim у моделі (drill-down наступний цикл, не блокує).

**+2026-05-18 (`Свод_РасчетыСПартнерами` → субконто Контрагент/Договор/
Партнёр):** 1С тепер наповнює виміри РС `Контрагент`/`Партнер`/`Договор`
для розрахункових Source з `АналитикаУчетаПоПартнерам`. **Модель PL.pbix
НЕ змінювалась** — зв'язки `Fact_Balance[Counterparty_ID]→Контрагенты`
та `[Contract_ID]→ДоговорыКонтрагентов` вже існують (з List, див. вище
рядки 303–305); `[SettlementObj_ID]→ОбъектыРасчetов` теж. Раніше
`Counterparty_ID/Contract_ID` у Fact_Balance були NULL для розрахунків
(регістр не наповнювався) → drill-down порожній. Після доробки 1С +
перепроведення янв2026 + `--run-once fact_balance --period 2026-01`
(run_id=290) деталі розрахунків мають `Counterparty_ID/Contract_ID/
Partner_ID` 100%, FK→Dim 0 orphans (`verify_olap_balance_raschety_
kontragent.py` PASS), повний баланс незмінний (Σ Sum_Close=0, Актив
288 787 750,11). **Дія користувача:** Power BI Desktop → **Refresh**
(Power Query тягне свіжий Fact_Balance) → **Ctrl+S**. Зміни моделі не
потрібні (нічого не створювалось).

> ⚠️ **РЕ-ETL ПО КОЖНОМУ МІСЯЦЮ ОКРЕМО (lesson 2026-05-19):** доробку
> 2026-05-18 (run_id=290) застосували **лише до янв2026**. Грудень2025 та
> лютий2026 у `Fact_Balance` лишались з пустим `Counterparty_ID` для
> `РасчетыСКлиентамиПоСрокам`/`РасчетыСПоставщикамиПоСрокам` → drill-down
> контрагентів у звіті порожній за ці місяці (1С-регістр
> `А_ОтчетБаланс_Свод` за грудень контрагентів МАВ — 1995/2038 +284/298;
> stale була тільки OLAP-копія). **Корінь:** `python main.py` (без
> прапорців) виконує лише `dim_catalogs/fact_pnl/fact_cashflow` —
> **`fact_balance` НЕ у дефолті** (див. ETL §Default mode), тож звичайний
> прогон + PBI Refresh Balance НЕ оновлює. Лікування — окремо per-period:
> `python main.py --run-once fact_balance --period 2025-12` (run_id=294,
> 7757) + `--period 2026-02` (run_id=295, 8321). Після цього Dec/Jan/Feb
> `cp_filled` збігається з 1С (Dec 284/1995, Feb 404/1970). **Правило:**
> після БУДЬ-ЯКОЇ 1С-доробки субконто `А_ОтчетБаланс_Свод` ре-ETL
> `fact_balance` треба прогнати **за КОЖЕН вже завантажений місяць**, не
> лише за поточний. Діагностика: `Python/test/test_dec2025_reg_
> counterparty.py` (Section 1 — 1С vs OLAP, без date-параметрів).
> **Дія користувача:** Power BI Desktop → **Refresh** (свіжий Fact_Balance
> Dec/Feb) → **Ctrl+S**.

> ⚠️ **Знайдено баг-шум авто-детекту (НЕ виправлено, поза запитом):**
> Power BI авто-створив зв'язки по generic-колонці `Hierarchy_Path`:
> `СтатьиДвиженияДенежныхСредств[Hierarchy_Path]↔Номенклатура` (BothDirections,
> One-One, **активна** — шкідлива: хибний крос-фільтр) + 3 неактивні
> (СтруктураПредприятия/СтатьиРасходов/Партнеры → Номенклатура). Рекомендація:
> видалити/деактивувати активну (узгодити з користувачем — це спільна модель).

### 13.2-bis Колонка `Fact_Balance[ТипНалога]` (2026-05-18, через MCP)

`Свод_ПрочиеАктивыПассивы_Прямой` LIVE → Fact_Balance +колонка **TaxType**
(`Перечисление.ТипыНалогов`, розшифровка ПАП.Аналитика статті «Налоги»).
Через MCP `column_operations Create`: модельна колонка **«ТипНалога»** на
Fact_Balance (1С-нотація, `SourceColumn=TaxType`, `DataType=String`,
`IsHidden=false`, `SummarizeBy=None` — дзеркало `Source`).

**+ Dim `ТипыНалогов` (2026-05-18, на вимогу користувача — образець
Dim_ObjektyRaschetov):** SQL `dbo.Dim_TaxTypes` (15 рядків: 14 enum +
`ПустаяСсылка`; ключ=метаім'я як Fact_Balance.TaxType, `TaxType_Name`=
синонім 1С UI, `EnumOrder`). DDL `scripts/ddl_dim_tax_types.sql` + сидер
`scripts/seed_dim_tax_types.py` (імена/синоніми з 1С метаданих — `_Enum1651`
у SQL не має імен). PL.pbix: таблиця **«ТипыНалогов»** — партиція
**нативний `[Query="SELECT TaxType, TaxType_Name, EnumOrder FROM
dbo.Dim_TaxTypes"]`** (навігатор кешує нові таблиці §13.1, native query це
обходить; state=NoData до Desktop Refresh). Колонки: `TaxType`(прих., ключ),
**«ТипНалога»** (видима, `SourceColumn=TaxType_Name`, SortBy=`EnumOrder`),
`EnumOrder`(прих.). Зв'язок **`Fact_Balance[ТипНалога]→ТипыНалогов[TaxType]`**
(Many→One, OneDirection, active) — створено MCP `relationship Create`
(підтверджено `relationship List`: 39 зв'язків). FK 100% (verify
`Fact_Balance.TaxType` без Dim = 0).
⚠️ **Навігатор SQL кешує схему (§13.1):** партиція Fact_Balance —
`Источник{[Item="Fact_Balance"]}[Data]`; MCP `partition Refresh` падає
`Столбец "TaxType" не существует в наборе строк` (кеш не бачить нову
SQL-колонку). Таблиця «ТипыНалогов» — native query, partition state=NoData
→ тригерить prompt дозволу нативного запиту. **Резолюція (як усі
Balance-зміни §13.7 + Dim_ObjektyRaschetov §13.2):** користувач у Power BI
Desktop робить **Refresh Fact_Balance + Refresh «ТипыНалогов»**
(інтерактивний — навігатор перечитує каталог, підтягує TaxType; підтвердити
будь-який credential/native-query prompt) → **Ctrl+S**. Метадані колонки +
таблиці + зв'язку вже в моделі (in-memory); дані заповняться після Desktop
Refresh. Очікувані
значення (січ.2026/ТОВ): Налоги по `Fact_Balance[ТипНалога]` НДС
9 246 711,36 / ДругиеНалоги 72 252,00 / НДФЛ 4 925,02 / ВоенныйСбор
1 368,07 / НачисленныйЕСВ 6 019,47 (== Карточка), решта статей TaxType=
`"ПустаяСсылка"`. Verify (SQL, до Desktop Refresh):
`scripts/verify_olap_balance_papdirect.py` PASS.

### 13.2-ter Dim `ТипПоказателя` + зв'язок (2026-05-18, через MCP)

Фінансист додав у `А_ОтчетБаланс_Свод` вимір **`ТипПоказателя`**
(`Перечисление.ВидыСтатейУправленческогоБаланса`: Актив/Пассив/
АктивПассив), заповнюється централізовано в `ПровестиБалансСвод`
формулою штатного `Отчет.УправленческийБаланс`
(`ВЫБОР КОГДА АктивПассив=ЗНАЧЕНИЕ(...АктивПассив) ТОГДА Пассив ИНАЧЕ
АктивПассив КОНЕЦ`; «Налоги»→**Пассив**). Мета: у Power BI поділ
Актив/Пассив як у штатному звіті (а не raw «АктивПассив»).

ETL: `Fact_Balance.TipPokazatelya` (`_Fld56131RRef`,
enum_resolver→кирилиця «Актив»/«Пассив», конвенція як Source/TaxType).
**+ Dim `Dim_TipPokazatelya`** (образець Dim_TaxTypes): SQL 4 рядки
(Актив/Пассив/АктивПассив + `ПустаяСсылка`; ключ=метаім'я,
`TipPokazatelya_Name`=синонім 1С UI, `EnumOrder`). DDL
`scripts/ddl_dim_tip_pokazatelya.sql` (+ `ALTER Fact_Balance ADD
TipPokazatelya`) + сидер `scripts/seed_dim_tip_pokazatelya.py`.

PL.pbix: користувач додав таблицю **«ТипПоказателя»** (4 кол.:
`TipPokazatelya` ключ, `TipPokazatelya_Name`, `EnumOrder`, `Loaded_At`).
Зв'язок через MCP **`Fact_Balance[TipPokazatelya]→ТипПоказателя
[TipPokazatelya]`** (Many→One, OneDirection, **active**).
⚠️ **УРОК — ложна авто-зв'язка по `EnumOrder`:** Power BI auto-detect
створив зайву зв'язку `ТипПоказателя[EnumOrder]→ТипыНалогов[EnumOrder]`
(active, BothDirections — хибний матч службового `EnumOrder`=0,1,2 двох
різних Dim; спотворював би `ТипыНалогов`). Через MCP `relationship
Delete` видалено; правильну `Fact_Balance[TipPokazatelya]→
ТипПоказателя` (auto-detected, була `isActive=false`) `relationship
Activate`. **Загальне правило:** після додавання Dim з колонкою
`EnumOrder`/`*_ID` перевіряти `relationship List` на хибні авто-зв'язки
між службовими стовпцями різних Dim і видаляти їх. `verify_olap_balance_
tippokazatelya.py` PASS (TipPokazatelya 0 NULL, FK→Dim 100%, «Налоги»→
Пассив, ПОЛНЫЙ БАЛАНС дек 278 093 267,32 / янв 288 787 750,11 ==
штатний звіт). Метадані зв'язку in-memory → **Ctrl+S**.

**+ Стандартизація під 1С-нотацію (2026-05-19, через MCP, образець
§13.2-bis ТипыНалогов / §13.8 еталон):** колонку `TipPokazatelya_Name`
перейменовано на видиму **«ТипПоказателя»** (`column_operations Rename`;
sourceColumn лишився `TipPokazatelya_Name`), `sortByColumn=EnumOrder`;
службові `TipPokazatelya` (ключ FK), `EnumOrder`, `Loaded_At` →
`isHidden=true` (`column_operations Update`). Зв'язок
`Fact_Balance[TipPokazatelya]→ТипПоказателя[TipPokazatelya]` цілий
(приховання ключа НЕ ламає зв'язок — прецедент §13.8 Cash_Account_ID;
`relationship List`=40, усі active, хибних авто-зв'язків між службовими
стовпцями нема). DAX-перевірка: таблиця `ТипПоказателя` 4 рядки, сорт по
`EnumOrder` (Актив=0 / Пассив=1 / Актив/Пассив=2 / (Не визначено)=99).
**In-memory → Ctrl+S.** Зріз по видимій колонці (Актив/Пассив);
службові ключ/порядок/Loaded_At сховані як у `ТипыНалогов`/
`СтруктураПредприятия`. **Уточнення 2026-05-19 (користувач у PBI
Desktop, збережено 11:35):** видиму колонку скорочено
**«ТипПоказателя» → «Тип»** → актуальне посилання `'ТипПоказателя'[Тип]`
(`column_operations List` підтвердив: `TipPokazatelya`/`EnumOrder`/
`Loaded_At` лишились прихованими, sortByColumn=EnumOrder, зв'язок
`Fact_Balance[TipPokazatelya]→ТипПоказателя` цілий).

### 13.3 DAX-міри (Table_Measures, displayFolder `Balance`) — 10 (+1 у `Balance\Фільтри`)

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
[Баланс Вих (абс)] = ABS([Баланс Вих])   // 2026-05-19; formatString #,0.00;-#,0.00;
[Видимість рядка Баланс] =               // 2026-05-19; displayFolder Balance\Фільтри; formatString 0
    VAR Close_ = SUM(Fact_Balance[Sum_Close])
    RETURN IF(Close_ <> 0, 1, BLANK())
```

Двосторонні `AktivPassiv` діляться по знаку Sum_Close (як Налоги, канон OD-9).

`[Баланс Вих (абс)]` (2026-05-19, через MCP `measure_operations Create`,
образець `Сума казна ДДС (абс)` / `Сума план ДДС (об'єкт) (абс)`) — модуль
`[Баланс Вих]`, щоб у матриці Актив/Пасив пасив (від'ємне сальдо)
відображався з додатним знаком. displayFolder `Balance`, формат як у
ДДС-(абс). DAX-перевірка (по `ТипПоказателя[Тип]`): Актив 845 205 614,22
(без змін) / Пассив −845 175 060,96 → **+845 175 060,96**. In-memory →
Ctrl+S.

`[Видимість рядка Баланс]` (2026-05-19, через MCP `measure_operations
Create`, образець `[Видимість рядка ДДС план-казна (місяць)]` /
`[Видимість рядка PL]`) — helper для приховування порожніх рядків у
матрицях Балансу: `1` якщо `Sum_Close (== [Баланс Вих]) <> 0`, інакше
`BLANK`. displayFolder **`Balance\Фільтри`**, formatString `0`.
**Застосування:** Filters pane матриці → `Filters on this visual` (НЕ `on
this page`) → поле `Видимість рядка Баланс` → `is not blank`. Перевірено
DAX (по `ТипПоказателя[Тип]`): Актив/Пассив (Sum_Close≠0) → `1`; рядки з
нульовим сальдо → `BLANK` (приховуються). Від'ємне сальдо (Пассив) дає
`1` — ховаються лише справжні нулі. In-memory → Ctrl+S.

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

### 13.8 Стандартизація нових Dim під 1С-нотацію (2026-05-17)

Нові/незачеплені довідники приведені до конвенції старих (еталон
`СтруктураПредприятия`: бізнес-ім'я RU видиме, **усе технічне `isHidden`**,
`sourceColumn` зберігається при rename). Через MCP `column_operations`
Rename+Update(IsHidden):

| Таблиця | Перейменовано (видиме RU) | Сховано (isHidden) |
|---|---|---|
| `Склады` | `Warehouse_Name→Склад` | `Warehouse_ID`, `ПометкаУдаления`(Marked_For_Deletion), `Loaded_At` |
| `Номенклатура` | `Item_Name→Номенклатура`, `Is_Group→ЭтоГруппа`, `Marked_For_Deletion→ПометкаУдаления` | `Item_ID`,`Item_Code`,`Parent_ID`,`ЭтоГруппа`,`ПометкаУдаления`,`Loaded_At`,`Hierarchy_Path`,`Hierarchy_Depth`,`Level1..5` |
| `Виды номенклатуры` | `ItemGroup_Name→Вид номенклатуры`, `Is_Group→ЭтоГруппа`, `Marked_For_Deletion→ПометкаУдаления` | `ItemGroup_ID`,`ItemGroup_Code`,`Parent_ID`,`ЭтоГруппа`,`ПометкаУдаления`,`Loaded_At` |
| `Dim_PAP_Articles` * | `PAP_Article_Name→СтатьяАктивовПассивов` | `PAP_Article_ID`,`PAP_Article_Code`,`Parent_ID`,`Is_Group`,`Marked_For_Deletion`,`Loaded_At` |
| `Dim_DenezhnyeSredstva` * | `Account_Name→Денежные средства` | `Cash_Account_ID`,`Organization_ID`,`Currency_Code`,`Loaded_At` |

\* Назви ТАБЛИЦЬ `Dim_PAP_Articles`/`Dim_DenezhnyeSredstva` **НЕ змінено**
(73 DAX-міри + зв'язки залежать; §13.1). Колонки-залежності мір/слайсерів
**не чіпані**: `Dim_PAP_Articles[AktivPassiv]`, `Dim_DenezhnyeSredstva[Account_Type]`
лишені видимими/без rename. Ключі `*_ID` лише сховані (не перейменовані) —
зв'язки цілі.

**Verified (DAX після змін):** Себест Close=83 627 719,44; `[Актив]`/
`[Контроль Актив-Пассив]` рахуються (міри по `[AktivPassiv]` не зламані);
`'Склады'[Склад]` drill-down Себест працює (Цех екстракції 10 117 180,38…);
`Dim_DenezhnyeSredstva[Account_Type]`=BankAccount 50 435 887,99 (зв'язок
`Cash_ID` цілий після приховання `Cash_Account_ID`). 73 міри + усі зв'язки OK.

> ⚠️ Зміни **in-memory** — користувач має зберегти **Ctrl+S** у Power BI
> Desktop (інакше rename/hide втратяться при закритті).

### 13.9 Ієрархії Dim (2026-05-17, образець `ИерархияСтатейДДС`)

Конвенція: `hierarchy Иерархия<Сущность>` з `hideMembers: hideBlankMembers`,
5 рівнів Level1..Level5 (приховані колонки). Створено через MCP
`user_hierarchy_operations`/`column_operations`. Перейменування таблиць
`Dim_PAP_Articles→СтатьиАктивовПассивов`, `Dim_DenezhnyeSredstva→ДенежныеСредства`
зроблено користувачем у PBI Desktop (DAX-fixup автоматичний — 73 міри цілі,
перевірено: Себест=83 627 719,44, [Актив]/[Контроль] рахуються).

| Dim (1С-ієрархічний) | Иерархия | Джерело Level1..5 | Стан |
|---|---|---|---|
| `Номенклатура` | `ИерархияНоменклатуры` | ETL Level1..5 (рекурс. CTE dim_items) — готові | ✅ DONE |
| `СтатьиАктивовПассивов` | `ИерархияСтатейАктивовПассивов` | calc PATH (Parent_ID) Level1..5 | ✅ DONE |
| `Виды номенклатуры` | `ИерархияВидовНоменклатуры` | calc PATH (Parent_ID) Level1..5 | ✅ DONE |
| `ФизическиеЛица` | `ИерархияФизическихЛиц` | calc PATH (Parent_ID) Level1..5 | ✅ DONE |
| `Склады` | `ИерархияСкладов` | **ІЄРАРХІЧНИЙ** (Иерархия груп і елементів, 2 рівні), **без Кода** (Длина кода=0). Recursive-CTE по `_Reference502` БЕЗ `_Code` (як dim_organizations). Dim_Warehouses 347 рядків, 303 групи, Level1..5 | ✅ DONE |
| `ДенежныеСредства` | — | складений з 8 довідників, не єдиний ієрарх. довідник | ✅ SKIP |

**Parent-child Level (PATH):** для довідників без ETL-Level (Parent_ID є):
calc-колонки `_ParentClean` (зануляє корінь/orphan через LOOKUPVALUE),
`_HPath = PATH(ID,_ParentClean)`, `Level{n} = LOOKUPVALUE(Name,ID,
PATHITEM(_HPath,n))` — усі `isHidden`; ієрархія на Level1..5. Візуально
== `ИерархияСтатейДДС`. Перевірено: Level1 заповнені (PAP «Денежные
средства/Дебиторская задолженность…», ФизЛица — ПІБ, Виды — Товары/Продукція).

**Склады — ІЄРАРХІЧНИЙ, recursive-CTE БЕЗ Кода.** Перша спроба впала
`[42S22] Недопустимое имя столбца "_Code"` — бо Длина кода=0 (нема `_Code`),
а НЕ через відсутність ієрархії. `_Reference502` має `_ParentIDRRef`/`_Folder`
(перевірено `scripts/probe_ref502_columns.py`: є _IDRRef/_Description/_Marked/
_ParentIDRRef/_Folder, нема _Code). `ddl/08_dim_warehouses.sql` +
`pipelines/dim_catalogs.json` `dim_warehouses` **виправлено**: recursive-CTE
по `_Reference502` БЕЗ `_Code`/Код (Warehouse_Code теж прибрано), решта як
dim_items. **Урок:** ієрархічність 1С — по Конфігуратору (Иерархический✓) /
фізичних `_ParentIDRRef`/`_Folder`; відсутність `_Code` ≠ неієрархічний
(catalog може мати Длина кода=0). Не плутати.

✅ **ВИКОНАНО (2026-05-17, агентом):** `apply_08_dim_warehouses.py` (PASS)
→ `main.py --run-once dim_catalogs` (run_id=261, Success, 61 685; Dim_Warehouses
347 рядків / 303 групи / Level1 заповнений) → у моделі через MCP додано 9
data-колонок (sourceColumn: Parent_ID/Is_Group/Hierarchy_Path/Hierarchy_Depth/
Level1..5, усі `isHidden`) → Refresh `Склады` → `ИерархияСкладов` на Level1..5
(HideBlankMembers). Перевірено DAX: Себест Sum_Close дрилиться по
`Склады[Level1]` (Глобине 27,3М; Цех металоконструкцій 28,3М; ПІДГІРЦІ 10,9М…),
зв'язок `Fact_Balance[Warehouse_ID]→Склады` цілий.

> ⚠️ **Усі 5 ієрархій + Dim-зміни — in-memory.** Зберегти **Ctrl+S** у Power
> BI Desktop. (Дані Dim_Warehouses в OlapBASERP уже оновлені ETL — Refresh
> моделі підтягне; при наступному відкритті після Ctrl+S усе на місці.)

### 13.10 Розширення ДоговорыКонтрагентов + ФінАгенти + ТипиДоговорів (2026-05-19)

#### Виконано через MCP (in-memory)

**Таблиця `ДоговорыКонтрагентов`** — 8 нових колонок (sourceColumn→RU-назва):

| sourceColumn | RU-назва в моделі |
|---|---|
| `Is_FinAgent_Contract` | Це договір фін.агента |
| `TipDogovora` | Тип договора |
| `Department_Name` | Підрозділ (бух.) |
| `Dept_OkazUslug_Name` | Підрозділ (послуги між підр.) |
| `Partner_Name` | Партнер |
| `Counterparty_Name` | Контрагент |
| `DDS_Article_Forced_Name` | Стаття ДДС (осн. примусово) |
| `Org_Buh_Name` | Організація (бух.) |

Технічна колонка `FinAgent_ID` — прихована (FK для snowflake-зв'язку).

**Нова таблиця `ФінАгенти`** (`Dim_FinAgents`): `FinAgent_ID` прихований,
`FinAgent_Name`→«ФінАгент»; решта тех. колонки приховані.

**Snowflake-зв'язок** `ДоговорыКонтрагентов[FinAgent_ID]→ФінАгенти[FinAgent_ID]`
(Many→One, active) — створено через MCP.

#### User-gated (виконати у Power BI Desktop)

Наступні дії **ще не виконані** (потребують Power BI Desktop):

1. **Refresh `ОбъектыРасчетов`** — 8 нових колонок (`TipRaschetov`, `TipObjektaRaschetov`,
   `Partner_Name`, `Department_Name`, `Counterparty_ID`, `Contract_ID`, `Object_ID`,
   `Object_Type_Name`) ще не в моделі; після Refresh перейменувати/сховати по аналогії.
2. **Додати таблицю `ТипиДоговорів`** — native query:
   ```sql
   SELECT TipDogovora, TipDogovora_Name, EnumOrder FROM dbo.Dim_TipyDogovorov
   ```
3. **Створити зв'язок** `ДоговорыКонтрагентов[TipDogovora]→ТипиДоговорів[TipDogovora]`
   (Many→One, active).
4. **Refresh + Ctrl+S** — MCP-зміни in-memory; нова snowflake-зв'язок потребує пересчёту
   моделі в Desktop. БЕЗ Ctrl+S усі зміни (MCP + нові зв'язки) втрачаються.

> ⚠️ Зміни MCP **in-memory** — обов'язково **Ctrl+S у Power BI Desktop** після
> всіх user-gated дій. Без збереження зміни сесії будуть втрачені.
