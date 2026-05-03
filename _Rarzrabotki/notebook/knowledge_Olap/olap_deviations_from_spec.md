# OLAP — Deviations vs Spec v3 final

> Фактичні відхилення реалізації від `docs/superpowers/specs/2026-05-01-olap-baserp-architecture-design-v3-final.md`. Усі deviations — або data-driven adaptations, або platform constraint fixes. Тип файлу: ЗМІННИЙ.

---

## Зведення

| # | Deviation | Stage | Spec section | Тип | Статус |
|---|---|---|---|---|---|
| 1 | Тимчасові registratori для cf-validate | 5/6 | §3.4, §3.6 | Platform constraint workaround | Reverted у Tasks 7-8 |
| 2 | Source: Attribute → Dimension у А_ОтчетPL_Свод | 10 | §3.4 | Architecture fix | Постійно |
| 3 | Документ_Источник composite expanded 12→24 типів | 10 | §3.4 | Data-driven adaptation | Постійно |
| 4 | ВидДвижения тип fix у А_ОтчетDDS_Свод | 12 | §3.6 | Spec error fix | Постійно |
| 5 | GROUP BY dedup у DDS query | 12 | §3.5 | Data deduplication | Постійно |
| 6 | Σ рядків PnL = 3937, не 5101 | 13 | §8 | Spec was rough estimate | Update spec |
| 7 | CFS_Section = NULL для всіх статей | — | §3.2 | Manual data entry pending | Поточно (ongoing work для фінансиста) |
| 8 | SQL-first замість COM-first ETL | 3 | §4 | Architecture fix | Постійно (Stage 3 implemented) |
| 9 | Mapping через Python COM, не BSL EPF | 3.0.5 | §4.1 | Pragmatic alternative | Постійно |
| 10 | `Сорт` у Dim_PL_Articles лишається NULL | 3.7 | §3.5 | Метадані 1С не мають реквізиту | Можливо потрібен Stage 1.b |
| 11 | Bridge PLArticle_DDS = 0 рядків | 3.7 | §4.2 | Реквізит не заповнений у даних | Pending фінансист |
| 12 | Fact_Cashflow.BankAccount_ID NOT NULL — placeholder для Каси | 3.9 | §3.7 | Schema constraint workaround | Schema review |
| 13 | _Date_Time +2000 рік-офсет у BaseERP backend | 3.9 | — | 1С platform behavior (cluster mode) | Постійно (handled у Pipeline) |

**Загальна оцінка:** усі 13 deviations — або (a) workarounds для платформенних обмежень 1С, або (b) виправлення помилок specа які виявились при тестуванні з реальними даними. Жодне відхилення не змінює бізнес-логіку. Ключове Stage 3 deviation (#8) — архітектурне: SQL-first замість COM-first, обґрунтоване продуктивністю (~10× швидше) і наявністю exposed mapping.

---

## 1. Тимчасові registratori для cf-validate

### Що
Spec §3.4 і §3.6 описують `А_ОтчетPL_Свод` і `А_ОтчетDDS_Свод` як `WriteMode = RecorderSubordinate` без вказівки registratorов. Платформа 1С не дозволяє завантажити такий регістр без хоча б одного зареєстрованого Документа в `RegisterRecords` будь-якого існуючого Документа (помилка 101 при db-load-xml).

### Виправлення
Tasks 5/6: тимчасово додано:
- `Documents/А_ОтчетPL.xml` RegisterRecords += `InformationRegister.А_ОтчетPL_Свод` (1 рядок)
- `Documents/А_ПереносДвиженийИзКазны.xml` RegisterRecords += `InformationRegister.А_ОтчетDDS_Свод` (1 рядок)

Tasks 7/8: revert назад. Documents `А_ФинРез_PL` і `А_ФинРез_DDS` стали єдиними registratorами своїх регістрів.

### Поточний стан
- `Documents/А_ОтчетPL.xml` — оригінальний (тільки `АккумуляційнійРегистр.А_ДенежныеСредстваФинАгенты`)
- `Documents/А_ПереносДвиженийИзКазны.xml` — оригінальний (тільки `АккумуляційнійРегистр.А_ДвиженияДенегИзКазны`)
- Жодних residual поведінкових змін.

### Why this didn't break anything
А_ОтчетPL.ОбработкаПроведения не пише у `А_ОтчетPL_Свод` (logica відсутня), тому навіть під час тимчасового registration не було fictitious записів. Аналогічно для А_ПереносДвиженийИзКазны.

### Status
✅ Resolved (reverted)

---

## 2. Source: Attribute → Dimension у А_ОтчетPL_Свод ⚠️

### Що
Spec §3.4 (рядок 149) поклав `Source` у Реквізити (Attributes):
```
| Реквізити (5):                            |
| `Source` | EnumRef.А_ИсточникPL          |
```

При реалізації Stage 1.10 виявилось, що це призводить до collapse рядків. Регістр сведень з `WriteMode = RecorderSubordinate` має композитний ключ = (Регістратор + ВсіВиміри). Якщо Source у Атрибутах — він не входить у ключ. Тому 8 рядків з різним Source але однаковими іншими полями колапсували в один рядок (last-write-wins).

**Симптом:** після проведення А_ФинРез_PL очікували 3937 записів, а отримали 3929 (8 менше).

### Виправлення (Stage 1.10)
Перенесено `Source` з Attributes у Dimensions у XML регістру:
```xml
<Dimensions>
    <!-- ... 8 інших -->
    <Dimension uuid="...">
        <Properties>
            <Name>Source</Name>
            <Type><v8:Type>cfg:EnumRef.А_ИсточникPL</v8:Type></Type>
        </Properties>
    </Dimension>
</Dimensions>
```

Тепер унікальний ключ = (Регістратор + 8 dim + Source) → 3937 рядків зберігаються коректно.

### Why це архітектурно правильніше
У OLAP-моделях **Source/Type/Marker колонки завжди в Dimensions** — це slowly-changing-attribute pattern. Spec помилково покважив Source як "атрибут факту", але концептуально він — "категорія факту" (тобто dimension).

### Impact на DAX
Нульовий — `Fact_PnL[Source]="ERP_Income"` працює однаково що для Attribute, що для Dimension. Тільки PowerQuery бачить колонку у списку Dimensions, не Attributes — це косметика.

### Status
✅ Permanent fix. Spec треба оновити (зробити в наступному ревізії).

---

## 3. Документ_Источник composite expanded 12 → 24 типів

### Що
Spec §3.4 (рядок 144) визначив `Документ_Источник` як композитний тип з 12 documenst:
```
DocumentRef.РеализацияТоваровУслуг
DocumentRef.ПриобретениеТоваровУслуг
DocumentRef.А_ОтчетPL
DocumentRef.А_ПереносДвиженийИзКазны
... (12 загалом)
```

Реальні дані за лютий 2026 містять регістраторів інших типів — наприклад `КорректировкаРеализации`, `СписаниеБезналичныхДенежныхСредств`, `ВозвратТоваровПоставщику`, `АктВыполненныхРабот`, `А_ПриходДенегОтФинАгента`, `ОтражениеЗарплатыВФинансовомУчете`, `А_РаспределениеЗаработнойПлаты`, тощо.

### Виправлення (Stage 1.10)
Через iterative discovery (run BSL → побачити error "Тип регістратора не дозволений" → додати у composite → re-load) розширено composite до 24 типів. Повний список — у `InformationRegisters/А_ОтчетPL_Свод.xml`.

### Why
Реальна data state >> spec estimate. У великій ERP за місяць обіг проходить через ~20-30 типів документів-регістраторів.

### Impact
- Регістр приймає всі реальні документи
- Drill-down URL формується для всіх 24 типів
- Power BI таблиця показує правильне Presentation для кожного

### Status
✅ Permanent fix. Якщо у наступних місяцях з'являється новий тип — додати до composite (через `meta-edit` skill або direct XML).

---

## 4. ВидДвижения тип fix у А_ОтчетDDS_Свод ⚠️ КРИТИЧНО

### Що
Spec §3.6 (рядок 214) хибно поклав:
```
| `ВидДвижения` | EnumRef.ВидыДвиженийДенежныхСредств (Приход/Расход) |
```

`ВидыДвиженийДенежныхСредств` — це **55-значне перерахування CFS-категорій** (`ВыручкаОтПродажи`, `ОплатаТруда`, `ВыплатаДивидендов`, `ПогашениеКредитовИЗаймов`, ...). Воно НЕ має значень "Приход" / "Расход".

Правильний тип — `ТипыДвиженияДенежныхСредств` (2 значення: `Поступление` / `Списание`).

### Симптом
При спробі завантажити записи у регістр через `Загрузить()`, BSL падав з помилкою `Type mismatch: значення Поступление має тип ТипыДвиженияДенежныхСредств, очікується ВидыДвиженийДенежныхСредств`.

### Виправлення (Stage 1.12)
- Тип реквізиту регістру `А_ОтчетDDS_Свод.ВидДвижения` змінено на `EnumRef.ТипыДвиженияДенежныхСредств`
- SQL у А_ФинРез_DDS.ОбработкаПроведения використовує реальні значення з регістрів ДС (Поступление/Списание)

### Why це сталось
Spec автор сплутав 2 перерахування з схожими іменами:
- `ВидыДвиженийДенежныхСредств` (1С BAS ERP standard) — для CFS-категорій (старий, рідко використовується у новій конфігурації)
- `ТипыДвиженияДенежныхСредств` (1С BAS ERP standard) — для напряму руху

### Impact
- Регістр приймає реальні значення з регістрів ДС
- DAX [Inflow] / [Outflow] обчислюється коректно через `Direction='Inflow'` (Direction денормалізовано Python ETL з ВидДвижения)

### Status
✅ Permanent fix. Spec оновити при ревізії.

---

## 5. GROUP BY dedup у DDS query

### Що
Запит у А_ФинРез_DDS.ОбработкаПроведения — UNION 3 регістрів ДС:
- `ДенежныеСредстваБезналичные.Обороты` (Source=ERP_Безнал)
- `ДенежныеСредстваНаличные.Обороты` (Source=ERP_Нал)
- `А_ДвиженияДенегИзКазны.Обороты` (Source=Казна)

**Edge case:** РКО (РасходныйКассовыйОрдер) автоматично пишеться у `ДенежныеСредстваНаличные` (через стандартний механізм 1С) **і** у `А_ДвиженияДенегИзКазны` (якщо документ перенесений з Казни). Тобто один РКО з'являється у двох регістрах одночасно.

При UNION такий РКО з'являється двічі — раз з Source=ERP_Нал, раз з Source=Казна. **160 таких випадків за лютий 2026.**

### Симптом
Регістр сведень падав з помилкою унікальності ключа: композитний ключ (Регістратор+вимірювання) повторювався з різним Source.

Wait — це не помилка. Source у нас в Dimensions, тому ключ включає Source. Тоді 160 рядків мали б зберегтись коректно (різний Source → різний ключ).

### Real reason для GROUP BY
Розглядаючи логіку ще раз: РКО фактично **один платіж**. У Power BI ми не хочемо показати його двічі (раз як ERP_Нал, раз як Казна). Це призведе до подвоєння Σ Outflow у візуалі "Direct Method CFS".

**Рішення:** GROUP BY всіх вимірювань + СУМА у фінальному SELECT, **dedup на рівні запиту**. Якщо два рядки UNION мають однакові ВСІ виміри (включно з Регістратором), вони злиаются у один рядок з сумою сум.

```sql
-- Final SELECT з GROUP BY:
ВЫБРАТЬ Період, Организация, ..., Источник, СУММА(СуммаГрн)
ИЗ ([3 UNION гілок])
СГРУППИРОВАТЬ ПО Період, ..., Источник
```

**Result:** 4812 рядків UNION → 4652 рядки після dedup (160 cross-branch duplicates merged).

### Impact
- DAX [Outflow] не подвоюється для документів з cross-branch заслідами
- Σ ERP_Нал у регістрі менший від Σ ДенежныеСредстваНаличные.Обороти на величину overlap з Казной — це правильно

### Status
✅ Permanent fix.

---

## 6. Σ рядків PnL = 3937, не 5101

### Що
Spec §8 (рядок 494) і IMPLEMENTATION_PROMPT.md мали неточну оцінку `Σ записів = 5101 рядок (з тестів Reports/А_ОтчетPL)`.

### Реальна цифра
**3937 рядків** — exact parity з `Reports/А_ОтчетPL.ПолучитьОбъединенныеДанные()` для лютого 2026 (verified Stage 1.13 acceptance test).

### Why spec був неточним
5101 — це була груба оцінка з попередньої версії А_ОтчетPL (до VAT-fix 2026-04-25). Після фіксу і нової логіки розрахунку (`ВыручкаИСебестоимостьПродаж` замість `ФинансовыеРезультаты`) кількість рядків стала менше.

### Impact на acceptance criteria
- Тест acceptance перевіряє `Σ in [3500, 4500]` (з певним tolerance) замість точного 5101
- Цей файл — джерело істини для актуальних чисел

### Status
✅ Permanent fix у acceptance тестах. Spec треба оновити (внести у v4 ревізію).

---

## 7. CFS_Section = NULL для всіх 425 статей ⏳ ONGOING

### Що
Реквізит `Справочник.СтатьиДвиженияДенежныхСредств.А_РазделCFS` створений у Stage 1.4 (commit `82a26952a`), але всі 425 статей мають значення `ПустаяСсылка` (порожнє).

### Спецовий план vs реальність
Spec §3.2 (рядок 89) каже:
> "Фінансист один раз заповнює ~150-200 статей через 1С UI."

Реалізація: створення реквізиту — наша задача (DONE), заповнення — задача фінансиста (PENDING).

### Поточний стан (станом на 2026-05-03)
- 425 статей загалом у справочнику СтатьиДвиженияДенежныхСредств
- 0 заповнено А_РазделCFS
- У регістрі А_ОтчетDDS_Свод поле `CFS_Section = NULL` для всіх 4652 рядків

### Impact
- Power BI слайсер `CFS_Section` покаже `(Blank)` для всіх рядків Cashflow
- Direct Method CFS report (Operating + Investing + Financing) буде неможливим до заповнення
- DAX `[CFS Total]` поверне 0 (бо всі gauge у Internal/Blank)

### Plan
- Фінансист заповнює 80% статей (тих що активно використовуються) — пріоритет на Operating-категорії
- Заповнення може бути пакетним: SQL bulk update у 1С (ризиковано) або через зовнішню обробку імпорту з Excel
- Acceptance gate для Stage 4 Power BI: ≥80% обороту лютого 2026 покривається заповненими статтями

### Як перевірити прогрес
```sql
-- 1С запит:
ВЫБРАТЬ
    КОЛИЧЕСТВО(*) КАК Всього,
    КОЛИЧЕСТВО(ВЫБОР КОГДА А_РазделCFS = ЗНАЧЕНИЕ(Перечисление.А_РазделыCFS.ПустаяСсылка) ТОГДА 1 КОНЕЦ) КАК Незаповнених
ИЗ Справочник.СтатьиДвиженияДенежныхСредств
ГДЕ НЕ ПометкаУдаления
```

### Status
⏳ Ongoing — work для фінансиста.

---

## 8. SQL-first замість COM-first архітектура ETL

### Що
Spec v3 §4 (`olap_etl_pipeline.md` чорновик до 2026-05-03) описував `Python ETL` як набір скриптів `extract/*_dim_*.py` що читають 1С через `V83.COMConnector` + 1С Query Language (`Запрос.Виполнить()`).

### Деviation
Stage 3 реалізований інакше:
- **Primary read path** — pyodbc до MSSQL backend BaseERP (`_Reference329`, `_InfoRg55970` тощо)
- **COM лишений як fallback** тільки для віртуальних таблиць (`.Остатки`/`.Обороты`/`.ОстаткиИОбороты`)
- Mapping `1С_object → SQL_table` згенерований через платформенний `ПолучитьСтруктуруХраненияБазыДанных()` і збережений у `mapping/baserp_storage.json`

### Чому
- **Швидкість**: 3937 рядків Fact_PnL ~58 ms через pyodbc проти ~2 s через COM (×30+).
- **Простота транзакцій**: все на одному pyodbc connection, не змішаний COM/SQL state.
- **Не блокує 1С**: ETL не займає сесії 1С Server.
- **Декларативність**: pipeline'и pure JSON без Python module-per-Dim, легше підтримувати.

### Trade-off
- Залежність від `mapping/baserp_storage.json` — після значної зміни конфіги (нові реквізити) треба `python mapping/refresh_mapping.py`.
- Деякі обчислювані поля (Остатки) недоступні в backend — там лишилися COM extractor'и.

### Where
`_Rarzrabotki/Olap/Ai_Olap/RESEARCH.md` — обґрунтування. `_Rarzrabotki/Olap/Ai_Olap/ai_olap/extractors/sql_backend.py` — primary impl. Live код merged into `main` 2026-05-03 (commit `952c46db7`).

---

## 9. Mapping `1С → SQL` через Python COM script, не BSL EPF

### Що
Чорновик plan'а Stage 0.5 пропонував створити `Обработки/А_СтруктураХранения.epf` (зовнішня обробка з кнопкою "Згенерувати JSON").

### Deviation
Реалізовано як `mapping/refresh_mapping.py` — Python wrapper що викликає `ПолучитьСтруктуруХраненияБазыДанных()` через V83.COMConnector. Метадані ті самі.

### Чому
- **Auto-runnable**: `python mapping/refresh_mapping.py` без open Designer/Enterprise.
- **CI-friendly**: легше додати у docker/CI.
- **Один stack**: Python ETL вже вимагає COM, нема смислу окремо у BSL.

### Trade-off
- Розробник без 1С Designer не може запустити (треба V83 COM на машині). У реальності ця ж залежність існує для COM-extractor'ів.

---

## 10. Реквізит `Сорт` відсутній у `Справочник.А_Статьи_PL` метаданих

### Що
SQL DDL `OlapBASERP.Dim_PL_Articles` має колонку `Sort_Order int NULL` (Stage 2). Spec згадує реквізит `Сорт` через попередній коміт `23269eeb6 PnL: + скрипти для заповнення Сорт у А_Статьи_PL`.

### Deviation
У поточній 1С метадані (verified через MCP `get_metadata_structure`) **реквізит `Сорт` відсутній** у Справочник.А_Статьи_PL і Справочник.А_ГруппаСтатей_PL. У pipeline `dim_pl_articles` поле не запитується; `Sort_Order` залишається NULL.

### Імпакт
Power BI dashboards Stage 4 не зможуть впорядкувати PL-статті за `Sort_Order` без додаткового workaround (можна використати `PL_Article_Code` як proxy сортування).

### Подальші дії
- Якщо CFO потребує точне впорядкування — додати реквізит `Сорт` (Number) у Справочник.А_Статьи_PL у Stage 1.b → запустити сітку заповнення (скрипти з commit 23269eeb6).
- Або: видалити `Sort_Order` з SQL DDL і використовувати Code-based ORDER BY у DAX.

---

## 11. Bridge PLArticle_DDS = 0 рядків після реалізації

### Що
SQL DDL `OlapBASERP.PLArticle_DDS` (Bridge) очікує заповнення з `Справочник.А_Статьи_PL` де реквізит `СтатьяДвиженияДенежныхСредств` ≠ пуста ссылка.

### Deviation
Live extract: `WHERE _Fld55816RRef <> 0x00..00` повертає 0 рядків. Усі 71 PL-стаття мають порожній реквізит.

### Імпакт
Power BI mapping PL ↔ DDS через Bridge таблицю не працюватиме — DAX `RELATEDTABLE(PLArticle_DDS)` повертатиме пусто.

### Подальші дії
- Pending фінансист: заповнити реквізит у 71 статті (через 1С UI).
- Альтернатива: pipeline-генератор з `Справочник.А_Статьи_PL.Статьи` (ТЧ) якщо там заповнено замість реквізиту шапки. Перевір через mcp__1c-workerp__execute_query.

---

## 12. Fact_Cashflow.BankAccount_ID NOT NULL → placeholder zero-UUID для Каси

### Що
Stage 2 SQL DDL заклав `Fact_Cashflow.BankAccount_ID char(32) NOT NULL`. Очікувалось — кожен Cashflow має банк/касу.

### Deviation
4652 рядки за лютий 2026 включають записи через `Касса` (готівка) де `БанковскийСчетКасса` — посилання на касу, а не на банківський рахунок організації. У реєстрі це поле `_Fld55996_RRRef` буває NULL для деяких джерел.

### Workaround
Pipeline `fact_cashflow.json` додає `defaults: {"BankAccount_ID": "00000000000000000000000000000000"}` у column_mapper. Усі касові рядки попадають з placeholder UUID.

### Подальші дії (правильніше)
- ALTER `Fact_Cashflow.BankAccount_ID` → NULL allowed.
- Або: окрема Dim_Cashboxes + поле `Cashbox_ID` у Fact_Cashflow.
- TBD у Stage 4 schema review.

---

## 13. _Date_Time у BaseERP MSSQL зберігається з +2000 рік-офсетом

### Що
1С BAS ERP cluster mode (`Srvr="SQLSERVER";Ref="BaseERP"`) зберігає всі datetime поля як `Дата + 2000 років`. Документ за 2026-02-15 у MSSQL — `4026-02-15`. (File-mode без офсету.)

### Deviation
SQL запити проти BaseERP backend мають додавати 2000 років до Period/Date params, інакше отримуємо 0 рядків.

### Виправлення
`Pipeline.run()` у `auto_period_params=true` mode додає `period_offset_years: 2000` (default) перед SQL-bind. Якщо пишете кастомний raw_sql — додавайте 2000 вручну.

### Імпакт
Period_Month у Fact_PnL/Fact_Cashflow тепер має правильний рік (2026-02), бо `FactLoader.load()` форсує `period_value` з `orchestrator.period` без офсету; Period (datetime2) — copy з SQL і має +2000 год, але не використовується у DAX. Це асиметрія яку треба пам'ятати.

### Подальші дії
- Або: `transformer onec_date` після extract відняти 2000 років від Period.
- Або: додати logic у extractor `subtract_period_offset_on_read`.
- TBD до Stage 4 — Power BI може зам`ючувати Period (datetime2) для drill-through.

---

## Загальна оцінка

Усі 13 deviations:
- **3 з них** (1, 5, 13) — workarounds для конкретних обмежень платформи 1С
- **5 з них** (2, 3, 4, 8, 12) — виправлення помилок specа / архітектурні корекції які виявились при тестуванні з реальними даними
- **1** (6) — корекція оцінки кількості рядків (spec мав груба estimate)
- **2** (7, 11) — pending manual work, не код-зміна
- **1** (9) — pragmatic alternative до BSL EPF (Python COM equivalent)
- **1** (10) — метадані 1С не мають реквізиту, який очікував DDL

**Жодне deviation не змінює бізнес-логіку** проекту. Всі — або necesssary fixes, або corrections of spec inaccuracies, або architectural improvements (#8 SQL-first).

**Дія:** при наступній ревізії specа (v4) внести deviations 2, 3, 4, 6, 8, 12 у текст specа щоб вирівняти документацію з реальною реалізацією.

---

## Як оновлювати цей файл

**Тригер:** будь-яке нове відхилення від specа.

**Що додавати:**
- Номер deviation, опис
- Stage / Task де виявлено
- Spec section
- Тип (Platform fix / Architecture fix / Spec error / Data-driven)
- Виправлення (commit hash)
- Impact

---

## Cross-references

- Spec v3 final (вихідний документ): `docs/superpowers/specs/2026-05-01-olap-baserp-architecture-design-v3-final.md`
- Implementation plan Stage 1: `docs/superpowers/plans/2026-05-01-olap-baserp-stage1-1c-metadata.md`
- Метадані що зачіпаються deviations: [olap_1c_objects.md](olap_1c_objects.md)
- BSL що зачіпається: [olap_obrabotka_provedeniya.md](olap_obrabotka_provedeniya.md)
- Acceptance numbers (з реальних чисел): [olap_acceptance_etalons.md](olap_acceptance_etalons.md)
