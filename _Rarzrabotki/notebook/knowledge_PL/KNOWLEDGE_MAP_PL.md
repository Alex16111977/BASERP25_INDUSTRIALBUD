# KNOWLEDGE MAP PL — Маппінг бази знань NotebookLM для PL-аналітики INDUSTRIALBUD

Notebook ID: af143439-3c76-42f8-a410-6367b5fd609f
Блокнот: INDUSTRIALBUD_PL
Зв'язаний блокнот (розробницький): INDUSTRIALBUD (3303acdb-2d7f-4879-9f13-78705ab3fb8c)
Дата створення маппінгу: 2026-04-22
Дата останнього upgrade: 2026-04-23 (Sec 11 FAQ + split-dumps + contextual metadata)
Фокус: аналітика PL план↔факт для фінансиста з 20-річним досвідом

## Архітектура upgrade v2 (2026-04-23)

| Категорія | Кількість | Призначення |
|---|---|---|
| Статичні довідники | 4 | catalog / dds_mapping / methodology / architecture |
| FAQ | 1 | pl_faq.md (з Sec 11 — 115 precision Q&A) |
| Помісячні dumps | 18 | 3 місяці × 6 buckets (summary/income/cost/opex/marketing/cash) |
| Delta-файли | 2 | Jan vs Dec, Feb vs Jan |
| **У NotebookLM разом** | **25** | |
| Query patterns (довідник на диску) | 1 | pl_query_patterns.md — НЕ завантажено |
| Архів старих моноліт-dumps | 3 | `_archive_pre_split/` — НЕ завантажено |
| **Всього файлів на диску** | **29** | |

## Призначення

Ця база знань відповідає на типові питання фінансиста:
- Що потрапило в PL-статтю X за лютий 2026?
- Які документи, які контрагенти, які суми?
- Що написано в коментарі до цієї статті у Excel?
- Як це змінилось у порівнянні з січнем 2026 / груднем 2025?
- Які ДДС-статті входять у PL-статтю і чому саме?

## Python-скрипти що будують цю базу знань

Усі Python-скрипти, що генерують файли у цій папці, розміщені разом зі знаннями у підпапці:

📂 **`Python/PL/`** — production-скрипти pipeline  
📂 **`Python/PL/test/`** — smoke-тести для перевірки після змін у 1С  
📄 **`Python/PL/README.md`** — **ДЕТАЛЬНИЙ опис кожного скрипта** (що робить, коли запускати, від яких елементів 1С-конфігурації залежить, що оновити при змінах)

**Коли ІІ треба змінити логіку збору знань** (наприклад, у 1С додали новий тип документа-регістратора, або перейменували реквізит):
1. Спочатку відкрий `Python/PL/README.md` → знайди відповідний сценарій змін (розділ «Типові сценарії змін у 1С і що робити»).
2. Відкрий відповідний скрипт і онови SQL-запити чи рендерер.
3. Запусти `Python/PL/test/test_queries.py` для валідації.
4. Повний прогін: `Python/PL/15_export_to_knowledge_pl.py`.

**Короткий перелік скриптів:**
- `15_export_to_knowledge_pl.py` — ГОЛОВНИЙ; генерує **20 md-файлів** (2 статичні catalog/mapping, 18 помісячних split-за-bucket, 2 дельти).
- `_erp_query.py` — COM-обгортка навколо `V83.COMConnector`; одна точка під'єднання.
- `_export_pl_knowledge_helpers.py` — 6 SQL-запитів до 1С + markdown-рендерери з **contextual metadata prefix** і **split-за-PL-групою** на 6 buckets. **Найбільш залежний від 1С-конфігурації.**
- `_compute_pl_aggregates.py` — топ-N агрегати по 3 місяцях → `_pl_aggregates.json`.
- `_render_faq.py` — рендерить `pl_faq.md` з агрегатів, містить **Sec 11: 115 precision Q&A** з точними сумами у 3 форматах (точне число + округлення + пропис словами) для максимального RAG retrieval-матчу.
- `test/test_erp_connection.py`, `test/test_queries.py`, `test/test_render.py` — smoke-тести.

## Як користуватись цим файлом

Коли користувач каже "обнови PL-знання" або "перевір актуальність PL-знань":

1. Прочитай цей файл повністю + `Python/PL/README.md`
2. Для кожного файлу знань порівняй "Дата генерації" з поточними датами:
   - Excel у `_Rarzrabotki/PL/<місяць>/*.xlsx`
   - Документ.А_ОтчетPL у 1С (Period)
   - Справочник.А_Статьи_PL (структура)
3. Якщо джерело змінилось — перегенеруй відповідний файл:
   - Статичні (каталог, маппінг): `python Python/PL/15_export_to_knowledge_pl.py --static-only`
   - Помісячні: `python Python/PL/15_export_to_knowledge_pl.py --period YYYY-MM`
   - Повний перезапуск: `python Python/PL/15_export_to_knowledge_pl.py` (+ потім `_compute_pl_aggregates.py` + `_render_faq.py` для FAQ)
4. Видали старе джерело з NotebookLM (source_delete з source_id)
5. Завантаж оновлене (notebook_add_text)
6. Онови "Дата генерації" та "NotebookLM source ID" в цьому файлі

## Потік даних

```
Excel (PL/{місяць}/*.xlsx)        ── фінансист готує план + коментарі
    │
    ├── PnL pipeline (scripts/01-14) ── імпорт у 1С
    │
    ▼
Документ.А_ОтчетPL.ДанныеОтчета   ── план у 1С, прив'язаний до Підрозділу/Напрямку
    │
    │  паралельно:
    ├── РегистрНакопления.ПрочиеРасходы, ПрочиеДоходы, ФинансовыеРезультаты ── факт ЕРП
    ├── РегистрНакопления.А_ДвиженияДенегИзКазны                             ── каса Казни
    │
    ▼
Отчет.А_ОтчетPL (СКД, 4 джерела, 7 CTE) ── фінансова сверка у UI 1С
    │
    ▼
15_export_to_knowledge_pl.py (новий скрипт) ── виписки у markdown
    │
    ▼
knowledge_PL/pl_dump_*.md  ── джерела для NotebookLM
    │
    ▼
NotebookLM INDUSTRIALBUD_PL (af143439-...) ── Q&A для фінансиста
```

### Зв'язок з основним блокнотом

Блокнот `INDUSTRIALBUD` (розробницький) має розробницький контекст:
- `a_otchet_pl_overview.md` — як влаштовані 7 CTE і 3 варіанти СКД
- `a_otchet_pl_adr_parity.md` — правила адаптивної уникальності ДДС
- `pnl_pipeline_architecture.md` — перелік 14+2 скриптів pipeline
- `registers_income_expense.md` — структура 7 регістрів

Блокнот `INDUSTRIALBUD_PL` (цей) — аналітичний фокус:
- Семантика статей для фінансиста (не CTE, не SKD-параметри)
- Номери документів, контрагенти, суми за період
- Коментарі з Excel + порівняння з 1С
- Дельти між місяцями

---

## Файли знань

> **Увага:** після upgrade 2026-04-23 архітектура змінилася з 10 моноліт-файлів на **25 split-файлів у NotebookLM + 4 на диску**. Старі source_id замінюються у процесі swap (див. розділ «NotebookLM swap log» нижче).

### A. Статичні довідники (4 файли, ПОСТІЙНІ)

Оновлюються рідко — при змінах конфігурації 1С. Містять contextual metadata prefix для AI-retrieval.

| # | Файл | NotebookLM title (keyword-rich) | Призначення |
|---|---|---|---|
| A1 | `pl_articles_catalog.md` | "PL Articles Catalog ІНДАСТРІАЛБУД — 68 статей 8 груп whitelist автосоздані" | Повна ієрархія Справочник.А_Статьи_PL: групи, коди, ТипСтатьи. Питання: які статті обов'язкові з ДДС, які whitelist (000000028, 000000056), які автосоздані. |
| A2 | `pl_dds_mapping.md` | "PL ↔ ДДС Mapping — матриця відповідностей Статьи_PL shared zontik" | Matrix PL↔ДДС: forward view (PL→ДДС), reverse view (ДДС→PL), uncovered ДДС, правило адаптивної уникальності. |
| A3 | `pl_report_architecture_analyst.md` | "PL звіт А_ОтчетPL архітектура аналітик — План Факт Каса ЦО СВОД data gap" | Архітектура звіту для фінансиста (без BSL): 3 колонки, коли легально розходяться, ЦО vs СВОД. |
| A4 | `pl_methodology.md` | "PL методологія ІНДАСТРІАЛБУД — що означає кожна стаття штраф ОС ЗП ЄСВ ПДФО" | Економічна семантика 8 груп і 68 статей. Антикейси (штрафи не в Собівартість). |

**Як оновити A1, A2:** `python Python/PL/15_export_to_knowledge_pl.py --static-only`  
**Як оновити A3, A4:** ручне редагування (рідко).

### B. FAQ (1 файл, ЗМІННИЙ — перегенерується після закриття місяця)

| # | Файл | NotebookLM title | Призначення |
|---|---|---|---|
| B1 | `pl_faq.md` | "PL FAQ ІНДАСТРІАЛБУД — 115+ precision Q&A точні суми методологія" | Прозові відповіді на типові питання (Sec 1-10) + **Sec 11: 115 precision Q&A з точними сумами** (TL;DR, топ-5 підрозділів/контрагентів/статей/документів/каси, delta, агрегат 3 міс). ~140 КБ. |

**Як оновити:**
```bash
python Python/PL/_compute_pl_aggregates.py   # → _pl_aggregates.json
python Python/PL/_render_faq.py              # → ../pl_faq.md
```

### C. Помісячні dumps (18 файлів, ЗМІННІ — split-за-bucket для RAG)

Для **кожного з 3 місяців** генерується 6 файлів split-за-PL-групою:

| Bucket | Файл | NotebookLM title pattern | Розмір (орієнт.) |
|---|---|---|---|
| summary | `pl_dump_{YYYY_MM}_01_summary.md` | "PL {Місяць Рік} INDUSTRIALBUD Summary — План {X}М Факт {Y}М Каса {Z}млрд Топ-підрозділ {назва}" | 6-7 КБ |
| income_revenue | `pl_dump_{YYYY_MM}_02_income_revenue.md` | "PL {Місяць Рік} INDUSTRIALBUD Доходи — виручка продажі донори ООН ВООЗ ПРООН" | 50-70 КБ |
| cost_of_goods | `pl_dump_{YYYY_MM}_03_cost_of_goods.md` | "PL {Місяць Рік} INDUSTRIALBUD Собівартість — матеріали ЗП виробнич. ІТР ПМ" | 800-930 КБ ⚠ |
| opex_admin | `pl_dump_{YYYY_MM}_04_opex_admin.md` | "PL {Місяць Рік} INDUSTRIALBUD Opex Адмін — загальновиробн. оренда банки ІТ" | 300-360 КБ |
| marketing_fin | `pl_dump_{YYYY_MM}_05_marketing_fin.md` | "PL {Місяць Рік} INDUSTRIALBUD Маркетинг Фінансова — реклама відсотки курсові" | 50-85 КБ |
| cash_anomalies | `pl_dump_{YYYY_MM}_06_cash_anomalies.md` | "PL {Місяць Рік} INDUSTRIALBUD Каса аномалії — Казна приплив відплив ЕРП БЕЗ PL" | 11 КБ |

**Місяці в роботі:** 2025-12 (Грудень 2025), 2026-01 (Січень 2026), 2026-02 (Лютий 2026) — **18 файлів разом**.

⚠ Cost-файли ~900 КБ — все одно велико для оптимального RAG. Планується подальший split (ЗП / Матеріали / Інше) у наступному upgrade v3.

**Як оновити:**
```bash
python Python/PL/15_export_to_knowledge_pl.py --period 2026-02
# → 6 файлів за місяць автоматично
```

### D. Delta-файли (2 файли, ЗМІННІ)

| # | Файл | NotebookLM title | Призначення |
|---|---|---|---|
| D1 | `pl_dump_delta_jan2026_vs_dec2025.md` | "PL delta Січень 2026 vs Грудень 2025 — Top moves нові зниклі контрагенти" | Динаміка за 2 місяці: топ-10 статей-змін, нові/зниклі контрагенти, порівняння коментарів Excel. |
| D2 | `pl_dump_delta_feb2026_vs_jan2026.md` | "PL delta Лютий 2026 vs Січень 2026 — Top moves нові зниклі контрагенти" | Те саме, Feb vs Jan. |

**Як оновити:** `python Python/PL/15_export_to_knowledge_pl.py --period 2026-02` (delta буде оновлена автоматично, якщо поперед. місяць вже згенеровано).

### E. Файли на диску (НЕ у NotebookLM)

| # | Файл | Призначення |
|---|---|---|
| E1 | `pl_query_patterns.md` | Довідник для фінансиста: **20 шаблонів запитів до NotebookLM** (✓ працює / ✗ не працює). Містить антипаттерни CoT-leak, quick reference таблицю «де що шукати». Створено 2026-04-23. |
| E2 | `_archive_pre_split/pl_dump_2025_12_december.md` | Архівна моноліт-версія грудня 2025 (900+ КБ). |
| E3 | `_archive_pre_split/pl_dump_2026_01_january.md` | Архів січень 2026. |
| E4 | `_archive_pre_split/pl_dump_2026_02_february.md` | Архів лютий 2026. |

**Цей файл (`KNOWLEDGE_MAP_PL.md`):** теж на диску, але **НЕ завантажено у NotebookLM** (раніше завантажували → плутало RAG, тестами підтверджено; див. memory `notebooklm_rag_best_practices.md`).

---

## NotebookLM swap log (статус source_id)

### Старі source_id (видалено 2026-04-23 під час upgrade v2):

| Файл старий | Старий source_id | Статус |
|---|---|---|
| pl_articles_catalog.md | efc43aa3-1c74-42a4-9d30-d988d8984b89 | DELETED |
| pl_dds_mapping.md | 12eb6866-e83f-4925-a404-00183ab385e0 | DELETED |
| pl_report_architecture_analyst.md | a724d03f-a91c-40af-a6d3-3d2ae1dccb9b | DELETED |
| pl_methodology.md | 4b8ed14a-aaae-4aae-961a-04497ca3625d | DELETED |
| pl_dump_2025_12_december.md (моноліт) | 1f5c5a2a-f584-43da-a99a-5cdd730bd475 | DELETED |
| pl_dump_2026_01_january.md (моноліт) | aed00671-0a9e-4575-9106-9626efccaac8 | DELETED |
| pl_dump_2026_02_february.md (моноліт) | 3fc8903d-9398-49d4-aecc-4f5c5d77daeb | DELETED |
| pl_dump_delta_jan2026_vs_dec2025.md | 89fc7b1f-6746-4012-8d81-597491d32bbc | DELETED |
| pl_dump_delta_feb2026_vs_jan2026.md | 9b59e904-5303-4247-8e5c-0707803f73e2 | DELETED |
| pl_faq.md | 18804978-526c-4e86-900b-bb53f1eb7e5d | DELETED |

### Нові source_id (завантажено 2026-04-23, 25 sources):

| Код | Файл | NotebookLM source_id |
|---|---|---|
| A1 | pl_articles_catalog.md | 18f0161d-ac61-4d20-9b44-fdfe1a3c72b2 |
| A2 | pl_dds_mapping.md | eac0beda-759f-4c94-8ad3-aa8119dac5d1 |
| A3 | pl_report_architecture_analyst.md | ffb2f9e3-cce3-4e30-b36d-10f11bfe44f4 |
| A4 | pl_methodology.md | 480663ff-07f3-4866-a557-38c661eb37a0 |
| B1 | pl_faq.md (v2 — fix Sec 1 dump references) | f84323c7-388c-43be-9ae5-0e5a47fac1d8 |
| C-Dec-01 | pl_dump_2025_12_01_summary.md | a7750342-02e6-4e5a-be8e-0e784fd69467 |
| C-Dec-02 | pl_dump_2025_12_02_income_revenue.md | 6b59c012-dc85-4956-8d77-8e44ebabdf0a |
| C-Dec-03 | pl_dump_2025_12_03_cost_of_goods.md | 0479a386-66ac-4efc-9fa4-46e0b9fc69a1 |
| C-Dec-04 | pl_dump_2025_12_04_opex_admin.md | dfabd4b7-f241-49da-9d1c-cee1db1aa3eb |
| C-Dec-05 | pl_dump_2025_12_05_marketing_fin.md | 76eeb353-6280-4038-9455-e6ea027cd253 |
| C-Dec-06 | pl_dump_2025_12_06_cash_anomalies.md | f23fa034-254d-4b55-827c-b2a96401a93b |
| C-Jan-01 | pl_dump_2026_01_01_summary.md | a7ba04cc-5a14-4bef-8bd9-075d37c68aa0 |
| C-Jan-02 | pl_dump_2026_01_02_income_revenue.md | af20a6ee-426c-4800-a405-2a9b4cf15d2b |
| C-Jan-03 | pl_dump_2026_01_03_cost_of_goods.md | 2b867011-d2d2-4bff-a48d-ce733cd5ac33 |
| C-Jan-04 | pl_dump_2026_01_04_opex_admin.md | 4ba449ef-a9e5-4b1c-ade5-cb92662c2e09 |
| C-Jan-05 | pl_dump_2026_01_05_marketing_fin.md | bc56ee78-2297-4e06-9f76-6585497e9b81 |
| C-Jan-06 | pl_dump_2026_01_06_cash_anomalies.md | 6f22afd9-5296-491a-8b19-fe94351f8205 |
| C-Feb-01 | pl_dump_2026_02_01_summary.md | 3a65f1e0-0f3e-4f40-a32c-bada44bc53b1 |
| C-Feb-02 | pl_dump_2026_02_02_income_revenue.md | e03748f6-d38a-4e52-89ce-8a0d11ce865f |
| C-Feb-03 | pl_dump_2026_02_03_cost_of_goods.md | 6e9d64ad-e158-444d-859e-40c3c4a64cb7 |
| C-Feb-04 | pl_dump_2026_02_04_opex_admin.md | b900a5f6-625e-4d82-aeff-8f37ca37355d |
| C-Feb-05 | pl_dump_2026_02_05_marketing_fin.md | 728e321e-5b38-4793-8420-49917ad5d706 |
| C-Feb-06 | pl_dump_2026_02_06_cash_anomalies.md | d0be0aa6-c036-4b27-872c-e1ad9a513051 |
| D1 | pl_dump_delta_jan2026_vs_dec2025.md | 04489c5c-9a31-44d1-a06c-c5f580c46874 |
| D2 | pl_dump_delta_feb2026_vs_jan2026.md | 66a5610a-66d3-4875-936b-bc0013f30bcf |

**Upgrade v2 завершено 2026-04-23.** Наступний крок — re-тест 10 питань для валідації target PASS rate ≥ 95%.

---

## Типи файлів знань

ПОСТІЙНИЙ — методологія, довідники, архітектура. Оновлювати при структурних змінах (раз на квартал або при зміні конфігурації).
ЗМІННИЙ — помісячні виписки і дельти. Оновлювати після закриття кожного місяця фінансистом.

## Regeneration команди

Усі скрипти — у `knowledge_PL/Python/PL/`. Детальний опис: `Python/PL/README.md`.

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_PL\Python\PL

# Smoke-тест перед повним запуском (рекомендовано після змін у 1С)
python test\test_erp_connection.py
python test\test_queries.py --period 2026-02

# Повний перезапуск (усі статичні + усі 3 місяці + дельти)
python 15_export_to_knowledge_pl.py

# Тільки один місяць
python 15_export_to_knowledge_pl.py --period 2026-02

# Тільки статика (каталог + маппінг)
python 15_export_to_knowledge_pl.py --static-only

# Без статики (швидше при оновленні тільки одного місяця)
python 15_export_to_knowledge_pl.py --period 2026-02 --skip-static

# FAQ (окремий крок — агрегати + рендер)
python _compute_pl_aggregates.py   # → _pl_aggregates.json (~40 сек)
python _render_faq.py              # → ../pl_faq.md (~1 сек)
```

## Як ІІ повинен оновлювати знання PL

Коли користувач каже "обнови PL":
1. Прочитати цей файл (KNOWLEDGE_MAP_PL.md)
2. Визначити які місяці треба оновити (за датою зміни Excel)
3. Запустити `15_export_to_knowledge_pl.py` з відповідними параметрами
4. Для кожного оновленого файлу:
   a. `source_delete(notebook_id=af143439-..., source_id=старий)`
   b. `notebook_add_text(notebook_id=af143439-..., title="<назва>", text=<вміст файлу>)`
   c. Оновити source_id та дату генерації у цьому файлі
