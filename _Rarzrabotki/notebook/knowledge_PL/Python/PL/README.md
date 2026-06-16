# Python-скрипти для побудови knowledge_PL/

Дата створення документа: 2026-04-22  
Призначення: генерують усі 10 файлів `*.md` у `_Rarzrabotki/notebook/knowledge_PL/` для NotebookLM-блокнота **INDUSTRIALBUD_PL** (`af143439-3c76-42f8-a410-6367b5fd609f`).

## Короткий зміст

```
C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_PL\Python\PL\
├── 15_export_to_knowledge_pl.py       # ← ГОЛОВНИЙ скрипт pipeline
├── _erp_query.py                       # COM-обгортка V83.COMConnector → ERP
├── _export_pl_knowledge_helpers.py     # 6 ERP-запитів + markdown-рендерери
├── _compute_pl_aggregates.py           # Топ-N агрегати для FAQ
├── _render_faq.py                      # Рендерер pl_faq.md з агрегатів
├── _pl_aggregates.json                 # Артефакт _compute_pl_aggregates (intermediate)
├── _aggregates_log.txt                 # Лог роботи _compute_pl_aggregates
├── _export_run_log.txt                 # Лог роботи 15_export_to_knowledge_pl
└── test/
    ├── test_erp_connection.py          # Smoke-тест COM-з'єднання
    ├── test_queries.py                 # Тест 6 ERP-запитів
    └── test_render.py                  # Тест markdown-рендерерів
```

Залежність від `_Rarzrabotki/Python/PnL/`:
- `config.py` → `CONN_ERP`, `EXCEL_FILES`, `JSON_DIR` (константи підключення та шляхів).
- `data/json/01_raw_sheets.json` → Excel-коментарі з PL-листів фінансиста.

> **ОНОВЛЕННЯ 2026-06-09.** Цей export-пайплайн читає з 1С (каталог `А_Статьи_PL` + документи `А_ОтчетPL`), тому
> автоматично отримує **очищені** дані: у import-пайплайні (PnL) додано **Шар 4 — реєстр фейк-статей** (структурний
> фільтр `REPORT_END_MARKERS` «все після підсумків = мусор» + `data/fake_articles.json`). Службові рядки фінансиста
> (`Договір підряду …` тощо) більше не потрапляють у каталог/документи → їх не буде і в knowledge-дампах.
> Деталі: `../../pl_pipeline_safeguards.md` (Шар 4), `../../../../Python/PnL/docs/fake_articles_registry.md`.
>
> **Стан даних у 1С (2026-06-09):** А_ОтчетPL 2026 — Січ 26 / Лют 28 / **Бер 30 / Кві 32** (Бер+Кві перезалиті з
> розрізом факту Ф1/Ф2). Knowledge-дампи поки до Лютого 2026 — щоб додати Бер/Кві у NotebookLM, розширити періоди
> у `_compute_pl_aggregates.py` і `_render_faq.py` на `2026-03`/`2026-04` (див. «Сценарій 3» нижче) і прогнати
> `15_export … --period 2026-03` / `--period 2026-04`.
>
> ⚠️ **Граблі import-пайплайну (для контексту):** повторний `02_extract_unique_articles.py` стирає uuid статей →
> у PnL обов'язково повторний `11_upload_articles.py` перед `08`, інакше документи заливаються порожніми. Цей
> export-пайплайн від цього не залежить (читає фінальний стан 1С), але якщо дампи раптом порожні за місяць —
> спершу перевірити, що документи `А_ОтчетPL` за цей місяць не порожні (`ДанныеОтчета`).

## Типовий workflow

### Повна регенерація всіх знань (після закриття нового місяця Excel)

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_PL\Python\PL
python 15_export_to_knowledge_pl.py      # генерує 7 *.md (статика + 3 місяці + 2 дельти)
python _compute_pl_aggregates.py         # обчислює агрегати → _pl_aggregates.json
python _render_faq.py                    # генерує pl_faq.md з агрегатів
```

Після цього — завантажити оновлені файли в NotebookLM через MCP `notebook_add_text` (попередньо видалити старі через `source_delete`).

### Тільки один місяць

```bash
python 15_export_to_knowledge_pl.py --period 2026-03 --skip-static
```

### Smoke-тест (перед великим запуском)

```bash
python test\test_erp_connection.py      # чи працює COM?
python test\test_queries.py             # чи виконуються усі 6 запитів?
python test\test_render.py              # чи валідні markdown-рендерери?
```

---

## Опис кожного скрипта

### 15_export_to_knowledge_pl.py — ГОЛОВНИЙ pipeline

**Робить:** генерує 7 файлів знань у `knowledge_PL/`:
- `pl_articles_catalog.md` — каталог А_Статьи_PL
- `pl_dds_mapping.md` — матриця PL ↔ ДДС
- `pl_dump_2025_12_december.md`, `pl_dump_2026_01_january.md`, `pl_dump_2026_02_february.md` — помісячні виписки (план + факт + каса + документи)
- `pl_dump_delta_jan2026_vs_dec2025.md`, `pl_dump_delta_feb2026_vs_jan2026.md` — дельти між місяцями

**Входи:**
- 1С БД BaseERP через `_erp_query` (COM) — довідник, маппінг, документ А_ОтчетPL, регістри ПрочиеРасходы/ПрочиеДоходы/ФинансовыеРезультаты/А_ДвиженияДенегИзКазны.
- `_Rarzrabotki/Python/PnL/data/json/01_raw_sheets.json` — Excel-коментарі (готує скрипт `01_extract_excel_to_json.py` в PnL pipeline).

**CLI:**
- `python 15_export_to_knowledge_pl.py` — всі 3 місяці + статика + дельти (час ~40 сек)
- `--period YYYY-MM` — тільки вказаний місяць + дельта з попереднім
- `--static-only` — тільки каталог + маппінг
- `--skip-static` — пропустити статику (швидше, якщо структура довідника не змінилась)

**Залежить від 1С-конфігурації:**
- Імена реквізитів ВСІХ запитів у `_export_pl_knowledge_helpers.py`. Якщо перейменують `А_СтатьяДвиженияДенежныхСредств` на будь-якому ПВХ → треба оновлювати тут.

**Що оновити при змінах 1С:**
- Якщо додали новий тип документа-регістратора до `ПрочиеРасходы` → розширити `Q_FACT_EXPENSES` (секція `ВЫБОР КОГДА Р.Регистратор ССЫЛКА Документ.X`). Те саме для `Q_FACT_INCOME_FR/PD`, `Q_CASH`.
- Якщо перейменовано реквізит `Подразделение.А_НаправлениеДеятельности` → оновити у 4 запитах (plan/rash/doh/cash).

---

### _erp_query.py — COM-обгортка

**Робить:** тонкий wrapper навколо `V83.COMConnector`. Експортує одну функцію `execute_query(query_text, parameters=None, max_rows=None)` що повертає `list[dict]`.

**Чому окремо:** щоб усі запити використовували один і той самий connection pool (глобальна `_CONN`), один patern conversion COM-значень (дати, refs, bool), одне місце для діагностики.

**Залежить від 1С-конфігурації:** практично ні — тільки від `config.CONN_ERP` (рядок підключення). Якщо міняється ім'я/пароль SQL-сервера чи бази — міняти у `_Rarzrabotki/Python/PnL/config.py`.

**Що оновити при змінах 1С:** рядок `CONN_ERP = 'Srvr="..."; Ref="..."; Usr="..."; Pwd="..."'` у `config.py`.

---

### _export_pl_knowledge_helpers.py — 6 ERP-запитів + рендерери

**Робить:** найбільший файл (~60 КБ). Містить:
- **6 SQL-запитів** (рядкові константи Q_*):
  - `Q_CATALOG` — повний `Справочник.А_Статьи_PL`
  - `Q_MAPPING` — ТЧ `Статьи` з ДДС
  - `Q_PLAN` — `Документ.А_ОтчетPL.ДанныеОтчета` (план)
  - `Q_FACT_EXPENSES` — `РегистрНакопления.ПрочиеРасходы.Обороты` з детализацією по регістратору + Контрагент
  - `Q_FACT_INCOME_FR` — доходи з `ФинансовыеРезультаты`
  - `Q_FACT_INCOME_PD` — доходи з `ПрочиеДоходы`
  - `Q_CASH` — `А_ДвиженияДенегИзКазны.Обороты` з детализацією
- **6 fetch-функцій** — обгортають запити + convert параметри.
- **Рендерери markdown** — `render_articles_catalog`, `render_dds_mapping`, `render_month_dump`, `render_delta`.
- **Утиліти** — `fmt_money`, `fmt_date`, `clean_type`, `load_excel_comments`.

**Залежить від 1С-конфігурації: СИЛЬНО.** Це серцевина pipeline. Усі імена реквізитів, типи документів-регістраторів, назви регістрів — тут.

**Що оновити при змінах 1С — детально за типами змін:**

| Зміна у 1С | Що оновити у цьому файлі |
|---|---|
| Новий реквізит у `А_Статьи_PL` який треба показувати у knowledge | `Q_CATALOG` + `render_articles_catalog` |
| Новий реквізит у ТЧ `Статьи` | `Q_MAPPING` + `render_dds_mapping` |
| Перейменовано `СуммаФома1` на `СуммаФорма1` | `Q_PLAN` (рядок `СУММА(ТЧ.СуммаФома1)`) |
| Додали тип документа у `ПрочиеРасходы.Регистратор` | `Q_FACT_EXPENSES` (3 блоки `ВЫБОР КОГДА Р.Регистратор ССЫЛКА`) |
| Перейменовано реквізит `Контрагент` у якомусь документі | Відповідна секція `ВЫРАЗИТЬ(...).Контрагент.Наименование` |
| Перейменовано `Подразделение.А_НаправлениеДеятельности` | 4 запити: Q_PLAN, Q_FACT_EXPENSES, Q_FACT_INCOME_FR, Q_FACT_INCOME_PD, Q_CASH |
| Зміна структури `А_ДвиженияДенегИзКазны` | `Q_CASH` |
| Новий тип документа у `А_ДвиженияДенегИзКазны.ДокументДвиженияКазны` | Q_CASH — додати у 3 блоки CASE (Номер, Дата, Контрагент) |

**Важливий нюанс:** для `.Номер` використовуємо `ВЫРАЗИТЬ(... КАК Строка(14))`, бо різні документи мають різну довжину Номера (див. `memory/doc_number_string_length.md`).

---

### _compute_pl_aggregates.py — агрегатор для FAQ

**Робить:** викликає 4 fetch-функції (plan, fact_expenses, fact_income, cash) для кожного з 3 місяців і обчислює:
- Totals: план / факт / каса приплив/відплив
- Top-15 підрозділів за планом
- Top-15 контрагентів за витратами
- Top-15 PL-статей за планом (з коментарями з Excel)
- Top-15 PL-статей за фактом
- Top-10 документів-регістраторів витрат
- Top-10 касових припливів

**Вихід:** `_pl_aggregates.json` (~60 КБ, intermediate-артефакт).

**Залежить від 1С-конфігурації:** через `_export_pl_knowledge_helpers.py`. Сам по собі не має прямих SQL.

**Що оновити при змінах 1С:** якщо додати новий зріз (наприклад, top-N по документних категоріях), — розширити функцію `aggregate_month` у цьому файлі.

**Період:** hardcoded `["2025-12", "2026-01", "2026-02"]`. При додаванні наступних місяців — розширити масив.

---

### _render_faq.py — рендерер pl_faq.md

**Робить:** читає `_pl_aggregates.json` і рендерить `pl_faq.md` (~60 КБ, 10 розділів з прозою + малі таблиці + Q&A).

**Чому окремо:** щоб не викликати ЕРП, якщо дані вже є у `_pl_aggregates.json` (швидка перегенерація тексту після правок у рендерері без запитів ЕРП).

**Залежить від 1С-конфігурації:** НІ. Рендерер не ходить у ЕРП.

**Що оновити:** текст методологічних Q&A (секція 8) — якщо змінюється методологія CFO. Список довідкових Q&A (секція 9).

**Формат для NotebookLM:** прозові абзаци З ЧИСЛАМИ ПЕРЕД таблицями, FAQ-блоки, короткі section headers як якорі для RAG.

---

## test/ — скрипти перевірки

### test_erp_connection.py

**Робить:** 5 кроків smoke-тесту COM-з'єднання:
1. Перевіряє `import win32com.client`
2. Перевіряє `import _erp_query`
3. `ВЫБРАТЬ 1 КАК Ping` — найпростіший запит
4. `КОЛИЧЕСТВО(*) ИЗ Справочник.А_Статьи_PL` — перевіряє доступ до довідника
5. `КОЛИЧЕСТВО(*) ИЗ Документ.А_ОтчетPL` — перевіряє документ

**Коли запускати:** при будь-якому підозрі на проблему з COM (зависання, помилки).

### test_queries.py

**Робить:** викликає усі 6 fetch-функцій з реальним періодом (за замовчуванням 2026-02) і друкує кількість рядків.

**Коли запускати:** після зміни реквізитів у 1С-конфігурації — перевірити що жоден запит не зламався.

**CLI:** `python test_queries.py --period 2026-02`

### test_render.py

**Робить:** тестує `fmt_money`, `fmt_date`, `clean_type` на крайніх випадках + `render_articles_catalog` і `render_dds_mapping` на фейкових даних.

**Коли запускати:** після правок у функціях render_* у `_export_pl_knowledge_helpers.py`. Не потребує ЕРП.

---

## Типові сценарії змін у 1С і що робити

### Сценарій 1: Додали новий тип документа-регістратора `ПрочиеРасходы`

Приклад: новий документ `Документ.А_СписанняМатеріалівПроекту` починає писати у `ПрочиеРасходы`.

**Що робити:**
1. Відкрий `_export_pl_knowledge_helpers.py`.
2. Знайди `Q_FACT_EXPENSES`.
3. У 3 блоках `ВЫБОР КОГДА Р.Регистратор ССЫЛКА` (`Номер`, `Дата`, `Контрагент`) додай нову гілку:
   ```
   КОГДА Р.Регистратор ССЫЛКА Документ.А_СписанняМатеріалівПроекту 
       ТОГДА ВЫРАЗИТЬ(...).Номер КАК Строка(14)
   ```
4. Запусти `test\test_queries.py --period 2026-02`. Якщо PASS → OK.
5. Повний прогін: `python 15_export_to_knowledge_pl.py`.

### Сценарій 2: Перейменували реквізит `А_СтатьяДвиженияДенежныхСредств` на `А_ДДС`

**Що робити:**
1. `grep` у `_export_pl_knowledge_helpers.py` на `А_СтатьяДвиженияДенежныхСредств` — покаже 6+ місць.
2. Заміни усі на `А_ДДС`.
3. `test\test_queries.py` → `test_render.py` → повний прогін.

### Сценарій 3: Додали новий місяць (березень 2026)

**Що робити:**
1. Відкрий `_Rarzrabotki/Python/PnL/config.py` → додай у `EXCEL_FILES` новий словник з шляхом до Excel-файлу березня.
2. Відкрий `_compute_pl_aggregates.py` → розшир список `["2025-12", "2026-01", "2026-02"]` → додай `"2026-03"`.
3. Відкрий `_render_faq.py` → розшир `UA_MONTHS` і `PERIOD_ORDER` так само.
4. Запусти `01_extract_excel_to_json.py` у PnL pipeline (щоб оновити `01_raw_sheets.json`).
5. Далі повний прогін: `15_export_to_knowledge_pl.py` + `_compute_pl_aggregates.py` + `_render_faq.py`.

### Сценарій 4: Додали нову PL-статтю у довідник

Без коду — тільки регенерація:
```bash
python 15_export_to_knowledge_pl.py --static-only
python _compute_pl_aggregates.py
python _render_faq.py
```

---

## Файли знань що генеруються

Усі у `C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_PL\`:

| Файл | Розмір | Скрипт-генератор | Тип для NotebookLM |
|---|---|---|---|
| `pl_articles_catalog.md` | ~13 КБ | 15_export | ПОСТІЙНИЙ |
| `pl_dds_mapping.md` | ~37 КБ | 15_export | ПОСТІЙНИЙ |
| `pl_dump_2025_12_december.md` | ~900 КБ | 15_export | ЗМІННИЙ |
| `pl_dump_2026_01_january.md` | ~740 КБ | 15_export | ЗМІННИЙ |
| `pl_dump_2026_02_february.md` | ~790 КБ | 15_export | ЗМІННИЙ |
| `pl_dump_delta_jan2026_vs_dec2025.md` | ~14 КБ | 15_export | ЗМІННИЙ |
| `pl_dump_delta_feb2026_vs_jan2026.md` | ~14 КБ | 15_export | ЗМІННИЙ |
| `pl_faq.md` | ~60 КБ | _compute + _render_faq | ЗМІННИЙ (headline для RAG) |
| `pl_methodology.md` | ~20 КБ | РУЧНИЙ (не генерується) | ПОСТІЙНИЙ |
| `pl_report_architecture_analyst.md` | ~11 КБ | РУЧНИЙ (не генерується) | ПОСТІЙНИЙ |
| `KNOWLEDGE_MAP_PL.md` | ~16 КБ | РУЧНИЙ | НЕ завантажувати в NotebookLM (meta) |

---

## Memory-записи пов'язані зі скриптами

- `memory/feedback_com_preference.md` — чому використовуємо Python COM, а не MCP HTTP.
- `memory/doc_number_string_length.md` — `Строка(14)` при `ВЫРАЗИТЬ`.
- `memory/notebooklm_rag_best_practices.md` — чому FAQ-формат + проза працюють для RAG.
