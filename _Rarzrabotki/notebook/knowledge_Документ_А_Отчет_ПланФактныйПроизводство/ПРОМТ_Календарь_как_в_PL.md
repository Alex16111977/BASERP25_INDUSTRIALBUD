# ПРОМТ ДЛЯ НОВОЙ СЕССИИ: привести Calendar витрины OlapFactory к виду, как в PL.pbix

> Скопировать блок «ЗАДАЧА» в чат новой сессии. Всё остальное в этом файле — контекст,
> инварианты и критерии приёмки; ИИ обязан выполнить их сам по этому файлу.

---

## ▼▼▼ КОПИРОВАТЬ В ЧАТ ▼▼▼

```
ЗАДАЧА: привести таблицу Calendar витрины OlapFactory и модели Power BI к тому же виду,
что в C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\PowerBi\PL.pbix

Регламент, контекст и критерии приёмки:
_Rarzrabotki/notebook/knowledge_Документ_А_Отчет_ПланФактныйПроизводство/ПРОМТ_Календарь_как_в_PL.md

Сначала прочитай этот файл целиком и KNOWLEDGE_MAP.md в той же папке, затем FINDINGS.md
(там 21 измеренный факт, включая три правила кавычек, на которых уже ломался PBIP).

Покажи мне план по шагам и развилку по диапазону дат — и только после моего «го» делай.
```

## ▲▲▲ КОНЕЦ БЛОКА ▲▲▲

---

# КОНТЕКСТ

## Что уже построено и работает

Контур «План-факт производства», внедрён 2026-08-01/02, все этапы приняты:

`Документ.А_Отчет_ПланФактныйПроизводство` (проведение) → `РегистрСведений.А_ПланФактПроизводство_Свод`
→ ETL `_Rarzrabotki\Olap_Factory\Py_Olap` → SQL-база **`OlapFactory`** →
`_Rarzrabotki\Olap_Factory\ПланФактВиробництва.pbip`

Модель Power BI: **10 таблиц, 22 меры, 10 связей**. Числа сходятся с регистром 1С до копейки
(План 16 214,750 год / 20 424 422,20 грн, Факт 15 092,000 год / 14 331 720,09 грн,
Виконання 3 930,958 год, ETC 9 036 958,39 грн).

**Полное описание контура — в этой же папке:** `KNOWLEDGE_MAP.md` (точка входа),
`05_olap_factory.md` (витрина и ETL), `06_powerbi_model.md` (модель), `FINDINGS.md` (грабли).

## Что не так с Calendar

| | Сейчас в `OlapFactory` | Эталон в `OlapBASERP` (его и видно в PL.pbix) |
|---|---|---|
| Колонок | **16** | **43** |
| Строк | 1 827 | 1 461 |
| Период | 2024-01-01 … **2028**-12-31 | 2024-01-01 … **2027**-12-31 |
| Чем заполняется | Python-сеялка в `Py_Olap\scripts\apply_ddl.py` | `Ai_Olap\scripts\calendar_dim_olapbaserp.sql` |

**Отсутствуют 29 колонок:** `Квартал`, `year_day`, `month_name`, `month_name_en`,
`month_name_short`, `month_name_short_en`, `week_day_str`, `week_day_str_en`, `week_day_str_ua`,
`week_day_short`, `week_day_short_en`, `year_week`, `Year_Quarter`, `year_quarter_en`, `date_str`,
`date_full_str`, `date_full_str_en`, `month_week`, `month_decade`, `month_last_day`,
`year_last_day`, `days_in_month`, `days_in_year`, `Year_Quarter_Sort`, `Year_Week_Sort`,
`Is_Working_Day`, `Days_To_Month_End`, `month_name_short_ua`, `year_month_short_ua`.

**Два имени у меня нестандартные** и в эталоне называются иначе:
`month_short_ua` → у эталона `month_name_short_ua`; `week_day_ua` → у эталона `week_day_str_ua`.

## Полный состав эталона `OlapBASERP.dbo.Calendar` (43 колонки, порядок как в таблице)

```
 1 date_               datetime NOT NULL   (PK)      23 year_quarter_en     varchar  NOT NULL
 2 year_               int      NOT NULL             24 date_str            varchar  NOT NULL
 3 quarter_            int      NOT NULL             25 date_full_str       nvarchar NULL
 4 Квартал             nvarchar NOT NULL             26 date_full_str_en    varchar  NULL
 5 month_              int      NOT NULL             27 month_week          int      NULL
 6 day_                int      NOT NULL             28 month_decade        int      NULL
 7 year_day            int      NOT NULL             29 month_last_day      nvarchar NULL
 8 week_               int      NOT NULL             30 year_last_day       nvarchar NULL
 9 week_day            int      NOT NULL             31 days_in_month       int      NULL
10 month_name          nvarchar NULL                 32 days_in_year        int      NULL
11 month_name_en       nvarchar NULL                 33 weekend             bit      NULL
12 month_name_ua       nvarchar NULL                 34 Period_Month        date     NOT NULL
13 month_name_short    nvarchar NULL                 35 Period_Quarter      date     NOT NULL
14 month_name_short_en nvarchar NULL                 36 Period_Year         date     NOT NULL
15 week_day_str        nvarchar NULL                 37 Year_Month_Sort     int      NOT NULL
16 week_day_str_en     nvarchar NULL                 38 Year_Quarter_Sort   int      NOT NULL
17 week_day_str_ua     nvarchar NULL                 39 Year_Week_Sort      int      NOT NULL
18 week_day_short      nvarchar NULL                 40 Is_Working_Day      bit      NOT NULL
19 week_day_short_en   nvarchar NULL                 41 Days_To_Month_End   int      NOT NULL
20 year_week           varchar  NOT NULL             42 month_name_short_ua nvarchar NULL
21 year_month          varchar  NOT NULL             43 year_month_short_ua nvarchar NULL
22 Year_Quarter        varchar  NOT NULL
```

## Ключевые файлы

| Что | Путь |
|---|---|
| Эталонный SQL календаря | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\scripts\calendar_dim_olapbaserp.sql` (229 строк, идемпотентный DROP+CREATE+INSERT) |
| Эталонный TMDL календаря | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\PowerBi\Industrial\PL.SemanticModel\definition\tables\Calendar.tmdl` (407 строк) |
| Мой DDL витрины | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap\scripts\ddl_olap_factory.sql` |
| Моя сеялка календаря | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap\scripts\apply_ddl.py` (функция `засеять_календарь`) |
| **Генератор модели — единственный источник истины** | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap\scripts\build_pbip.py` |
| Валидатор модели | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap\scripts\validate_pbip.py` |
| Приёмка витрины | `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap\tests\verify_olap_factory.py` |

Подключения: SQL `sa` на `localhost` (он же `SQLSERVER`), строки в `Py_Olap\.env`.
Python — только через `Py_Olap\.venv\Scripts\python.exe`.

---

# ЧТО НАДО СДЕЛАТЬ

## Шаг 1. Витрина: заменить Calendar на эталонный

Самый надёжный путь — **скопировать живую таблицу**, а не перезапускать `.sql`: файл
`calendar_dim_olapbaserp.sql` в шапке заявляет 40 колонок, а в живой таблице их 43
(`month_name_short_ua` и `year_month_short_ua` добавлены позже) — то есть скрипт устарел.
Базы на одном сервере, поэтому межбазовый запрос работает:

```sql
-- структура: скопировать DDL живой OlapBASERP.dbo.Calendar 1:1
-- данные:
INSERT INTO OlapFactory.dbo.Calendar (<явный список 43 колонок>)
SELECT <тот же список> FROM OlapBASERP.dbo.Calendar;
```

Обновить `ddl_olap_factory.sql` (новое определение таблицы, ASCII-only) и `apply_ddl.py`
(вместо Python-сеялки — копирование из `OlapBASERP`; если её нет, fallback на прежнюю логику
с понятным сообщением).

## Шаг 2. Модель: переписать таблицу Calendar в `build_pbip.py`

Взять за образец `PL.SemanticModel\definition\tables\Calendar.tmdl`. Обязательно перенести:

* `dataCategory: Time` на таблице;
* `isUnique` + `isKey` на `date_` — это и есть «пометить как таблицу дат»;
* **`sortByColumn`** на всех текстовых колонках, иначе месяцы и дни недели сортируются
  по алфавиту: `Квартал`→`quarter_`; `month_name`, `month_name_en`, `month_name_ua`,
  `month_name_short`, `month_name_short_en`→`month_`; `week_day_str`, `week_day_str_en`,
  `week_day_str_ua`, `week_day_short`, `week_day_short_en`→`week_day`;
* служебные колонки скрыть (`isHidden`), оставив видимыми те, что реально нужны в отчёте.

## Шаг 3. Связь и меры

Связь сейчас: `Факти.Period` → `Calendar.Дата`. Если колонка переименовывается в `date_`
(как в эталоне) — **обязательно** поправить `relationships.tmdl` в генераторе.

Четыре меры используют `REMOVEFILTERS('Calendar')` — имя таблицы не меняется, их править
не нужно, но проверить после сборки обязательно.

## Шаг 4. Пересборка и проверка

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap
.venv\Scripts\python.exe scripts\apply_ddl.py
.venv\Scripts\python.exe scripts\build_pbip.py
.venv\Scripts\python.exe scripts\validate_pbip.py
.venv\Scripts\python.exe tests\verify_olap_factory.py
```

---

# 🔴 ИНВАРИАНТЫ — сломать легко, заметить трудно

1. **Правило периода.** `План, год`, `План, к-сть`, `План, грн` и
   `Плановий розрахунок матеріалів, грн` обёрнуты в `REMOVEFILTERS('Calendar')` — они обязаны
   оставаться неподвижными при любом слайсере дат, а `Факт` и `Виконання` — следовать за окном.
   Это перенос правила контура 1С в DAX; проверять после каждой правки модели.
2. **Все таблицы читаются native query с ЯВНЫМ списком колонок.** Коннектор `Sql.Database`
   кеширует схему, и при `[Schema=,Item=]` новые колонки не появятся даже после Full Refresh.
   43 колонки придётся перечислить явно.
3. **При добавлении колонки править ДВА места** в `build_pbip.py`: определение колонки в TMDL
   И список колонок в native query. Пропуск второго = колонка молча не доедет (FINDINGS §21).
4. **Кавычки, три разных правила** (на них уже ломались три раза, FINDINGS §19, §20):
   * TMDL: идентификатор с пробелом/дефисом — в одинарных кавычках (`column 'Назва місяця'`);
     кириллица без пробелов кавычек НЕ требует;
   * DAX: имя таблицы с не-ASCII символами — в одинарных кавычках (`SUM('Факти'[X])`),
     иначе мера уходит в `SemanticError`;
   * ссылки на меры (`[План, год]`) кавычек не требуют даже с запятой.
   Всё это проверяет `validate_pbip.py` — не отдавать файл пользователю без зелёного прогона.
5. **`Ai_Olap` не трогать.** Читать из `OlapBASERP` можно, писать в него и менять его
   `refresh_mapping.py`/пайплайны — нельзя. Это чужой контур PnL/Cashflow/Balance.
6. **Оффсет +2000.** Даты BaseERP хранятся с ним; в витрине он уже снят. Календарь берётся
   из `OlapBASERP`, где даты нормальные — дополнительных преобразований не нужно.

---

# РАЗВИЛКА, которую надо вынести пользователю

**Диапазон дат.** Эталон покрывает 2024-01-01 … 2027-12-31, текущий календарь витрины —
до 2028-12-31. Варианты:

* **скопировать эталон как есть** (2024–2027) — полная идентичность с PL.pbix, но витрина
  теряет 2028 год; сейчас фактов там нет, инвариант «0 строк вне Calendar» не пострадает;
* **расширить до 2028** (или дальше) — надо адаптировать генерацию, зато запас на будущее;
  при этом состав колонок останется эталонным.

Спросить пользователя, **не решать самому**: это про горизонт планирования, а не про технику.

---

# КРИТЕРИИ ПРИЁМКИ

Всё должно быть зелёным, иначе работа не закончена:

1. `OlapFactory.dbo.Calendar` — **43 колонки**, состав и типы совпадают с `OlapBASERP.dbo.Calendar`
   (сверить программно, не глазами); число строк соответствует выбранному диапазону.
2. `tests\verify_olap_factory.py` — проходит; в частности **0 строк факта вне Calendar** и
   по-прежнему 0 FK-сирот по всем связям.
3. `scripts\validate_pbip.py` — 0 ошибок; таблиц 10, мер 22, связей 10.
4. Модель открывается в Power BI Desktop **без ошибок** и без жёлтых треугольников на мерах.
5. Через `powerbi-modeling-mcp` (подключиться к живому экземпляру: `ListLocalInstances` →
   `Connect` → `dax_query_operations Execute`) проверить ЧИСЛАМИ, а не состоянием:
   * весь период: План 16 214,750 год / 20 424 422,20 грн, Факт 15 092 год / 14 331 720,09 грн,
     Виконання 3 930,958, ETC 9 036 958,39;
   * окно `'Calendar'[<колонка даты>]` за 01.06.2026–24.07.2026: План и ETC **не меняются**,
     Факт 15 092 → 4 648 год и 14 331 720,09 → 7 292 005,76 грн,
     Виконання 3 930,958 → 3 497,132.
6. Сортировка в визуале: месяцы идут январь→декабрь, а не по алфавиту (это и есть проверка
   `sortByColumn`).
7. Изменения перенесены в git-ветку и закоммичены; `FINDINGS.md` дополнен, если всплыло
   что-то новое.

---

# ЧЕГО НЕ ДЕЛАТЬ

* ❌ не перезапускать `calendar_dim_olapbaserp.sql` вслепую — он устарел на 3 колонки
  и содержит `USE OlapBASERP` (запишет в чужую базу);
* ❌ не собирать `.pbix` скриптом — формат бинарный, работаем через `.pbip`;
* ❌ не отдавать файл пользователю без зелёного `validate_pbip.py` и без проверки чисел в DAX;
* ❌ не менять состав измерений/ресурсов факта заодно — задача только про календарь;
* ❌ не переименовывать таблицу `Calendar`: на неё завязаны `REMOVEFILTERS` в четырёх мерах.
