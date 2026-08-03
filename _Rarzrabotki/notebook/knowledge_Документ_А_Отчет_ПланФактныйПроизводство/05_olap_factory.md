# Витрина OlapFactory и ETL-проект Py_Olap (План-факт производства)

Файл описывает **третье и четвёртое звено** контура: SQL-витрину `OlapFactory` и ETL-проект,
который её наполняет.

```
Документ.А_Отчет_ПланФактныйПроизводство (проведение)
  → РегистрСведений.А_ПланФактПроизводство_Свод   (_InfoRg56577)
  → ETL Py_Olap → база OlapFactory (Fact + Dim + Calendar)
  → vw_Fact_PlanFact → Power BI (ПланФактВиробництва.pbip)
```

Отчёт 1С `А_ПланФактныйПроизводствоПолный` в контур не входит — он остался эталоном сверки.

---

## 1. Почему отдельный проект и отдельная база

Из `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\README.md` (строки 3–5):
контур **независим от `_Rarzrabotki/Olap/Ai_Olap`** — своя база `OlapFactory`, свой ETL-проект,
свой WHITELIST маппинга.

Что это снимает: перегенерация `mapping/baserp_storage.json` здесь **не может сдвинуть `_FldNNN`
у пайплайнов PnL / Cashflow / Balance** — те читают свой JSON в своём проекте. Список пайплайнов
по умолчанию в `Py_Olap/main.py` (строки 59–62) содержит ровно два id: `dim_factory` и
`fact_planfact_proizvodstvo`, то есть `main.py` без флагов не трогает ничего чужого.

Движок (`ai_olap/`) — клон движка Ai_Olap; отличаются только `pipelines/*.json`, `mapping/`,
`scripts/`, `tests/` и список `ALL_DEFAULT_PIPELINES`.

## 2. Структура каталога

`C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\`

| Путь | Что это |
|---|---|
| `Py_Olap\.env` | `BASERP_SQL_DSN`, `OLAP_SQL_DSN`, `BASERP_COM_CONN`, `LOG_LEVEL`, `MAPPING_PATH` (шаблон — `.env.example`) |
| `Py_Olap\ai_olap\` | движок: `extractors / transformers / loaders / orchestrator / utils / core` |
| `Py_Olap\mapping\refresh_mapping.py` | WHITELIST из 12 объектов 1С → `baserp_storage.json` |
| `Py_Olap\pipelines\dim_factory.json` | 7 справочников-измерений, `full_reload` |
| `Py_Olap\pipelines\fact_planfact_proizvodstvo.json` | факт из `_InfoRg56577` |
| `Py_Olap\scripts\ddl_olap_factory.sql` + `apply_ddl.py` | DDL витрины + зеркалирование Calendar из `OlapBASERP` + сев Dim_RowTypes |
| `Py_Olap\scripts\introspect_planfakt_fields.py` | печать физических `_InfoRg / _Fld` из маппинга |
| `Py_Olap\scripts\build_pbip.py` | генератор модели Power BI (TMDL) |
| `Py_Olap\tests\verify_olap_factory.py` | приёмка витрины: SQL == 1С |
| `ПланФактВиробництва.pbip / .SemanticModel / .Report` | проект Power BI |

Подключения (`Py_Olap\ai_olap\core\connections.py`): к `BaseERP` — pyodbc **`readonly=True`**,
к `OlapFactory` — pyodbc с `autocommit=False` и `fast_executemany`. COM к ERP используется только
в `refresh_mapping.py` и в приёмке.

## 3. Объекты витрины OlapFactory

DDL: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap\scripts\ddl_olap_factory.sql`

| Объект | Тип | Ключ / примечание |
|---|---|---|
| `Fact_PlanFactProizvodstvo` | таблица | `Fact_ID bigint IDENTITY` PK; 8 ресурсов; индексы `IX_Fact_PFP_Period (Period_Month, Department_ID)`, `IX_Fact_PFP_RowType (RowType)` |
| `Dim_Organizations` / `Dim_Departments` / `Dim_Stages` / `Dim_Items` / `Dim_CommonNames` / `Dim_Units` / `Dim_Individuals` | таблицы | PK `char(32)` — hex-UUID 1С без дефисов |
| `Dim_RowTypes` | таблица | PK `RowType varchar(20) COLLATE Cyrillic_General_CI_AS`, + `RowType_Sort`, `RowType_Group` |
| `Calendar` | таблица | PK `date_ datetime`; **зеркало `OlapBASERP.dbo.Calendar` 1:1** — 43 колонки, 2024-01-01 … 2027-12-31 = 1 461 день (см. §3.1) |
| `Dim_Documents` | таблица | первичка (закупка / табель / ВРМ / карточка СС) + `ObjectDept_ID` — объект 15м/30м, см. §3.2 |
| `ETL_Runs` | таблица | журнал прогонов: `Script`, `Period_Month`, `Rows_Loaded`, `Status`, `Error` |
| `vw_Fact_PlanFact` | представление | факт + разрешённая аналитика + `ObjectDept_ID` |
| `Dim_Departments_Tree` | представление | `Level1..Level5` + `Hierarchy_Path` |

🔴 **Внешних ключей (FOREIGN KEY) в DDL нет ни одного.** Целостность проверяется не БД,
а скриптом приёмки (`verify_olap_factory.py`, блок 6.5). Если вы добавите FK-констрейнты —
сломается `DimLoader`: он делает `TRUNCATE TABLE`, а TRUNCATE запрещён для таблицы, на которую
ссылается FK.

Числа замера 2026-08-01 (README, строки 66–83): Fact 4 396 строк, Dim_Items 42 603,
Dim_Departments 413, Dim_CommonNames 322, Dim_Individuals 764, Dim_Units 63,
Dim_Organizations 11, Dim_Stages 10, Dim_RowTypes 4, Calendar 1 827
(с 02.08.2026 Calendar — **1 461**, см. §3.1).

### 3.1 🔴 Calendar — зеркало эталона, а не собственная сеялка

С 02.08.2026 `OlapFactory.dbo.Calendar` — **точная копия живой `OlapBASERP.dbo.Calendar`**:
43 колонки, те же имена и типы, 1 461 строка за 2024-01-01 … 2027-12-31. Это та самая таблица,
которую видно в `PowerBi\Industrial\PL.pbix`; благодаря копии обе модели считают и подписывают
даты одинаково. Прежняя Python-сеялка на 16 колонок (2024…2028) осталась только аварийным
путём — если соседняя база недоступна.

Механика — `apply_ddl.py`, функция `синхронизировать_календарь()`:
структура читается из `OlapBASERP.INFORMATION_SCHEMA.COLUMNS` и воспроизводится один в один
(при расхождении таблица пересоздаётся), строки копируются межбазовым `INSERT … SELECT`
с явным списком колонок. Обе базы на одном сервере, поэтому межбазовый запрос работает.

Чего не делать:
* ❌ не заводить состав колонок руками в `ddl_olap_factory.sql` — у эталона есть колонка с
  кириллическим именем `Квартал`, а файл держится ASCII-only; и любая ручная копия устареет,
  как устарел сам эталонный скрипт;
* ❌ не перезапускать `Olap\Ai_Olap\scripts\calendar_dim_olapbaserp.sql` — он отстал от живой
  таблицы на 3 колонки (`month_name_short_ua`, `year_month_short_ua` добавлены позже) и
  содержит `USE OlapBASERP`, то есть пишет в **чужой** контур PnL/Cashflow/Balance;
* ❌ не писать в `OlapBASERP` вообще — только чтение.

Сверяется программно: `verify_olap_factory.py`, блок 6.8 — состав, типы, порядок колонок,
число строк и посимвольное совпадение всех значений (`EXCEPT` по всем 43 колонкам).

### 3.2 🔴 ПодразделениеОбъекта (15м/30м) выводится, а не хранится

Эталонный отчёт `А_ПланФактныйПроизводствоПолный` имеет разрез `ПодразделениеОбъекта`
(титул «Підрозділ-об'єкт»), а в регистре свёртки его нет и не будет — обоснование в
`02_registr_struktura.md` §8. С 03.08.2026 разрез всё равно доступен в Power BI, потому что
**выводится из уже загруженных данных**, без единой правки регистра и без перепроведения:

`ПодразделениеОбъекта` — реквизит карточки `Справочник.А_СтруктураСебестоимости`, а все
не-фактовые ветки свёртки жёстко привязаны к конкретной карточке. Отсюда цепочка:

| Тип строки | Как получаем объект | Заполнено |
|---|---|---|
| План | измерение `Документ` = сама карточка СС → её `ПодразделениеОбъекта` | да |
| Виконання | измерение `Документ` = ВРМ → реквизит шапки `СтруктураСебестоимости` → карточка | да |
| Факт | нет источника — и не должно быть: в отчёте у факта объект пуст («котёл» на итоге подразделения) | нет |
| ПланНаФакт | ETC считается по (Подразделение, общее название) без СС — дробить нечем | нет |

Реализация: колонка `Dim_Documents.ObjectDept_ID` (наполняется в `raw_sql` шага `dim_documents`:
`_Reference56133._Fld56558RRef` напрямую, `_Document56405` — через `_Fld56407RRef`), а
`vw_Fact_PlanFact` подтягивает её `LEFT JOIN`-ом по `Document_ID`. В факт-таблице колонки нет
намеренно: связь «строка → карточка СС» уже хранится в измерении `Документ`, дубль стал бы
вторым источником правды.

Контроль — `verify_olap_factory.py`, блок 6.9: заполненность по типам строк сверяется с этой
таблицей, а сумма плана по объектам обязана совпасть с планом подразделения.

## 4. Разбор raw_sql факта

`pipelines\fact_planfact_proizvodstvo.json`, строка 11 — один `SELECT` из одной таблицы:

```sql
FROM _InfoRg56577 r WHERE r._Active = 0x01 AND r._Period >= ? AND r._Period < ?
```

Соответствие проверено по `mapping\baserp_storage.json` (сгенерирован 2026-08-01T23:04:21,
платформа в поле `version`):

| Колонка витрины | Физическое поле | Поле регистра 1С |
|---|---|---|
| `Recorder_ID` | `_RecorderRRef` | Регистратор |
| `Period` | `_Period` | Период |
| `Organization_ID` | `_Fld56578RRef` | Организация |
| `Department_ID` | `_Fld56579RRef` | Подразделение |
| `ExecutorDept_ID` | `_Fld56580RRef` | ПодразделениеИсполнитель |
| `RowType` | `_Fld56581RRef` | ТипСтроки |
| `Block` / `Category` | `_Fld56582` / `_Fld56583` | Блок / Категория |
| `Stage_ID` | `_Fld56584RRef` | Этап |
| `Work_ID` | `_Fld56585RRef` | Работа (тип — Номенклатура, поэтому FK на `Dim_Items`) |
| `CommonName_ID` | `_Fld56586RRef` | ОбщееНазвание |
| `Analytics_ID` / `Analytics_Text` | `_Fld56587_RRRef` / `_Fld56587_S` | Аналитика (составной тип) |
| `Document_ID` | `_Fld56588_RRRef` | Документ (составной тип) |
| `Unit_ID` | `_Fld56597RRef` | Единица |
| `ItemName` | `_Fld56598` | НоменклатураНаименование |
| `PlanHours … ETC_UAH` | `_Fld56589 … _Fld56596` | ПланЧасы, ФактЧасы, ВиконанняЧасы, ПланКол, ФактКол, ПланГрн, ФактГрн, ПланНаФактГрн |

Оговорка: в `baserp_storage.json` для составных реквизитов присутствует только `_RRRef`-часть
(`Аналитика → _Fld56587_RRRef`). Имя `_Fld56587_S` (строковая часть составного типа) в маппинге
**не фигурирует** — оно выведено из платформенной раскладки `_TYPE/_S/_RRRef`; при перегенерации
маппинга проверять его отдельно.

**Почему JOIN к документу-регистратору не нужен** (описание пайплайна, строка 3): регистр
периодический — у него собственное поле `_Period`, и вся свёртка уже сделана в 1С при проведении
документа. Дата берётся из регистра, а не из шапки документа; `_RecorderRRef` тянется просто как
атрибут (drill-back в 1С), соединение к `_Document56571` не выполняется. Отбор `_Active = 0x01` —
только активные записи, как в запросах 1С по умолчанию.

## 5. Цепочка трансформеров

Порядок из пайплайна (строка 14): `varbinary_to_uuid → enum_resolver → period_offset_fix →
column_mapper`. Реализация — `Py_Olap\ai_olap\transformers\`.

1. **`varbinary_to_uuid`** — `VARBINARY(16)` → `char(32)` lowercase hex; 16 нулевых байт → `None`
   (чтобы FK не ссылались в пустоту); 1 байт → `bool`.
2. **`enum_resolver`** — UUID перечисления → строка из `FROZEN_ENUMS`. Карта строится один раз
   на процесс: `SELECT _IDRRef, _EnumOrder FROM _Enum56576` и zip с замороженным списком.
   🔴 Порядок шагов важен: `varbinary_to_uuid` уже превратил пустую ссылку в `None`, поэтому
   `enum_resolver` (строки 278–283) подставляет для неё литерал `"ПустаяСсылка"`, а не NULL —
   `Fact.RowType` объявлен `NOT NULL`.
3. **`period_offset_fix`** — снимает +2000-летний офсет кластерного бэкенда (год ≥ 3000 → минус
   2000), обнуляет время (иначе не работает Date Range slicer по datetime) и деривирует
   `Period_Month = date(год, месяц, 1)`.
4. **`column_mapper`** — переименование в целевую схему; **колонки вне `column_map` отбрасываются**
   (`keep_extra` по умолчанию False). Добавили колонку в `raw_sql`, но не в `column_map` — она
   молча не доедет до таблицы.

Загрузка (`ai_olap\loaders\`): `mode: idempotent_period` → `FactLoader`
(`DELETE FROM Fact WHERE Period_Month = ?` + вставка). Если `--period` не задан, оркестратор
(`orchestrator\pipeline.py`, строки 64–70) **подменяет режим на `full_reload`**, то есть факт
перезаливается целиком через `TRUNCATE`. Диапазон параметров `?` ставит `auto_period_params`
(строки 39–51): с периодом — `[месяц+2000, следующий месяц+2000]`, без периода —
`[2001-01-01, 9999-01-01]`. Вставка — `bulk_insert` пачками по 1000 с `fast_executemany`;
список колонок читается из `sys.columns`, IDENTITY и `Loaded_At` пропускаются, отсутствующие
ключи → NULL.

## 6. 🔴 WHITELIST: порядок действий при изменении состава регистра

`Py_Olap\mapping\refresh_mapping.py`, строки 31–46 — 12 полных имён (регистр, документ-регистратор,
перечисление, 9 справочников). Скрипт вызывает `ПолучитьСтруктуруХраненияБазыДанных` **только по
этому массиву**; объект не в списке = его физической таблицы в JSON нет, и `resolve()` упадёт
(`introspect_planfakt_fields.py` печатает «НЕ НАЙДЕН в маппинге»). В текущем JSON 29 записей —
больше 12, потому что табличные части объектов попадают отдельными строками
(`_Reference589_VT39129` и т. п.).

Почему это опасно: `_FldNNN` — **платформенные номера, зависящие от порядка реквизитов**.
Добавили/удалили измерение или ресурс в `А_ПланФактПроизводство_Свод` — номера сдвигаются, а
`raw_sql` продолжает молча читать чужие колонки (типы совпадают → ошибки не будет, будут
неверные суммы).

Порядок (README, строки 45–48):

1. `.venv\Scripts\python.exe mapping\refresh_mapping.py` — перегенерировать `baserp_storage.json`;
2. `.venv\Scripts\python.exe scripts\introspect_planfakt_fields.py` — получить актуальные `_Fld`;
3. руками поправить `raw_sql` **и** `column_map` в `pipelines\fact_planfact_proizvodstvo.json`;
4. при новой колонке — `ALTER TABLE` в `ddl_olap_factory.sql` + `apply_ddl.py`, дописать колонку
   в список в `scripts\build_pbip.py` и пересобрать модель;
5. `tests\verify_olap_factory.py` — приёмка.

## 7. 🔴 FROZEN_ENUMS: чем грозит переименование значений

`Py_Olap\ai_olap\transformers\enum_resolver.py`, строки 38–43:

```python
"Перечисление.А_ТипыСтрокПланФактПроизводство": [
    "План",        # order 0
    "Факт",        # order 1
    "Виконання",   # order 2
    "ПланНаФакт",  # order 3
],
```

Список **позиционный**: значение подставляется по `_EnumOrder` (`FROZEN_ENUMS[meta][n]`), а не по
имени в 1С. Отсюда три следствия:

* переставили значения местами в конфигурации 1С → те же строки лягут на другие записи, факт
  «План» станет «Факт», ошибки не будет;
* добавили значение в 1С, не дописав в список → `TransformError: enum order N out of frozen list …
  Configuration changed?` — это единственный случай, когда ETL честно падает;
* **переименовали строку здесь** → регресс мер Power BI: эти литералы используются в DAX буквально
  (комментарий строк 32–36 и `build_pbip.py`, строки 14–15). Одновременно ломается связь
  `Fact.RowType → Dim_RowTypes.RowType`: это PK-строка, а не суррогат.

Тот же список продублирован в `scripts\apply_ddl.py`, строки 30–35, где он сеет `Dim_RowTypes`
вместе с группой периода:

```python
("План", 0, "Незалежний"), ("Факт", 1, "Період"),
("Виконання", 2, "Період"), ("ПланНаФакт", 3, "Незалежний"),
```

`RowType_Group` — это и есть **правило периода контура в данных**: «Період» читается за выбранный
период, «Незалежний» — целиком. Правки нужны в обоих файлах, иначе сев и резолвер разъедутся.
Сев `Dim_RowTypes` идемпотентен: `apply_ddl.py` пропускает таблицу, если там уже есть строки —
изменённый список **не** перезапишет старые значения, таблицу надо чистить руками.
`Calendar` ведёт себя иначе: он каждый прогон перезаливается из эталона (§3.1), поэтому любые
ручные правки календаря в витрине будут молча потеряны — править надо эталон в `OlapBASERP`.

## 8. Два представления

* **`vw_Fact_PlanFact`** (DDL, строки 195–207) — весь факт плюс вычисленная колонка:
  `COALESCE(i.Item_Name, ind.Individual_Name, NULLIF(f.Analytics_Text, N''), N'(не задано)')
  AS Analytics_Name`, где `Dim_Items` и `Dim_Individuals` присоединены по одному и тому же
  `f.Analytics_ID`. Составной реквизит «Аналитика» (номенклатура / физлицо / свободный текст)
  превращается в одну читаемую колонку — благодаря этому в модели Power BI нет неоднозначных
  связей и нет DAX-лукапов. Power BI читает именно представление, не таблицу.
* **`Dim_Departments_Tree`** (строки 163–188) — рекурсивный CTE: корень — `Parent_ID IS NULL OR
  Parent_ID = REPLICATE('0',32)` (пустая ссылка 1С после `varbinary_to_uuid` может прийти и как
  NULL, и как 32 нуля), развёртка ограничена `Hierarchy_Depth < 5`, отдаёт `Level1..Level5` и
  `Hierarchy_Path` через ` / `. Готовый drill «Производство → МД IRS 2026 → 15 м → №1» без
  отдельного пайплайна. Иерархия глубже 5 уровней будет обрезана.

## 9. Команды запуска

Рабочий каталог: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap_Factory\Py_Olap`

```bat
.venv\Scripts\python.exe scripts\apply_ddl.py                     :: DDL + копия Calendar из OlapBASERP + сев Dim_RowTypes
.venv\Scripts\python.exe mapping\refresh_mapping.py               :: перегенерация baserp_storage.json (COM)
.venv\Scripts\python.exe scripts\introspect_planfakt_fields.py    :: физические _InfoRg / _Fld

.venv\Scripts\python.exe main.py                                  :: весь конвейер: справочники + факт
.venv\Scripts\python.exe main.py --period 2026-07                 :: то же, факт только за месяц
.venv\Scripts\python.exe main.py --run-once dim_factory
.venv\Scripts\python.exe main.py --run-once fact_planfact_proizvodstvo
.venv\Scripts\python.exe main.py --run-once fact_planfact_proizvodstvo --period 2026-07
.venv\Scripts\python.exe main.py --validate                       :: схема пайплайнов
.venv\Scripts\python.exe main.py --refresh-mapping                :: сброс КЕША (JSON НЕ перегенерирует)
.venv\Scripts\python.exe main.py --scheduled                      :: APScheduler по cron из пайплайнов

.venv\Scripts\python.exe scripts\build_pbip.py                    :: пересборка модели Power BI
.venv\Scripts\python.exe tests\verify_olap_factory.py             :: приёмка: SQL == 1С
```

Порядок при доработке 1С (README, строка 43): перепровести документы → `main.py` → Refresh в Power BI.
`--refresh-mapping` и `mapping\refresh_mapping.py` — **разные вещи**: первый только чистит кеш
резолвера в процессе (`main.py`, строки 27–29).

Расписания в пайплайнах (используются только режимом `--scheduled`): `dim_factory` — `0 1 * * *`,
`fact_planfact_proizvodstvo` — `0 2 * * *`.

## 10. Что проверяет приёмка

`Py_Olap\tests\verify_olap_factory.py` — 7 блоков: 6.1 последние `ETL_Runs` = Success;
6.2 число строк Fact == число записей регистра (COM-запрос к 1С); 6.3 суммы всех 8 ресурсов
SQL == 1С с допуском `< 0.005`; 6.4 разрез по подразделениям; 6.5 FK-сироты по 8 связям
(`Department_ID`, `ExecutorDept_ID`, `Organization_ID`, `Stage_ID`, `Work_ID`, `CommonName_ID`,
`Unit_ID`, `RowType`) — `Analytics_ID`, `Document_ID` и `Recorder_ID` намеренно **не** проверяются
как составные/служебные; 6.6 `RowType` без пустых; 6.7 `Period` в человеческом диапазоне (год
≤ 2100 — контроль снятого офсета +2000) и полностью покрыт `Calendar`.

Эталон 2026-08-01 (README, строки 81–83): План 16 214,750 год / 20 424 422,20 грн,
Факт 15 092,000 год / 14 331 720,09 грн, Виконання 3 930,958 год, ETC 9 036 958,39 грн.

## 11. Мелочи, о которые спотыкаются

* `apply_ddl.py` режет `.sql` по `"\nGO"` и выполняет батчами в autocommit. Разделитель `GO`
  обязан стоять с начала строки, иначе батч склеится и упадёт.
* Кириллица в `ddl_olap_factory.sql` не используется намеренно (комментарий строки 4): sqlcmd
  портит кодировку, поэтому кириллические значения (`Dim_RowTypes`) вставляются параметрами
  pyodbc из Python. По той же причине там нет и `CREATE TABLE Calendar`: у эталона колонка
  называется `Квартал` — таблица целиком создаётся из Python по структуре эталона (§3.1).
* ETL ходит в SQL по DSN из `.env` (`SERVER=localhost`), а сгенерированная модель Power BI —
  на `Sql.Database("SQLSERVER", "OlapFactory", …)` (`build_pbip.py`, строки 26 и 51). Это один
  хост под разными именами; при переносе править оба места.
* Каждый прогон пишет строку в `ETL_Runs` (`orchestrator\runner.py`): при исключении —
  `Status='Failed'` и полный traceback в `Error`. Первое место, куда смотреть при разборе.
