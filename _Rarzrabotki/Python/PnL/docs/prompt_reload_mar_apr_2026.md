# ПРОМТ: Перезалить Март 2026 + Апрель 2026 в 1С из обновлённого Excel

Ты работаешь в `C:\Configuration_downloads\BASERP25` (BAS ERP 2.5 INDUSTRIALBUD, 1С 8.3.20+). Пайплайн PnL —
`_Rarzrabotki/Python/PnL/`. Подключение к ERP — `config.py` → `CONN_ERP` (`Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"`).
Интерпретатор: `..\venv\Scripts\python.exe` (fallback `C:\Python313\python.exe`).

## Цель

**Перезалить** (refresh) плановые отчёты PnL за **Март 2026** и **Апрель 2026** из обновлённых финансистом Excel:
- `C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Березень_26\!PL по компании Березень 2026.xlsx` (period 2026-03-31)
- `C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Апрель_26\!PL по компании Квітень 2026.xlsx` (period 2026-04-30)

Организация — **ТОВ "ІНДАСТРІАЛБУД"**. Один лист Excel = один `Документ.А_ОтчетPL` (на подразделение). Пайплайн
**идемпотентный**: `12_upload_documents.py` ищет существующий документ по `Дата + ПодразделениеСтрока` в пределах
месяца и **обновляет** его (не плодит дубли). Поэтому перезаливка безопасна.

## Текущее состояние БД (проверено 2026-06-09 через MCP `execute_query`)

| Период | `А_ОтчетPL` (проведено) | `А_РасшифровкаЛистов` |
|---|---|---|
| 2026-01 | 26 | №2 |
| 2026-02 | 28 | №3 |
| **2026-03** | **30** | **№4** |
| **2026-04** | **32** | **№5** |

`config.EXCEL_FILES` **УЖЕ содержит** записи Березень/Квітень 2026 (последние 2 записи) — **менять config НЕ нужно**.
Override `МД МХП ОРІЛЬ` (апрель) уже добавлен в `MANUAL_SHEET_TO_STRUCT_OVERRIDES`.

> ⚠️ **Готча MCP:** не фильтруй `Дата >= '2026-03-01'` строковым параметром через MCP `execute_query` — строка-дата
> НЕ биндится в виртуальную/обычную таблицу (вернёт пусто = ложь). Используй серверные `ГОД(Дата)` /
> `НАЧАЛОПЕРИОДА(Дата, МЕСЯЦ)` БЕЗ параметра-даты, либо Python COM с `datetime`.

## Архитектура pipeline (известна, заново не открывать)

| Скрипт | `--month` | Назначение |
|---|---|---|
| `01_extract_excel_to_json` | — | Excel → `01_raw_sheets.json` (читает ВСЕ `config.EXCEL_FILES`; различает месяцы по `month_header`, не по path) |
| `02_extract_unique_articles` | — | дедуп статей |
| `03_extract_dds_from_erp` | — | `СтатьиДвиженияДенежныхСредств` → 04 |
| `04_match_articles` | — | fuzzy PL↔ДДС + `MANUAL_DDS_OVERRIDES` → 05 |
| `04b_match_statya_dohodov` | — | `СтатьиДоходов` + `MANUAL_STATYA_DOHODOV_OVERRIDES` → 10 |
| `05_build_groups` | — | группы PL → 03 |
| `06_extract_struct_from_erp` | — | `СтруктураПредприятия` → 06 |
| `07_match_sheets_to_struct` | — | fuzzy листы↔подразделения + `MANUAL_SHEET_TO_STRUCT_OVERRIDES` → 07 |
| `10_upload_groups` | — | в 1С `А_ГруппаСтатей_PL` (идемпотентно) |
| `11_upload_articles` | post-filter | в 1С `А_Статьи_PL` (locked-aware: `А_РучнаяКорректировка=Истина` → пропуск) |
| `13_fill_rasshifrovka_listov` | **YES** | find-or-create `А_РасшифровкаЛистов` за месяц + перезаполнить ТЧ Расшифровка |
| `08_prepare_documents` | **YES** | собрать документы месяца → `08_documents_to_import.json` |
| `12_upload_documents` | **YES** | залить/обновить `А_ОтчетPL` (find-or-create по Дата+ПодразделениеСтрока, проводит) |

**Критический порядок:** `10`+`11` (заливают каталоги и **backfill UUID** в JSON) идут **ДО** `08` (который читает
UUID из JSON). Иначе в документах поле `Группа` пустое (см. `reload_instructions.md`).

## НЕ запускать (жёстко)

- ❌ `09_delete_existing_docs.py` — снесёт ВСЕ `А_ОтчетPL` (все года).
- ❌ `13_fill_rasshifrovka_listov.py` **без `--month`** — затрёт ТЧ Расшифровка ВСЕХ документов №1–№5 (ручные правки финансиста).
- ❌ `17_upload_pl_sort.py` — `Сорт` зафиксирован по эталону Лютий 2026, новый прогон не нужен.
- ❌ Менять `month_header`/удалять записи 2024–2025 в `config.EXCEL_FILES`.
- ❌ Удалять `А_РасшифровкаЛистов №1` (декабрь, эталон с правками финансиста).

## Алгоритм

```powershell
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
$py = "..\venv\Scripts\python.exe"   # fallback: C:\Python313\python.exe
```

### Шаг 0 — Pre-checks
- Excel-файлы март/апрель на месте (пути выше). Конфиг их уже содержит — **config не трогаем**.
- ERP по COM: `& $py scripts\_verify_excel_vs_pl.py` (без ошибок).
- Снять эталон ДО: `А_ОтчетPL` по месяцам и `А_РасшифровкаЛистов` (запросы из «Верификации» ниже) — ожидаем 30/32 + №4/№5.
- ⚠️ **Ручные правки:** `13 --month` **перезапишет** ТЧ Расшифровка документов №4 (март) / №5 (апрель). Если финансист
  правил привязку листов→подразделений в этих документах вручную — правки затрутся (бэкап в `data/json/13_backup_*.json`).
  Если такие правки были и важны — перед прогоном перенести их в `MANUAL_SHEET_TO_STRUCT_OVERRIDES` (config), чтобы `07` их учёл.

### Шаг 1 — Каноники `01–07` (БЕЗ `--month`) — рефреш JSON из обновлённого Excel
```powershell
& $py scripts\01_extract_excel_to_json.py
& $py scripts\02_extract_unique_articles.py
& $py scripts\03_extract_dds_from_erp.py
& $py scripts\04_match_articles.py
& $py scripts\04b_match_statya_dohodov.py
& $py scripts\05_build_groups.py
& $py scripts\06_extract_struct_from_erp.py
& $py scripts\07_match_sheets_to_struct.py
```
**Следить (обновлённый Excel мог принести новое):**
- **Новые статьи** в `02_unique_articles.json` → если это коммент-арифметика финансиста (числа/опечатки) → в
  `EXTRA_COMMENT_NAMES`; если вариант имени существующей статьи → в `ARTICLE_NAME_ALIASES` (см. `pnl_excel_fake_articles_filter.md`).
- **Новые листы** (МД/МК-филиалы, появившиеся в марте/апреле) с fuzzy-score < 85 → добавить в `MANUAL_SHEET_TO_STRUCT_OVERRIDES`
  (struct_uuid/direction взять из `06_struct_predpr.json` или `execute_query` по `СтруктураПредприятия`). Проверить `data/reports/unmapped_sheets.txt`.
- При правках config/overrides — перезапустить `02`/`07` соответственно.

### Шаг 2 — Pre-flight: защита статей
```powershell
& $py scripts\_mass_set_manual_flag.py --dry-run
& $py scripts\_mass_set_manual_flag.py
```
Если есть НЕ-locked статьи (`НЕ А_РучнаяКорректировка`) — `11` упадёт `Не удалось записать "X"`. Флаг ставится через
`ОбменДанными.Загрузка=Истина` (обход гардов `ПередЗаписью`).

### Шаг 3 — Каталоги `10`, `11`
```powershell
& $py scripts\10_upload_groups.py
& $py scripts\11_upload_articles.py
```
Ожидается `locked≈60+`, `created=` числу действительно новых статей (новые пишутся с `А_РучнаяКорректировка=Ложь` —
финансист потом ставит галку в 1С).

### Шаг 4 — Цикл по двум месяцам: Расшифровка → cleanup → документы
```powershell
foreach ($m in "2026-03","2026-04") {
    Write-Host "=== $m : расшифровка ==="
    & $py scripts\13_fill_rasshifrovka_listov.py --month $m --dry-run
    & $py scripts\13_fill_rasshifrovka_listov.py --month $m
}
# Очистить служебные строки (расчет 2026 / Метрики / PL_Свод) во всех документах — идемпотентно
& $py scripts\_cleanup_rasshifrovka_garbage.py
foreach ($m in "2026-03","2026-04") {
    Write-Host "=== $m : документы ==="
    & $py scripts\08_prepare_documents.py --month $m
    & $py scripts\12_upload_documents.py --month $m --limit 1   # тест на 1 документе
    & $py scripts\12_upload_documents.py --month $m             # полный прогон
}
```
- `13 --month` находит существующий №4/№5 по period-window (`Дата МЕЖДУ &НачалоМесяца И &КонецМесяца`) и **перезаполняет** ТЧ из обновлённого Excel.
- `12 --month` обновляет существующие 30/32 документа (find-or-create) — **дубли не плодятся**.

## Acceptance / Верификация (серверные даты, без MCP-string-параметра)

```sql
ВЫБРАТЬ НАЧАЛОПЕРИОДА(Дата, МЕСЯЦ) КАК Месяц, КОЛИЧЕСТВО(*) КАК Док,
    СУММА(ВЫБОР КОГДА Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Проведено
ИЗ Документ.А_ОтчетPL ГДЕ НЕ ПометкаУдаления
СГРУППИРОВАТЬ ПО НАЧАЛОПЕРИОДА(Дата, МЕСЯЦ) УПОРЯДОЧИТЬ ПО Месяц
```
Ожидается:
- **2026-03 и 2026-04 присутствуют, Док == Проведено** (если листы Excel не менялись — 30 и 32; если финансист
  добавил/убрал листы — счётчик соответственно изменится — это нормально, сверить с числом непустых листов месяца).
- **Прочие месяцы (2024-01..2026-02) НЕ изменились** (2024=396, 2025=331, 2026-01=26, 2026-02=28).

```sql
ВЫБРАТЬ ГОД(Дата) КАК Год, КОЛИЧЕСТВО(*) КАК Док ИЗ Документ.А_РасшифровкаЛистов
ГДЕ НЕ ПометкаУдаления СГРУППИРОВАТЬ ПО ГОД(Дата)
```
Ожидается 2026 = **4** (№2–№5) — НЕ задвоились (13 обновил существующие №4/№5, не создал новые).

Контроль данных (опц.): `& $py scripts\_verify_excel_vs_pl.py` и `_verify_pl_vs_adr.py` (отчёт `data/reports/verification_pl_vs_adr.csv`).

## Риски (из `pnl_monthly_import_pattern.md`, `prompt_load_2025_full_year.md`)
- **Один Excel-файл = много month_header'ов** не про этот случай (март/апрель — отдельные файлы, только колонки своего месяца) — пересечения нет.
- **Перезапуск `13` без `--month`** = катастрофа (затрёт №1–№5). Всегда `--month` + `--dry-run` сначала.
- **Новый филиал апреля** (`МД МХП ОРІЛЬ`) override уже есть; если в обновлённом Excel ещё новые листы — добавить override до `07`.
- **`12` идемпотентен**, но если в обновлённом Excel УБРАЛИ лист-подразделение — старый документ останется (find-or-create не удаляет). Сверить список подразделений до/после; лишние — пометить на удаление вручную в 1С.

## Follow-up (вне scope этого промта)
- Регенерация knowledge для NotebookLM (после загрузки): пайплайн `knowledge_PL/Python/PL/` — расширить хардкод-периоды
  `["2025-12","2026-01","2026-02"]` в `_compute_pl_aggregates.py` и `_render_faq.py` на `2026-03`/`2026-04`, прогнать
  `15_export_to_knowledge_pl.py --period 2026-03` / `--period 2026-04` + `_compute_pl_aggregates` + `_render_faq` → залить в NotebookLM. **Отдельная задача.**

## Источники
- `_Rarzrabotki/Python/PnL/README.md` (пайплайн, маппинг колонок, идемпотентность, locked-флаг).
- `_Rarzrabotki/Python/PnL/docs/reload_instructions.md` (порядок 10/11 ДО 08).
- `_Rarzrabotki/Python/PnL/docs/prompt_load_2025_full_year.md` (эталон `--month`-перезаливки, acceptance).
- `_Rarzrabotki/Python/PnL/config.py` (`EXCEL_FILES`, `MANUAL_SHEET_TO_STRUCT_OVERRIDES`, `ARTICLE_NAME_ALIASES`, `EXTRA_COMMENT_NAMES`).
- Память: `pnl_monthly_import_pattern.md` (правило `--month`, что НЕ запускать), `pnl_excel_fake_articles_filter.md` (фильтр коммент-статей).
