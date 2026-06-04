# ПРОМТ: Загрузить весь 2024 год (Январь–Декабрь) из декабрьского Excel 2024

Ты работаешь в `C:\Configuration_downloads\BASERP25` (BAS ERP 2.5 INDUSTRIALBUD на 1С:Предприятие 8.3.20+). Пайплайн PnL — в `_Rarzrabotki/Python/PnL/`. Подключение к ERP — `config.py` → `CONN_ERP`. Внешний интерпретатор: `..\venv\Scripts\python.exe`.

## Цель

Загрузить в 1С плановые отчёты PnL за **январь–декабрь 2024** (12 месяцев) из **одного декабрьского Excel-файла 2024**:

```
C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx
```

— этот файл содержит **колонки за весь 2024 год** (col=14 → Январь 2024, …, col=69 → Декабрь 2024), плюс годовые итоги 2022, 2023.

В БД сейчас **отсутствует весь 2024 год**, нужно создать 12 новых документов `А_РасшифровкаЛистов` (№17..№28 по нумерации) и ~280–340 проведённых `А_ОтчетPL`.

## Контекст — что уже есть и что подтверждено

### Состояние БД на 2026-05-21 (verified MCP)

`Документ.А_РасшифровкаЛистов` (16 записей):

| № | Дата | Период |
|---|------|--------|
| 000000006 | 31.01.2025 10:00 | Январь 2025 |
| 000000007 | 28.02.2025 10:00 | Февраль 2025 |
| 000000008 | 31.03.2025 09:00 | Март 2025 |
| ... | ... | ... |
| 000000016 | 30.11.2025 10:00 | Ноябрь 2025 |
| 000000001 | 31.12.2025 12:00 | Декабрь 2025 |
| 000000002 | 31.01.2026 12:00 | Январь 2026 |
| 000000003 | 28.02.2026 12:00 | Лютий 2026 |
| 000000004 | 31.03.2026 09:00 | Березень 2026 |
| 000000005 | 30.04.2026 09:00 | Квітень 2026 |

`Документ.А_ОтчетPL` — **445 проведённых** за 16 месяцев (январь 2025 – апрель 2026), 100% проведены.

### Структура Excel 2024 (подтверждено через openpyxl)

В листе `PL_ЦО` строка 2:

```
col=3  '2022'                 (годовой итог — НЕ загружаем)
col=8  '2023'                 (годовой итог — НЕ загружаем)
col=14 datetime(2024,1,1)     ← Январь 2024
col=19 datetime(2024,2,1)     ← Февраль 2024
col=24 datetime(2024,3,1)     ← Март 2024
col=29 datetime(2024,4,1)     ← Апрель 2024
col=34 datetime(2024,5,1)     ← Май 2024
col=39 datetime(2024,6,1)     ← Июнь 2024
col=44 datetime(2024,7,1)     ← Июль 2024
col=49 datetime(2024,8,1)     ← Август 2024
col=54 datetime(2024,9,1)     ← Сентябрь 2024
col=59 datetime(2024,10,1)    ← Октябрь 2024
col=64 datetime(2024,11,1)    ← Ноябрь 2024
col=69 datetime(2024,12,1)    ← Декабрь 2024
```

**36 visible листов** (vs 35 у 2025), **97 hidden** (отфильтруются). Из 36 видимых:
- 2 утилитных (`Метрики`, `расчет 2024`) — отсеются через `find_month_columns is None`;
- 1 сводный (`PL_Свод`) — также без datetime-маркера, отсеется;
- 1 PL_-сводный без direction (`PL_Староконстантинов`) — нужен override;
- 16 PnL-листов **новых** для каталога 07_mapping (в 2024 году работали филиалы, закрытые к 2025);
- 16 уже известны (Крушинка, Ретал.ГПУ, Глобино-2, PL_ЦО, PL Строительство Свод и т.д.).

### Архитектура pipeline — уже исправлена (2026-05-21)

`13_fill_rasshifrovka_listov.py` корректно работает с **множественными month_header'ами на один Excel**:
- `find_doc_by_file(conn, excel_path, period_str=ef["period"])` — приоритет поиск по period window.
- `create_doc_for_period(conn, excel_path, period_str=ef["period"])` — дата создаваемого документа берётся из явного period_str.
- Basename-fallback используется ТОЛЬКО когда `period_str=None` (legacy).
- Помеха `next((e for e in EXCEL_FILES if e["path"] == p))` устранена — main() передаёт `period_str` явно.

Скрипт `_cleanup_rasshifrovka_garbage.py` уже существует в `scripts/` (идемпотентный cleanup ТЧ от `расчет YYYY`, `Метрики`, `PL_Свод` + перепроведение).

Каталог `А_Статьи_PL` содержит **68 защищённых** (А_РучнаяКорректировка=Истина) статей, плюс несколько SKIP-NEW договоров подряда.

## Алгоритм

### Шаг 1 — Расширить `config.EXCEL_FILES` на 12 записей

Добавить в `_Rarzrabotki/Python/PnL/config.py` блок **в начало** `EXCEL_FILES` (перед записями 2025):

```python
EXCEL_FILES = [
    # === 2024 ГОД (один файл, 12 month_header'ов; col=14..69 step=5) ===
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-01-31", "month_header": "2024-01-01", "label": "Январь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-02-29", "month_header": "2024-02-01", "label": "Февраль 2024"},   # 2024 — високосный
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-03-31", "month_header": "2024-03-01", "label": "Март 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-04-30", "month_header": "2024-04-01", "label": "Апрель 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-05-31", "month_header": "2024-05-01", "label": "Май 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-06-30", "month_header": "2024-06-01", "label": "Июнь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-07-31", "month_header": "2024-07-01", "label": "Июль 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-08-31", "month_header": "2024-08-01", "label": "Август 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-09-30", "month_header": "2024-09-01", "label": "Сентябрь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-10-31", "month_header": "2024-10-01", "label": "Октябрь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-11-30", "month_header": "2024-11-01", "label": "Ноябрь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-12-31", "month_header": "2024-12-01", "label": "Декабрь 2024"},
    # === 2025 (11+1 запис, без изменений) ===
    # ...
]
```

**Високосный год**: февраль 2024 — 29 дней, period=`2024-02-29`. Не используй `2024-02-28`.

### Шаг 2 — Добавить overrides для 16 новых листов

В `MANUAL_SHEET_TO_STRUCT_OVERRIDES` (в config.py) добавить (все UUID **подтверждены через MCP**):

```python
# === 2024 ГОД — филиалы, закрытые/изменённые к 2025 ===
"PL_Староконстантинов": {
    "struct_uuid": "9db438f5-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Староконстантинов",
    "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Строительство",
    "include_children": True,
    "reason": "Сводный лист Староконстантинов (PL_-префикс) = подразделение Строительства",
},
"Молоко Вітчизни": {
    "struct_uuid": "9d2c90df-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Молоко Вітчизни",
    "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Строительство",
    "include_children": False,
    "reason": "Подразделение 2024 года, Строительство",
},
"КМД Путровка": {
    "struct_uuid": "9d2c911d-1ae7-11f0-80dc-00155d235309",
    "struct_name": "КМД Путровка",
    "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Закрытые обьекты",
    "include_children": False,
    "reason": "Объект закрыт к 2025, направление Закрытые объекты",
},
"Турбина-3": {
    "struct_uuid": "9d2c90b2-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Турбина 3",
    "direction_uuid": "9d2c84a1-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Девелопмент",
    "include_children": False,
    "reason": "В Excel имя 'Турбина-3', в 1С — 'Турбина 3' (без дефиса)",
},
"Нежин 3": {
    "struct_uuid": "9d2c8fc5-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Нежин 3 ",   # ВАЖНО: с пробелом в конце!
    "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Закрытые обьекты",
    "include_children": False,
    "reason": "Объект закрыт; в 1С имя с trailing space",
},
"Глобино": {
    "struct_uuid": "9d2c90ca-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Глобино",
    "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Закрытые обьекты",
    "include_children": False,
    "reason": "Объект 'Глобино' закрыт к 2025 (отдельный от 'Глобино-2' и 'МК Глобино')",
},
"Приднепровский_металл": {
    "struct_uuid": "9dc455b5-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Приднепровский-металл",
    "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Закрытые обьекты",
    "include_children": False,
    "reason": "В Excel underscore, в 1С — дефис",
},
"Солома_Элеватор": {
    "struct_uuid": "9d2c8c54-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Солома Элеватор",
    "direction_uuid": None,  # направление не задано в 1С
    "direction_name": None,
    "include_children": False,
    "reason": "В Excel underscore, в 1С — пробел; направление не задано",
},
"МД Первомайск ПА": {
    "struct_uuid": "9d2c9319-1ae7-11f0-80dc-00155d235309",
    "struct_name": "МД Первомайск ПА",
    "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Закрытые обьекты",
    "include_children": False,
    "reason": "Точное совпадение, объект закрыт к 2025",
},
"НГУ": {
    "struct_uuid": "9d2c91e6-1ae7-11f0-80dc-00155d235309",
    "struct_name": "НГУ",
    "direction_uuid": None,
    "direction_name": None,
    "include_children": False,
    "reason": "Точное совпадение, направление не задано",
},
"МД ООН": {
    "struct_uuid": "9d2c8fb1-1ae7-11f0-80dc-00155d235309",
    "struct_name": "МД ООН",
    "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Производство",
    "include_children": False,
    "reason": "Отдельный от 'МД ООН 2025' (тот появился в 2025)",
},
"МД ВООЗ": {
    "struct_uuid": "9d2c90c9-1ae7-11f0-80dc-00155d235309",
    "struct_name": "МД ВООЗ",
    "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Производство",
    "include_children": False,
    "reason": "Отдельный от 'МД ВООЗ  2025' (тот появился в 2025)",
},
"ЕМС": {
    "struct_uuid": "9d2c92f3-1ae7-11f0-80dc-00155d235309",
    "struct_name": "ЕМС малий",
    "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
    "direction_name": "Производство",
    "include_children": False,
    "reason": "В Excel короткое 'ЕМС', в 1С — 'ЕМС малий'",
},
"Форд Тягач": {
    "struct_uuid": "9dc44b0c-1ae7-11f0-80dc-00155d235309",
    "struct_name": "Тягач Форд  1842Т",
    "direction_uuid": None,
    "direction_name": None,
    "include_children": False,
    "reason": "Машина: в 1С 'Тягач Форд  1842Т' с двумя пробелами",
},
"MAN Blue KA1783HI": {
    "struct_uuid": "9de72248-1ae7-11f0-80dc-00155d235309",
    "struct_name": "MAN blue № 1 KA1783HI",
    "direction_uuid": None,
    "direction_name": None,
    "include_children": False,
    "reason": "В Excel 'MAN Blue KA1783HI', в 1С — 'MAN blue № 1 KA1783HI'",
},
"MAN green KA2790IE": {
    "struct_uuid": "9de7226d-1ae7-11f0-80dc-00155d235309",
    "struct_name": "MAN green KA2790IE",
    "direction_uuid": None,
    "direction_name": None,
    "include_children": False,
    "reason": "Точное совпадение",
},
```

### Шаг 3 — Прогнать каноники (01–07) БЕЗ `--month`

```powershell
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
$py = "..\venv\Scripts\python.exe"

& $py scripts\01_extract_excel_to_json.py
& $py scripts\02_extract_unique_articles.py
& $py scripts\03_extract_dds_from_erp.py
& $py scripts\04_match_articles.py
& $py scripts\04b_match_statya_dohodov.py
& $py scripts\05_build_groups.py
& $py scripts\06_extract_struct_from_erp.py
& $py scripts\07_match_sheets_to_struct.py
```

**Ожидается:** `01_raw_sheets.json` → 28 периодов (16 существующих + 12 новых 2024), `02_unique_articles.json` → ~66–72 статьи (2024 может добавить несколько уникальных). Проверить `02_unique_articles.json` на фейковые статьи-комментарии:

```bash
python -c "import json; d=json.load(open(r'data/json/02_unique_articles.json', encoding='utf-8')); only=[a for a in d['articles'] if all('2024' in p for p in a['periods'])]; print('Только 2024:', len(only)); [print(' ', a['name']) for a in only]"
```

Если появились комментарии с числами + единицей валюты (\dгрн, Nтыс, тис, млн) — фильтр `_COMMENT_VALUE_RE` в excel_parser сработает. Если **точные имена** без чисел типа `Доход 3,7+0,7-рахсод 8,8-4,5` — добавить в `config.EXTRA_COMMENT_NAMES` и перепрогнать `01`+`02`.

### Шаг 4 — Pre-flight: НЕ-locked статьи

```powershell
& $py scripts\_mass_set_manual_flag.py --dry-run
& $py scripts\_mass_set_manual_flag.py   # только если dry-run показал >0 unlocked
```

В текущем состоянии БД (2026-05-21) **все 68 статей locked** → этот шаг будет no-op.

### Шаг 5 — Каталоги (10, 11)

```powershell
& $py scripts\10_upload_groups.py
& $py scripts\11_upload_articles.py
```

Если в 2024 году появились **новые статьи без ДДС** → SKIP-NEW. Если новые статьи с типом Доход — будут созданы (создание через `mgr.СоздатьЭлемент()`, требуется ручная проверка финансистом и установка А_РучнаяКорректировка). После создания **обязательно** запустить `_mass_set_manual_flag.py` чтобы при следующем прогоне не пытаться обновить.

### Шаг 6 — А_РасшифровкаЛистов: 12 новых документов

```powershell
foreach ($m in "2024-01","2024-02","2024-03","2024-04","2024-05","2024-06","2024-07","2024-08","2024-09","2024-10","2024-11","2024-12") {
    Write-Host "=== $m ==="
    & $py scripts\13_fill_rasshifrovka_listov.py --month $m --dry-run
    & $py scripts\13_fill_rasshifrovka_listov.py --month $m
}
```

После каждого запуска `13_fill_rasshifrovka_listov.py` создаст документ `А_РасшифровкаЛистов № N+1` с правильной датой (последний день месяца, 12:00). **Документы №1–№16 НЕ затронутся** — period window их игнорирует.

Ожидается итог: 12 новых документов с датами 2024-01-31, 2024-02-29, …, 2024-12-31. Все с ТЧ Расшифровка по ~33–36 строк, ~30 с Подразделением, ~5 свод-строк.

### Шаг 7 — Cleanup ТЧ от служебных листов

```powershell
& $py scripts\_cleanup_rasshifrovka_garbage.py
```

Идемпотентен: удалит из ТЧ всех документов (включая 12 новых) строки с ИмяЛиста IN `{"расчет 2024", "расчет 2025", "расчет 2026", "Метрики", "PL_Свод"}` и перепроведёт. Существующие 16 документов не пострадают (там уже очищено).

### Шаг 8 — А_ОтчетPL за 12 месяцев

```powershell
foreach ($m in "2024-01","2024-02","2024-03","2024-04","2024-05","2024-06","2024-07","2024-08","2024-09","2024-10","2024-11","2024-12") {
    Write-Host "=== $m ==="
    & $py scripts\08_prepare_documents.py --month $m
    & $py scripts\12_upload_documents.py --month $m --limit 1
    & $py scripts\12_upload_documents.py --month $m
}
```

Каждая итерация: `08` соберёт ~25–32 документа в JSON, `12 --limit 1` создаст 1 тестовый, `12` (без limit) — UPD-ит первый + NEW остальные. Скрипт идемпотентен по (Дата, ПодразделениеСтрока).

## Acceptance criteria

### Минимум

1. `EXCEL_FILES` содержит 28 записей: 12 (2024) + 12 (2025) + 4 (2026). Записи 2025 и 2026 не изменились.
2. `MANUAL_SHEET_TO_STRUCT_OVERRIDES` содержит на 16 ключей больше (новые 2024-листы).
3. `01_raw_sheets.json` → 28 периодов.
4. `02_unique_articles.json` → 60–75 статей (изменение ± от текущих 66).
5. `11_upload_articles.py`: `locked=63` (или больше), `created=0..3` (новые 2024-статьи без ДДС → SKIP-NEW), `failed=0`.
6. 12 новых документов `А_РасшифровкаЛистов` с правильными датами (контроль через `execute_query`).
7. `_cleanup_rasshifrovka_garbage.py` после прогона: для №17..№28 удалено по 3 строки (`расчет 2024`, `Метрики`, `PL_Свод`).
8. `А_ОтчетPL` за каждый 2024-месяц: 25–32 документа, все `Проведен = True`, 0 failed.
9. Существующие 2025 и 2026 А_ОтчетPL не задеты (445 документов до — должно остаться 445 + ~300 = ~745).

### Финальная верификация

```sql
ВЫБРАТЬ
    НАЧАЛОПЕРИОДА(Дата, МЕСЯЦ) КАК Месяц,
    КОЛИЧЕСТВО(*) КАК Документов,
    СУММА(ВЫБОР КОГДА Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Проведено
ИЗ Документ.А_ОтчетPL
ГДЕ НЕ ПометкаУдаления
СГРУППИРОВАТЬ ПО НАЧАЛОПЕРИОДА(Дата, МЕСЯЦ)
УПОРЯДОЧИТЬ ПО Месяц
```

Ожидается 28 строк (январь 2024 – апрель 2026), все Документов == Проведено.

```sql
ВЫБРАТЬ Номер, Дата ИЗ Документ.А_РасшифровкаЛистов
ГДЕ НЕ ПометкаУдаления УПОРЯДОЧИТЬ ПО Дата
```

Ожидается 28 строк: первый — `№N 31.01.2024`, последний — `№5 30.04.2026`.

## Прецеденты и риски

### Риск 1: Високосный февраль 2024

Феraль 2024 имеет 29 дней. В `EXCEL_FILES` для февраля укажи `period="2024-02-29"`, иначе `create_doc_for_period` создаст документ с `datetime(2024, 2, 28, 12, 0, 0)` — ошибки не будет, но дата некорректная. Лучше точная.

### Риск 2: Filтр `_COMMENT_VALUE_RE` пропустил коммент

Если в 2024 финансист добавил комментарий с арифметикой типа `Доход 3,7+0,7-рахсод 8,8-4,5` (см. прецедент 2025) — добавить в `config.EXTRA_COMMENT_NAMES` и перепрогнать `01`+`02`.

### Риск 3: Закрытые направления (`Закрытые обьекты`)

Несколько 2024-филиалов привязаны к направлению `Закрытые обьекты` (UUID `9d021b71-1ae7-11f0-80dc-00155d235309`). Это **легальное** направление 1С (с опечаткой 'обьекты'), не помеха для записи документа.

### Риск 4: `13` без `--month` — катастрофа

Точно как для 2025: запуск `13` без `--month` пробежит по всем 28 EXCEL_FILES записям, затрёт ТЧ Расшифровка всех 16 существующих документов (включая ручные правки финансиста в №1!) и создаст 12 новых. **Каждый запуск 13 — обязательно с `--month YYYY-MM`.** Перед боевым — `--dry-run`.

### Риск 5: Подразделение `Глобино` ≠ `Глобино-2`

В Excel 2024 есть лист `Глобино` (без суффикса). В 1С есть как `Глобино` (Закрытые объекты), `Глобино-2` (Строительство) и `МК Глобино` (Производство). Override явно указывает `9d2c90ca-1ae7-11f0-80dc-00155d235309` = старый закрытый объект. **НЕ путать** с активным `Глобино-2`.

### Риск 6: Имя в 1С с trailing space

`Нежин 3 ` (с пробелом в конце) — реальное имя в 1С. Не trim. Поиск через `НайтиПоНаименованию` обычно работает корректно, но для отображения в 1С это будет `Нежин 3 `.

## Что НЕ делать

- **НЕ удалять** существующие записи 2025/2026 в `EXCEL_FILES`.
- **НЕ запускать** `13` без `--month`.
- **НЕ запускать** `09_delete_existing_docs.py`.
- **НЕ запускать** `17_upload_pl_sort.py`.
- **НЕ менять** уже существующие записи в `MANUAL_SHEET_TO_STRUCT_OVERRIDES` (только добавлять новые ключи).
- **НЕ trim** имена подразделений с пробелами (например `Нежин 3 `).
- **НЕ объединять** `МД ООН` (2024) и `МД ООН 2025` — это разные подразделения с разными UUID.
- **НЕ объединять** `Глобино` (закрыт) и `Глобино-2` / `МК Глобино` (активные).

## Как запустить

```powershell
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
$py = "..\venv\Scripts\python.exe"

# 1. Обновить config.py (Шаги 1 и 2)
# (через Edit-инструмент в IDE)

# 2. Каноники
& $py scripts\01_extract_excel_to_json.py
& $py scripts\02_extract_unique_articles.py
& $py scripts\03_extract_dds_from_erp.py
& $py scripts\04_match_articles.py
& $py scripts\04b_match_statya_dohodov.py
& $py scripts\05_build_groups.py
& $py scripts\06_extract_struct_from_erp.py
& $py scripts\07_match_sheets_to_struct.py

# 3. Pre-flight
& $py scripts\_mass_set_manual_flag.py --dry-run
# & $py scripts\_mass_set_manual_flag.py   # только если dry-run показал >0

# 4. Каталоги
& $py scripts\10_upload_groups.py
& $py scripts\11_upload_articles.py

# 5. А_РасшифровкаЛистов (12 новых документов)
foreach ($m in "2024-01","2024-02","2024-03","2024-04","2024-05","2024-06","2024-07","2024-08","2024-09","2024-10","2024-11","2024-12") {
    & $py scripts\13_fill_rasshifrovka_listov.py --month $m --dry-run
    & $py scripts\13_fill_rasshifrovka_listov.py --month $m
}

# 6. Cleanup
& $py scripts\_cleanup_rasshifrovka_garbage.py

# 7. А_ОтчетPL (12 месячных пачек)
foreach ($m in "2024-01","2024-02","2024-03","2024-04","2024-05","2024-06","2024-07","2024-08","2024-09","2024-10","2024-11","2024-12") {
    & $py scripts\08_prepare_documents.py --month $m
    & $py scripts\12_upload_documents.py --month $m --limit 1
    & $py scripts\12_upload_documents.py --month $m
}

# 8. Финальная верификация — два execute_query из раздела Acceptance criteria
```

## Источники информации

- `_Rarzrabotki/Python/PnL/README.md` — пайплайн и маппинг колонок.
- `_Rarzrabotki/Python/PnL/docs/prompt_load_2025_full_year.md` — родной шаблон, прецедент годового импорта.
- `_Rarzrabotki/Python/PnL/scripts/13_fill_rasshifrovka_listov.py` (после фикса 2026-05-21) — `find_doc_by_file(..., period_str)` и `create_doc_for_period(..., period_str)`.
- `_Rarzrabotki/Python/PnL/scripts/_cleanup_rasshifrovka_garbage.py` — готовый скрипт очистки.
- Memory `pnl_monthly_import_pattern.md` — порядок прогона, `--month` правило, фикс 2026-05-21.
- Memory `pnl_excel_fake_articles_filter.md` — фильтр комментариев финансиста.
- Прецедент **2026-05-21** (2025 год: январь–ноябрь): 11 документов А_РасшифровкаЛистов + 301 А_ОтчетPL созданы без потери существующих данных.
- Probe Excel 2024 (2026-05-21): подтверждены 12 month-колонок (col=14..69 step=5), 36 visible листов, 16 новых для 07_mapping.
- MCP-поиск подразделений 2024 (2026-05-21): найдены UUID 16 филиалов через `Справочник.СтруктураПредприятия ПОДОБНО %...%`.

---

**Перед началом:**
1. Сделать MCP-снимок текущего состояния: `А_РасшифровкаЛистов` (16) и `А_ОтчетPL` помесячно (445 за январь 2025 – апрель 2026).
2. Если расходится — стоп, уточнить.
3. При первой ошибке (особенно в шаге 6 или 8) — НЕ продолжать вслепую, поднять backup из `data/json/13_backup_*.json` или `09_import_log.json`.
