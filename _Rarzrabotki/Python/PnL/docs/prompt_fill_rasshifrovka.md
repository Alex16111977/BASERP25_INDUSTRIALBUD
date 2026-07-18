# ПРОМТ: Доработать заполнение документа `Документ.А_РасшифровкаЛистов` через Python-скрипт

Ты работаешь в `C:\Configuration_downloads\BASERP25` (конфигурация BAS ERP 2.5 INDUSTRIALBUD на 1C:Enterprise 8.3.20+). Подключение к ERP описано в `_Rarzrabotki/Python/PnL/config.py` → `CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="..."'`.

## Цель

Написать Python-скрипт, который заполнит ТЧ `Расшифровка` у 3-х существующих документов `Документ.А_РасшифровкаЛистов` **всеми видимыми листами** соответствующих Excel-файлов PnL (Декабрь 2025, Январь 2026, Лютий 2026). В колонку `Подразделение` подставить UUID подразделения **только** для тех листов, где известно **точно** (по `07_mapping_sheet_to_struct.json`). Остальные оставить пустыми — финансист дозаполнит вручную.

Назначение документа: **маппинг «лист Excel ↔ подразделение 1С»** через бизнес-процесс (финансист правит UI документа, не JSON).

## Контекст — что УЖЕ есть

### В базе 1С
Через `mcp__1c-workerp__execute_query` подтверждено (2026-04-18) — 3 существующих документа:

| № | Дата | `ИмяФайла` (как сейчас в базе) |
|---|------|-------------------------------|
| 000000001 | 31.12.2025 | `!PL по компании Декабрь 2025.xlsx` |
| 000000002 | 31.01.2026 | `!PL по компании Январь 2026.xlsx` |
| 000000003 | 28.02.2026 | `!PL по компании Лютий 2026.xlsx` |

**ВАЖНО:** в базе хранится **basename** (только имя файла, без пути). В `config.py` → `EXCEL_FILES[i]["path"]` — **полный путь**. Это требует двух разных значений для поиска (по basename) и для сохранения (полный путь). См. Шаг 3.

**Документ №000000001 не пустой:** финансист внёс 4 строки (Крушинка, Ретал.ГПУ, МХП Катеринополь, а также ручной маппинг `PL_ЦО_Строительство → Строительство` — этого маппинга НЕТ в `07_mapping.json`). Скрипт обязан **полностью перезаписать** ТЧ — последняя строка потеряется, финансист допишет после импорта. Перед очисткой ОБЯЗАТЕЛЬНО сделать **backup** текущей ТЧ (см. Шаг 2a).

Документы №2 и №3 — ТЧ пустая.

### В проекте `_Rarzrabotki/Python/PnL/`
- `config.py` → `EXCEL_FILES` (список из 3 dict'ов с `path`, `period`, `month_header`, `label`) и `CONN_ERP`.
- `scripts/utils/com_connect.py` → функции `connect_erp()` (V83.COMConnector), `uuid_str(conn, ref)`, `run_query(conn, text, params)`.
- `data/json/07_mapping_sheet_to_struct.json` → 14 записей маппинга вида:
  ```json
  {"sheet_name": "Крушинка", "struct_uuid": "9d2c8f1a-1ae7-11f0-80dc-00155d235309",
   "struct_name": "Крушинка", "match_type": "exact", "confidence": 100, ...}
  ```
  Возможные значения `match_type`: `exact` | `fuzzy` | `manual_override` | `none`.
- `scripts/12_upload_documents.py` → **образец** работы с COM для документов (функция `ref_by_uuid`, паттерн `СоздатьДокумент / ПолучитьОбъект`, `Записать()`).

## Точные метаданные `Документ.А_РасшифровкаЛистов`

Подтверждены через `mcp__1c-workerp__get_metadata_structure(metaType="Documents", name="А_РасшифровкаЛистов")`:

**Шапка:**
| Реквизит | Тип | Назначение |
|---|---|---|
| `Номер` | Строка | стандартный |
| `Дата` | Дата | стандартная |
| `ИмяФайла` | Строка | полный путь к Excel-файлу |
| `Комментарий` | Строка | произвольный |

**Табличная часть `Расшифровка`:**
| Реквизит | Тип | Что класть |
|---|---|---|
| `ИмяЛиста` | Строка | точное имя листа Excel (как есть, без обрезки) |
| `Подразделение` | `СправочникСсылка.СтруктураПредприятия` | UUID из 07_mapping или пусто |
| `ПодразделениеСтрока` | Строка(150+) | **копия** `ИмяЛиста` (привязка к Excel остаётся даже если Подразделение пусто) |
| `НаправлениеДеятельности` | `СправочникСсылка.НаправленияДеятельности` | UUID направления из 07_mapping.direction_uuid (подтянут из 06 или задан в override) |
| `ВключатьДочерние` | `Булево` | `True` для сводных листов (PL_ЦО, Техника, PL Строительство Свод и т.п. — из 07_mapping.include_children) |

**Добавлено 2026-04-18** (миграция НаправлениеДеятельности): строки с `ВключатьДочерние=True` должны иметь заполненное `НаправлениеДеятельности` (валидация в `ПередЗаписью` документа А_ОтчетPL отказывает записи). См. `docs/2026-04-18-napravlenie-deyatelnosti-plan.md`.

## Алгоритм скрипта

Создать `_Rarzrabotki/Python/PnL/scripts/13_fill_rasshifrovka_listov.py`.

### Шаг 1 — Загрузить источники
```python
import json
from pathlib import Path
from openpyxl import load_workbook
import config
from utils.com_connect import connect_erp

mapping_raw = json.loads((config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8"))
# dict: sheet_name → {struct_uuid, match_type, ...}
mapping_by_sheet = {m["sheet_name"]: m for m in mapping_raw["mappings"]}

# Список match_type, для которых Подразделение подставляется автоматически:
AUTO_MATCH_TYPES = {"exact", "fuzzy", "manual_override"}
```

### Шаг 2 — Читать листы Excel напрямую (НЕ из 01_raw_sheets.json)

01_raw_sheets.json отфильтровал summary-листы. Для `А_РасшифровкаЛистов` нужны **ВСЕ видимые** листы — их финансист сам увидит и решит что куда.

```python
def read_visible_sheets(excel_path: str) -> list[str]:
    wb = load_workbook(excel_path, read_only=True)
    names = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    wb.close()
    return names
```

### Шаг 2a — Backup существующей ТЧ (ОБЯЗАТЕЛЬНО перед очисткой)

Перед любой модификацией дампим текущую ТЧ в `data/json/13_backup_<номердок>_<YYYYMMDD_HHMMSS>.json`:

```python
from datetime import datetime
from utils.com_connect import uuid_str

def backup_existing_rows(conn, doc_ref, out_dir: Path):
    obj = doc_ref.ПолучитьОбъект()
    rows = []
    for row in obj.Расшифровка:
        rows.append({
            "ИмяЛиста": str(row.ИмяЛиста),
            "ПодразделениеСтрока": str(row.ПодразделениеСтрока),
            "Подразделение_uuid": uuid_str(conn, row.Подразделение) if row.Подразделение else None,
            "Подразделение_имя": str(row.Подразделение) if row.Подразделение else None,
        })
    if not rows:
        return None  # пустая ТЧ — backup не нужен
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = out_dir / f"13_backup_{str(obj.Номер).strip()}_{ts}.json"
    out.write_text(
        json.dumps({"doc_number": str(obj.Номер), "backed_up_at": ts,
                    "rows_count": len(rows), "rows": rows},
                   ensure_ascii=False, indent=2),
        encoding="utf-8")
    return out
```

Backup делается **даже в режиме `--dry-run`** (операция read-only).

### Шаг 3 — Найти документ в базе (поиск по basename, запись полного пути)

В базе `ИмяФайла` хранится как `!PL по компании Декабрь 2025.xlsx`, в config.py — полный путь. Ищем по basename, при записи обновляем до полного пути.

**КРИТИЧНО:** `ИмяФайла` — **Строка неограниченной длины** (memo). Оператор `=` запрещён («Нельзя сравнивать поля неограниченной длины»). Использовать только `ПОДОБНО` с префиксным шаблоном `%basename`, чтобы покрыть оба случая: в базе только basename ИЛИ уже полный путь.

```python
from pathlib import Path

def find_doc_by_file(conn, excel_path: str):
    basename = Path(excel_path).name  # '!PL по компании Декабрь 2025.xlsx'
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ Документ.А_РасшифровкаЛистов
    ГДЕ ИмяФайла ПОДОБНО &BasenameLike И НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
    q.УстановитьПараметр("BasenameLike", f"%{basename}")  # совпадает и с basename, и с полным путём
    tz = q.Выполнить().Выгрузить()
    if tz.Количество() > 0:
        return tz.Получить(0).Ссылка
    return _find_doc_by_period(conn, excel_path)

def _find_doc_by_period(conn, excel_path: str):
    """Fallback: если basename не нашёлся (напр. переименовали файл)."""
    # ef['period'] = 'YYYY-MM-DD' → находим ef по совпадению пути, берём period
    ef = next((e for e in config.EXCEL_FILES if e["path"] == excel_path), None)
    if not ef:
        return None
    y, m, d = map(int, ef["period"].split("-"))
    dt = datetime(y, m, d, 12, 0, 0)
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ Документ.А_РасшифровкаЛистов
    ГДЕ Дата = &Дата И НЕ ПометкаУдаления"""
    q.УстановитьПараметр("Дата", dt)
    tz = q.Выполнить().Выгрузить()
    return tz.Получить(0).Ссылка if tz.Количество() > 0 else None
```

**Если ни basename, ни дата не нашли** — создать новый через `conn.Документы.А_РасшифровкаЛистов.СоздатьДокумент()` с `Дата = dt` и `ИмяФайла = excel_path` (полный путь). Для текущих 3-х документов этот путь кода НЕ должен срабатывать.

### Шаг 4 — Заполнить ТЧ для каждого документа

Функция принимает `dry_run: bool`. Если `True` — не делает `Очистить()/Добавить()/Записать()`, но проходит весь алгоритм и возвращает детальный разбор по каждой строке. Возвращает не только агрегаты, но и `rows` — для per-row лога.

```python
def fill_rasshifrovka(conn, doc_ref, excel_path, sheets, mapping, dry_run=False):
    obj = doc_ref.ПолучитьОбъект()
    rows_detail = []

    if not dry_run:
        obj.ИмяФайла = excel_path  # полный путь — для будущих запусков
        obj.Расшифровка.Очистить()

    for sheet_name in sheets:
        m = mapping.get(sheet_name)
        detail = {
            "sheet_name": sheet_name,
            "match_type": m["match_type"] if m else "absent",
            "struct_uuid": None,
            "struct_name": None,
            "action": "empty",
        }

        if not dry_run:
            row = obj.Расшифровка.Добавить()
            row.ИмяЛиста = sheet_name
            row.ПодразделениеСтрока = sheet_name[:150]

        if m and m.get("match_type") in AUTO_MATCH_TYPES and m.get("struct_uuid"):
            detail["struct_uuid"] = m["struct_uuid"]
            detail["struct_name"] = m.get("struct_name")
            detail["action"] = "filled"
            if not dry_run:
                uid = conn.NewObject("УникальныйИдентификатор", m["struct_uuid"])
                row.Подразделение = conn.Справочники.СтруктураПредприятия.ПолучитьСсылку(uid)

        rows_detail.append(detail)

    if not dry_run:
        obj.Записать()  # БЕЗ проведения — движений нет

    total = len(rows_detail)
    with_dep = sum(1 for r in rows_detail if r["action"] == "filled")
    return {
        "total": total,
        "with_dep": with_dep,
        "empty_dep": total - with_dep,
        "rows": rows_detail,
    }
```

### Шаг 5 — Главный цикл

Поддержать флаг `--dry-run` через argparse. При `--dry-run`:
- Печатать `[DRY-RUN]` префикс у всех логов изменения.
- Не вызывать `Записать()`.
- Имя лога: `13_fill_rasshifrovka_log_dryrun.json` (обычный: `13_fill_rasshifrovka_log.json`).
- Backup делать всё равно (он read-only).

```python
import argparse
from datetime import datetime

def read_visible_and_hidden(excel_path):
    wb = load_workbook(excel_path, read_only=True)
    vis, hid = [], []
    for ws in wb.worksheets:
        (vis if ws.sheet_state == "visible" else hid).append(ws.title)
    wb.close()
    return vis, hid

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="Напечатать план заполнения без модификации базы.")
    args = ap.parse_args()

    mode = "DRY-RUN" if args.dry_run else "LIVE"
    print(f"=== Режим: {mode} ===\n")

    conn = connect_erp()
    mapping_raw = json.loads((config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8"))
    mapping_by_sheet = {m["sheet_name"]: m for m in mapping_raw["mappings"]}

    entries = []
    for ef in config.EXCEL_FILES:
        path, label = ef["path"], ef["label"]
        print(f"--- {label} ---")
        visible, hidden = read_visible_and_hidden(path)
        print(f"  Листов видимых: {len(visible)}, скрытых (пропущены): {len(hidden)}")

        doc_ref = find_doc_by_file(conn, path)
        if not doc_ref:
            raise RuntimeError(f"Документ для '{path}' не найден")

        backup_file = backup_existing_rows(conn, doc_ref, config.JSON_DIR)
        if backup_file:
            print(f"  Backup: {backup_file.name}")

        res = fill_rasshifrovka(conn, doc_ref, path, visible, mapping_by_sheet,
                                dry_run=args.dry_run)
        print(f"  Всего: {res['total']}, с Подразделением: {res['with_dep']}, "
              f"пустых: {res['empty_dep']}")

        entries.append({
            "file": label,
            "excel_path": path,
            "doc_uuid": uuid_str(conn, doc_ref),
            "total": res["total"],
            "with_dep": res["with_dep"],
            "empty_dep": res["empty_dep"],
            "rows": res["rows"],
            "skipped_hidden": hidden,
            "backup_file": backup_file.name if backup_file else None,
        })

    suffix = "_dryrun" if args.dry_run else ""
    out = config.JSON_DIR / f"13_fill_rasshifrovka_log{suffix}.json"
    out.write_text(
        json.dumps({"run_at": datetime.now().isoformat(),
                    "mode": mode, "entries": entries},
                   ensure_ascii=False, indent=2),
        encoding="utf-8")
    print(f"\nЛог: {out}")

if __name__ == "__main__":
    main()
```

## Правила маппинга Подразделения

| `match_type` из 07_mapping | Действие |
|---|---|
| `exact` (confidence = 100) | **Ставим UUID** |
| `fuzzy` (confidence ≥ 85) | **Ставим UUID** (очевидные опечатки: Укрнафта1модуль → Укрнафта 1 Модуль, Черноморск → Чорноморськ, МКБориспільська → МК Бориспільска) |
| `manual_override` | **Ставим UUID** (решение финансиста: Техника → Логистика, ПРОООН Черкаси ДСНС → МД ПРООН Черкаси  ДСНС — с двумя пробелами) |
| `none` | **Оставляем пусто** |
| Листа нет в 07_mapping | **Оставляем пусто** (summary-листы типа `PL_ЦО_Строительство`, `PL_Свод`, и новые листы — финансист допишет) |

Во всех случаях `ПодразделениеСтрока = ИмяЛиста` (привязка к Excel сохраняется).

## Acceptance criteria

### Dry-run (выполнять ПЕРВЫМ)

1. `python scripts\13_fill_rasshifrovka_listov.py --dry-run` завершается без Exception.
2. В stdout префикс `=== Режим: DRY-RUN ===`, для каждого файла выведено «Листов видимых/скрытых» и «Всего/с Подразделением/пустых».
3. `data/json/13_fill_rasshifrovka_log_dryrun.json` создан. Содержит 3 entries с `rows[]`, `skipped_hidden[]`.
4. Для документа №000000001 создан backup: `data/json/13_backup_000000001_*.json` с 4 строками (исходные ручные правки финансиста).
5. База не изменилась (повторный MCP-запрос по `ТЧ.Расшифровка` даёт ту же картину — №1 содержит 4 строки, №2 и №3 пусты).

### Боевой запуск (без `--dry-run`)

1. Скрипт завершается без Exception.
2. `data/json/13_fill_rasshifrovka_log.json` создан, содержит 3 записи с полями `{file, excel_path, doc_uuid, total, with_dep, empty_dep, rows[], skipped_hidden[], backup_file}`.
3. Backup документа №000000001 создан (`13_backup_000000001_*.json`).
4. MCP-проверка:
   ```
   ВЫБРАТЬ Док.Номер, Док.Дата, Док.ИмяФайла,
          КОЛИЧЕСТВО(ТЧ.НомерСтроки) КАК Строк,
          СУММА(ВЫБОР КОГДА ТЧ.Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
                     ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд
   ИЗ Документ.А_РасшифровкаЛистов КАК Док
   ЛЕВОЕ СОЕДИНЕНИЕ Документ.А_РасшифровкаЛистов.Расшифровка КАК ТЧ
     ПО ТЧ.Ссылка = Док.Ссылка
   ГДЕ НЕ Док.ПометкаУдаления
   СГРУППИРОВАТЬ ПО Док.Номер, Док.Дата, Док.ИмяФайла
   УПОРЯДОЧИТЬ ПО Док.Дата
   ```
   Ожидается (подсчёт для 2026-04-18):
   - 3 строки. `Строк`: **31 / 29 / 31** (декабрь/январь/лютий).
   - `СПодразд`: **12 / 12 / 13**.
     - Декабрь — все бизнес-листы из 07_mapping, кроме отсутствующих в файле (`ПРОООН Черкаси ДСНС`, `Черноморск`).
     - Январь — все, кроме `МКБориспільська`, `ПРОООН Черкаси ДСНС`.
     - Лютий — все, кроме `МКБориспільська`.
   - `ИмяФайла` у всех 3-х = **полный путь** `C:\Configuration_downloads\...\*.xlsx`.
5. В клиенте 1С открыть документ №000000001:
   - ТЧ: 31 строка со всеми видимыми листами Декабря.
   - `Подразделение` заполнено для 12 листов: Крушинка, Ретал.ГПУ, МХП Катеринополь, Астарта. Тищенки, КМД-2, Нежин 4, Глобино-2, МКБориспільська, Кернел Кухня, Укрнафта1модуль, КПП Червоний хрест, Техника (= Логистика).
   - Пусто для `PL_Свод`, `PL_ЦО_Строительство`, `PL_ЦО`, `MAN *`, `ЦО_*`, `Метрики`, и др.
   - `ПодразделениеСтрока` у всех 31 строк = `ИмяЛиста`.

## Что НЕ делать

- **НЕ запускать** боевой режим без предварительного `--dry-run` и визуального осмотра лога.
- **НЕ пропускать** backup — даже если ТЧ выглядит пустой, `backup_existing_rows` обязателен перед `Очистить()` (функция сама решит, сохранять ли файл).
- **НЕ искать** документ по полному пути — в базе хранится только basename. Искать через `Path(excel_path).name`, записывать `obj.ИмяФайла = excel_path` (полный путь).
- **НЕ создавать** новые документы, если 3 существующих найдены. Fallback по дате применяется только если basename не нашёлся (переименование файла). Создание нового — лишь как крайний случай.
- **НЕ угадывать** Подразделение для листов, которых нет в 07_mapping (оставлять пустым).
- **НЕ фильтровать** summary-листы — все видимые (`sheet_state == "visible"`) попадают в ТЧ.
- **НЕ править** `07_mapping_sheet_to_struct.json` — этот файл управляется `07_match_sheets_to_struct.py`.
- **НЕ проводить** документ — `Записать()` без параметра `Проведение` (движений нет).
- **НЕ обрезать** имя листа при копировании в `ИмяЛиста` (поле достаточное).

## Как запустить

```powershell
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL

# 1. Dry-run — ОБЯЗАТЕЛЬНО перед боевым запуском (база не меняется, лог и backup создаются).
.\..\venv\Scripts\python.exe scripts\13_fill_rasshifrovka_listov.py --dry-run

# 2. Боевой запуск (ТЧ перезаписывается; backup сохраняется в data/json/).
.\..\venv\Scripts\python.exe scripts\13_fill_rasshifrovka_listov.py

# 3. Верификация через MCP (см. Acceptance criteria, боевой запуск, пункт 4).
```

Откат №000000001 к ручным правкам финансиста (если нужно) — взять `data/json/13_backup_000000001_*.json` и восстановить строки руками через клиент 1С (для одного документа быстрее, чем писать скрипт восстановления).

## Прецедент имени из 07_mapping

Внимание на имя `"ПРОООН Черкаси ДСНС"` (3 буквы "О" подряд — опечатка финансиста в листе Excel) → в 07_mapping замаплена через manual_override на `"МД ПРООН Черкаси  ДСНС"` (два пробела после «Черкаси», так в справочнике СтруктураПредприятия). При чтении листов из Excel использовать имя **как есть** — именно под ним ищется маппинг.

## Источник информации

Промт обновлён 2026-04-18 после проверки боевой базы и Excel-файлов:
- Метаданные `Документ.А_РасшифровкаЛистов` — подтверждены через `get_metadata_structure`.
- **`ИмяФайла` в базе = basename** без пути — подтверждено `execute_query` на 2026-04-18.
- **Документ №000000001 не пустой** — 4 строки ручной работы финансиста, включая `PL_ЦО_Строительство → Строительство` (отсутствует в 07_mapping).
- Количество видимых листов Excel: **31 / 29 / 31** (декабрь/январь/лютий) — проверено `openpyxl.sheet_state == 'visible'`.
- Покрытие бизнес-листов из `07_mapping` по файлам: **12 / 12 / 13** из 14:
  - декабрь — нет `ПРОООН Черкаси ДСНС`, `Черноморск`;
  - январь — нет `МКБориспільська`, `ПРОООН Черкаси ДСНС`;
  - лютий — нет `МКБориспільська`.
- Структура `07_mapping_sheet_to_struct.json` — 14 записей (9 exact + 3 fuzzy + 2 manual_override).
- Паттерн COM-работы с документами — заимствован из `scripts/12_upload_documents.py` (`ref_by_uuid`).
- Утилиты подключения — из `scripts/utils/com_connect.py` (`connect_erp`, `uuid_str`).

Если что-то изменилось (финансист что-то поправил, файлы переименованы) — перед написанием скрипта перепроверь через MCP `execute_query` и `openpyxl`, не додумывай.
