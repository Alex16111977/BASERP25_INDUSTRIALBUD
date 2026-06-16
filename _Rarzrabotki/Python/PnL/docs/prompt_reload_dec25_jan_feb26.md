# ПРОМТ: Перезалить Декабрь 2025 + Январь 2026 + Лютий 2026 в 1С (А_ОтчетPL)

Каталог: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL`. Интерпретатор: `..\venv\Scripts\python.exe`
(fallback `C:\Python313\python.exe`). ERP: `config.CONN_ERP` (SQLSERVER/BaseERP/Администратор/24043). Орг — ТОВ "ІНДАСТРІАЛБУД".

## Цель
Идемпотентно **перезалить** план PnL за **2025-12, 2026-01, 2026-02** из обновлённого Excel:
- `..\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx` (period 2025-12-31)
- `..\PL\Январь_26\!PL по компании Январь 2026.xlsx` (period 2026-01-31)
- `..\PL\Лютий_26\!PL по компании Лютий 2026.xlsx` (period 2026-02-28)

⚠️ Декабрьский файл = **12 month-колонок за весь 2025** (col=19..74, step 5). Это значит запись «Декабрь 2025» в
`config.EXCEL_FILES` и записи 2025-01..11 (тот же path) — НЕ путать; фильтруем строго `--month 2025-12`.
`config.EXCEL_FILES` **уже содержит** все 3 месяца — **config не менять**.

## Состояние ДО (проверить через MCP/ГОД(Дата), без string-параметра-даты!)
А_ОтчетPL: **2025-12=28, 2026-01=26, 2026-02=28** (проведены). А_РасшифровкаЛистов: №1(дек), №2(янв), №3(фев).
Документ №000000001 (дек) — эталон с ручными правками финансиста: **НЕ удалять**.

## 12 (idempotent: find-or-create по Дата+ПодразделениеСтрока) — обновит существующие, дублей не плодит.

## Алгоритм
```powershell
$py = "..\venv\Scripts\python.exe"; $d = "scripts"
# 1) Каноники 01-07 (БЕЗ --month) — рефреш JSON из обновлённого Excel
foreach ($s in "01_extract_excel_to_json","02_extract_unique_articles","03_extract_dds_from_erp","04_match_articles","04b_match_statya_dohodov","05_build_groups","06_extract_struct_from_erp","07_match_sheets_to_struct") { & $py "$d\$s.py" }
# 2) Каталоги (backfill UUID в 02_unique_articles.json — КРИТИЧНО перед 08!)
& $py "$d\_mass_set_manual_flag.py" --dry-run   # если НеLocked>0 → прогнать без --dry-run
& $py "$d\10_upload_groups.py"; & $py "$d\11_upload_articles.py"
# 3) По месяцам: 13 (dry→live) → cleanup → 08 → 12
foreach ($m in "2025-12","2026-01","2026-02") {
  & $py "$d\13_fill_rasshifrovka_listov.py" --month $m --dry-run
  & $py "$d\13_fill_rasshifrovka_listov.py" --month $m
}
& $py "$d\_cleanup_rasshifrovka_garbage.py"
foreach ($m in "2025-12","2026-01","2026-02") {
  & $py "$d\08_prepare_documents.py" --month $m
  & $py "$d\12_upload_documents.py" --month $m --limit 1
  & $py "$d\12_upload_documents.py" --month $m
}
```

## 🚨 Граблі (из инцидента 2026-06-09 — ОБЯЗАТЕЛЬНО)
1. **Повторный `02` СТИРАЕТ uuid статей** → если после правок прогнать `01`+`02`, ОБЯЗАТЕЛЬНО снова `10`+`11`
   ПЕРЕД `08`. Иначе `12` зальёт документы **ПУСТЫМИ** (`rows=0`, ДанныеОтчета пустые). Контроль: после `11`
   проверить `data/json/02_unique_articles.json` → `with_uuid == len(articles)`.
2. **Фейк-статьи** уже фильтруются автоматически (Шар 4: `config.REPORT_END_MARKERS` + `data/fake_articles.json` +
   `excel_parser`). После `02` проверить `02_unique_articles.json` на новые подозрительные имена (числа/договоры в
   названии после итогов) — если просочились, добавить в `data/fake_articles.json` (1 строка) и повторить `01`+`02`+`10`+`11`.
   Уже созданные фейки чистит `scripts/_cleanup_fake_articles.py`.
3. `13 --month` ПЕРЕЗАПИСЫВАЕТ ТЧ Расшифровка месяца (backup в `data/json/13_backup_*.json`). №1(дек) с ручными
   правками — если важны, перенести в `MANUAL_SHEET_TO_STRUCT_OVERRIDES` до прогона.

## НЕ запускать: `09_delete_existing_docs.py` (снесёт всё); `13` без `--month` (затрёт №1-№3); `17_upload_pl_sort.py`.

## Верификация (серверные ГОД/НАЧАЛОПЕРИОДА, НЕ string-дата в MCP)
```sql
ВЫБРАТЬ НАЧАЛОПЕРИОДА(ТЧ.Ссылка.Дата,МЕСЯЦ) КАК М, КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ТЧ.Ссылка) КАК Док,
  КОЛИЧЕСТВО(*) КАК Строк, СУММА(ТЧ.Сумма) КАК Σ
ИЗ Документ.А_ОтчетPL.ДанныеОтчета КАК ТЧ ГДЕ ГОД(ТЧ.Ссылка.Дата) В (2025,2026) И НЕ ТЧ.Ссылка.ПометкаУдаления
СГРУППИРОВАТЬ ПО НАЧАЛОПЕРИОДА(ТЧ.Ссылка.Дата,МЕСЯЦ) УПОРЯДОЧИТЬ ПО М
```
Критерий: 2025-12 / 2026-01 / 2026-02 — **Док=ДокСДанными** (нет пустых, Σ≠0), счётчики ≈28/26/28 (могут
измениться, если в Excel добавили/убрали листы — сверить с числом непустых листов). Прочие месяцы (2024, 2025-01..11,
2026-03/04) — **без изменений**. А_РасшифровкаЛистов 2025-12..2026-02 не задвоились. Orphan (лист скрыт в Excel,
док остался) — пометить на удаление **вручную в 1С** (агент авто-удалять боевые доки не вправе).

## Источники
`docs/prompt_reload_mar_apr_2026.md` (эталон того же паттерна), `docs/reload_instructions.md` (порядок 10/11 ДО 08),
`README.md`, `docs/fake_articles_registry.md`, memory `pnl_monthly_import_pattern` / `pnl_excel_fake_articles_filter`.
