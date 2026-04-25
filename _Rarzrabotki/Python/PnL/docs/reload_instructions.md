# Инструкция: перезаливка 3-х месяцев документов А_ОтчетPL с фиксами маппинга

**Контекст**: после правки `config.py` (убрана «Техника» из SUMMARY_SHEET_PATTERNS, добавлены MANUAL_*_OVERRIDES) и `07_mapping_sheet_to_struct.json` (добавлены Техника→Логистика, ПРОООН→МД ПРООН) требуется пересобрать пайплайн и перезалить документы в 1С.

---

## Предварительные проверки (перед перезаливкой)

1. **Excel-файлы на месте** (декабрь/январь/лютий). Пути в `config.py`:
   - `C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx`
   - `C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Январь_26\!PL по компании Январь 2026.xlsx`
   - `C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Лютий_26\!PL по компании Лютий 2026.xlsx`
2. **ERP доступна** через COM. Тест: `python scripts/_verify_excel_vs_pl.py` — должен вернуть 272 строки без ошибок.
3. **Backup** текущих документов А_ОтчетPL (опционально, пайплайн идемпотентен).

---

## Пошаговая команда

Запускать из директории `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL`:

```powershell
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL

# 1. Пересобрать Excel→JSON (лист «Техника» теперь парсится)
python scripts/01_extract_excel_to_json.py

# 2. Уникальные статьи (обнаружит новые статьи «Техника»-листа)
python scripts/02_extract_unique_articles.py

# 3. ДДС из 1С (обновить справочник)
python scripts/03_extract_dds_from_erp.py

# 4. Маппинг статья→ДДС (с применением MANUAL_DDS_OVERRIDES)
python scripts/04_match_articles.py
# Ожидаемый вывод: manual_override=5 (Логистические, Телекомуникационные, Юридические, ИТ, Услуги банков)

# 4b. Маппинг доходных статей
python scripts/04b_match_statya_dohodov.py
# Ожидаемый вывод: manual_override=3

# 5. Группы PL (создаёт 03_groups_pl.json БЕЗ UUID)
python scripts/05_build_groups.py

# 6. Структура подразделений из 1С
python scripts/06_extract_struct_from_erp.py

# 7. Маппинг листов Excel ↔ подразделения (с MANUAL_SHEET_TO_STRUCT_OVERRIDES)
python scripts/07_match_sheets_to_struct.py
# Ожидаемый вывод: manual_applied=X, manual_added=Y (Техника + ПРОООН)

# === КРИТИЧЕСКИЙ ПОРЯДОК: 10 и 11 ДО 08 ===
# 10 и 11 заливают группы и статьи в 1С и ОБНОВЛЯЮТ JSON-файлы UUID'ами.
# Если запустить 08 раньше — 08 прочитает JSON БЕЗ UUID'ов, и в документах
# А_ОтчетPL поле "Группа" в ТЧ ИтогиПоГруппам будет ПУСТЫМ.

# 10. Залить группы PL (обновляет 03_groups_pl.json с UUID-ами)
python scripts/10_upload_groups.py

# 11. Залить/синхронизировать статьи PL (обновляет 02_unique_articles.json)
python scripts/11_upload_articles.py

# 8. Подготовить документы к загрузке (теперь берёт UUID-ы из обновлённых JSON)
python scripts/08_prepare_documents.py
# Ожидаемое количество документов: 34 + 3 новых (Техника × 3 месяца)

# 12. Залить документы А_ОтчетPL (идемпотентно)
python scripts/12_upload_documents.py
```

---

## Проверка результата

После 12_upload_documents.py запустить верификационные скрипты:

```powershell
# Контроль 1: Excel ↔ документ А_ОтчетPL
python scripts/_verify_excel_vs_pl.py
# Должно быть 272 строки, ≥ 247 OK (допустимы 25 с техническим расхождением по "Финансовая деятельность")

# Контроль 2: документ А_ОтчетPL ↔ АДР (1С-факт)
python scripts/_verify_pl_vs_adr.py
# Показывает реальные расхождения PL vs АДР по каждому (подразделению, периоду)
# Смотреть файл: data/reports/verification_pl_vs_adr.csv
```

Проверить в 1С:
1. Открыть документ А_ОтчетPL для нового подразделения «Логистика» — декабрь 2025 (если в Excel есть лист «Техника» за декабрь).
2. Открыть документ для МД ПРООН Черкаси ДСНС — декабрь 2025.
3. На документе № 000000029 (Астарта декабрь) — нажать «Открыть сверку с 1С», убедиться что отчёт открывается с правильным периодом и подразделением.

---

## Откат (если что-то сломалось)

Пайплайн идемпотентен — повторный запуск 12_upload_documents.py перезапишет документы. Если нужно полностью откатить:

```powershell
python scripts/09_delete_existing_docs.py  # удалит все документы А_ОтчетPL
# Затем восстановить старый config.py из git и прогнать 01-12 заново
```

Метаданные (документы, справочники, реквизиты) — **не меняются** этим пайплайном; только данные в документах.
