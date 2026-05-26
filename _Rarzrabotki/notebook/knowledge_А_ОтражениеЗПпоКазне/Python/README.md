# Python/ — диагностический пайплайн knowledge_А_ОтражениеЗПпоКазне

> 4 скрипта Python COM, проверяющие/документирующие реальную структуру 3-х внешних баз (BuhKazn, zup_1, zup_2) + ERP-документа.

---

## Зависимости

- `C:\Python313\python.exe` (имеет pywin32)
- COM-доступ к серверу SQLSERVER через `V83.COMConnector`
- Connection-строки → см. [KNOWLEDGE_MAP.md §«Связанные внешние объекты»](../KNOWLEDGE_MAP.md)

---

## Список скриптов

### `test/discovery_kazna_bdds_struktura.py`

**Назначение:** документирует структуру `РегистрНакопления.БДДС` в базе BuhKazn (Казна) — измерения, ресурсы, реквизиты. Делает тестовый запрос на 5 строк (любой проведённый РаспределениеЗП/РаспределениеФ2). Извлекает UUID регистратора через `kazna.string(...)` (lowercase).

**Запуск:**
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_А_ОтражениеЗПпоКазне\Python\test\discovery_kazna_bdds_struktura.py"
```

**Артефакт:** `_artifacts/kazna_bdds_schema.md`

**Ожидаемый вывод:** 9 измерений, 4 ресурса, 5 реквизитов + таблица 5 строк примеров + UUID регистратора[0].

---

### `test/discovery_zup_struktura.py`

**Назначение:** документирует структуру **3-х регистров zup_1** (ВзаиморасчетыПоНДФЛ, ВзносыВФонды, УдержанияРаботниковОрганизаций) и **документа `НачислениеЗарплатыРаботникам` в zup_2** (Реквизиты + 3 ТЧ).

**Запуск:**
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_А_ОтражениеЗПпоКазне\Python\test\discovery_zup_struktura.py"
```

**Артефакт:** `_artifacts/zup_schema.md`

**Ожидаемый вывод:** метаданные регистров/документа + 3 строки тестового запроса (если ДатаВремя в COM-параметре работает; иначе skipped — структура всё равно собрана).

**Known issue:** тест-запросы могут упасть с ошибкой `Тип не определен 'ДатаВремя'` — это известная проблема COM-маршаллинга tz-aware дат. Метаданные при этом извлекаются корректно. См. [LESSONS.md §10](../LESSONS.md).

---

### `test/verify_uuid_lookup_kazna_erp.py`

**Назначение:** проверяет cross-base UUID lookup 1:1 для 2-х типов документов:
- Казна.`Документ.РаспределениеЗаработнойПлаты` → ERP.`Документ.А_РаспределениеЗаработнойПлаты`
- Казна.`Документ.РаспределениеФ2` → ERP.`Документ.РаспределениеФ2`

Для каждого:
1. Находит последний проведённый документ в Казне
2. Извлекает UUID через `kazna.string(doc.Ссылка.УникальныйИдентификатор())`
3. Создаёт `УникальныйИдентификатор` в контексте ERP
4. `Документы.X.ПолучитьСсылку(уид).ПолучитьОбъект() <> None`
5. Сравнивает Номер + Дата (должны совпадать)

**Запуск:**
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_А_ОтражениеЗПпоКазне\Python\test\verify_uuid_lookup_kazna_erp.py"
```

**Артефакт:** `_artifacts/uuid_pair_example.md`

**Ожидаемый вывод:** 2 пары UUID с пометкой ✓ найден + совпадение Номер/Дата.

---

### `test/verify_test_doc_loading.py`

**Назначение:** snapshot реального документа `А_ОтражениеЗПпоКазне` (последний по дате) — количество строк в каждой из 7 ТЧ + первые 3 строки каждой (с базовыми реквизитами).

**Запуск:**
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_А_ОтражениеЗПпоКазне\Python\test\verify_test_doc_loading.py"
```

**Артефакт:** `_artifacts/test_doc_sample.md`

**Ожидаемый вывод:** заголовок документа + размеры 7 ТЧ + по 3 строки каждой ТЧ (значения примитивов читаемы, COM-references отображаются как `<COMObject <unknown>>`).

---

## Структура папки

```
Python/
├── README.md                       (этот файл)
├── test/
│   ├── discovery_kazna_bdds_struktura.py
│   ├── discovery_zup_struktura.py
│   ├── verify_uuid_lookup_kazna_erp.py
│   └── verify_test_doc_loading.py
└── _artifacts/
    ├── kazna_bdds_schema.md
    ├── zup_schema.md
    ├── uuid_pair_example.md
    └── test_doc_sample.md
```

---

## Запуск всех 4 скриптов

```powershell
$py = "C:\Python313\python.exe"
$dir = "C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_А_ОтражениеЗПпоКазне\Python\test"
& $py "$dir\discovery_kazna_bdds_struktura.py"
& $py "$dir\discovery_zup_struktura.py"
& $py "$dir\verify_uuid_lookup_kazna_erp.py"
& $py "$dir\verify_test_doc_loading.py"
```

Или через MCP `python-runner`:
```python
mcp__python-runner__run_command(command='C:\\Python313\\python.exe "..."')
```

---

## Регрессия и валидация

После значимых изменений в `Documents/А_ОтражениеЗПпоКазне/Ext/ObjectModule.bsl` или в обмене ConvertERP/Казна — стоит:

1. Прогнать все 4 скрипта
2. Сравнить `_artifacts/uuid_pair_example.md` с golden state в [FINDINGS.md §4](../FINDINGS.md) — UUID должен по-прежнему мапиться 1:1
3. Сравнить размеры ТЧ из `_artifacts/test_doc_sample.md` с ожидаемыми (для эталонного документа `№000000005` от 30.04.2026 — 7 ТЧ заполнены, см. [FINDINGS.md §5](../FINDINGS.md))

Если размеры значительно отличаются (например, было 926 строк РаспределениеКазна, стало 50) — высока вероятность регрессии в `ЗагрузитьРаспределениеКазна_ОтражениеЗП` или сбой обмена в `ConvertERP/`.

---

## Cross-references

- [KNOWLEDGE_MAP.md](../KNOWLEDGE_MAP.md)
- [FINDINGS.md](../FINDINGS.md) — golden state артефактов
- [LESSONS.md](../LESSONS.md) — антипаттерны (особенно §10 про TZ-чувствительные datetime)
- CLAUDE.md «Rule #-1 — Python COM test BEFORE BSL»
- CLAUDE.md «Cross-base UUID lookup»
