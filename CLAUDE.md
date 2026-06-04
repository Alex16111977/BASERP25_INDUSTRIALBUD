# CLAUDE.md - AI ASSISTANT INSTRUCTION: BAS ERP 2.5 INDUSTRIALBUD

## 🚨🚨🚨 RULE #-1 — MOST CRITICAL — TEST ALL QUERIES IN PYTHON FIRST 🚨🚨🚨

**ABSOLUTE MANDATORY RULE — НІ ОДНОГО ВИКЛЮЧЕННЯ:**

**Перш ніж писати ЛЮБИЙ запит мовою 1С (1C Query Language) у код модуля, обробки, документа — ОБОВ'ЯЗКОВО спочатку перевірити цей запит через Python COM тест.**

### Чому це критично:
- Запити 1С мають **зарезервовані слова** які складно запам'ятати (`ПО`, `ИЗ`, `ГДЕ`, `КАК` тощо)
- Алиаси типу `ПО` (Подразделение Организаций) ламають синтаксис `ЛЕВОЕ СОЕДИНЕНИЕ ... ПО ...`
- Поля у віртуальних таблицях `.Обороты()` мають префікси `Приход/Расход/Оборот` — без тесту здогадатися неможливо
- Метадані мають дублікати кодів, неочевидні типи реквізитів — без перевірки = помилки в продакшені
- **Помилка у запиті ламає всю функцію** — користувач отримує `Ошибка при вызове метода контекста (Выполнить)` під час бойової роботи

### Обов'язковий процес:

```
1. Написати запит у Python (1:1 як буде у BSL)
2. Виконати через COM: erp.NewObject("Запрос") + .Execute().Выгрузить()
3. Якщо помилка — виправити, тестувати знову
4. ТІЛЬКИ після успішного запуску → перенести запит у BSL код
5. F7 — оновити конфігурацію
```

### Шаблон Python тесту запиту:
```python
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ...
ИЗ ...
"""
q.SetParameter("П", значення)
try:
    r = q.Execute().Выгрузить()
    print(f"OK, рядків={r.Количество()}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL: {e.excepinfo[2]}")
    else:
        print(f"FAIL: {e}")
```

### Антипаттерни (НЕ РОБИТИ):
- ❌ "Я знаю синтаксис, тестувати не буду" — **ВСІ помилки запитів сталися саме через цю думку**
- ❌ Скопіювати запит з іншого модуля без перевірки на реальній базі — структура регістру могла змінитись
- ❌ Використовувати алиаси що збігаються з ключовими словами (`ПО`, `ИЗ`, `КАК`)
- ❌ Покладатися на метадані без `execute_query` для перевірки реальних полів віртуальних таблиць

### Pattern, який ВЖЕ знайшов помилки:
- `ВзаиморасчетыПоНДФЛ.Доход` → треба `ДоходПриход` (бо це сальдовий регістр)
- `ЛЕВОЕ СОЕДИНЕНИЕ ... КАК ПО` → треба `КАК ПодрОрг` (ПО — зарезервоване)
- `Сотрудник.Физлицо.Наименование` у `.Обороты()` де Сотрудник — **реквізит**, не вимірювання → недоступний
- `НалогиБухгалтерия.СуммаДохода` поле додано пізніше → перші запити повертали 0

**Lesson learned:** Кожна помилка яку користувач отримав у Designer/Enterprise — це наслідок мого ігнорування цього правила. Жодного винятку: **немає Python тесту = немає коду в BSL**.

---

## CRITICAL RULES (READ AT THE START OF EVERY SESSION!)

### Rule #0 - Golden Sequence for Metadata
BEFORE ANY 1C query:
1. `list_metadata_objects` -> find objects
2. `get_metadata_structure` -> get EXACT structure
3. **`Python тест запиту`** -> перевірити синтаксис на реальній базі (Rule #-1!)
4. `execute_query` -> use ONLY real field names

**NEVER invent field names, registers, documents!**

### Rule #1 - CRUD Operations
- `find_document_ref` -> get UUID BEFORE modification
- `unpost_document` -> cancel posting BEFORE modification
- `update_document` -> modify data
- `post_document` -> post after changes

### Rule #2 - Step-by-Step Development
Break the task into steps. Test after each step!

### Rule #5 - ALWAYS Consult NotebookLM Before Writing Code for Documents
When creating/modifying document code (especially programmatic document creation):
1. **ALWAYS query NotebookLM FIRST** — ask about hidden fields, flags, side effects
2. If NotebookLM returns empty — `refresh_auth` and retry, or READ `_Rarzrabotki/notebook/knowledge/document_hidden_fields.md` directly
3. **NEVER skip this step** even if you "think you know" — the knowledge base has critical details

**Critical flags for programmatic document creation:**
- `А_Обработан = Истина` — WITHOUT THIS: exchange procedure will overwrite your data on next save!
- `А_ВведенВЕРП = Истина` — WITHOUT THIS: exchange procedure clears Товары, РасшифровкаПлатежа!
- `А_ОбработанКазна = Истина` — for ПКО/РКО: WITHOUT THIS: РасшифровкаПлатежа gets cleared!
- `Курс = 1, Кратность = 1` — WITHOUT THIS: СуммаВзаиморасчетов = 0!

**Lesson learned:** Skipped NotebookLM consultation → wrote А_Обработан=Ложь → systemic bug that corrupts document data on next save. The knowledge base explicitly documents the correct values.

### Rule #3 - Query Language
All 1C queries use Russian field names. Example:
```
SELECT Ref, Date, Number FROM Document.ПоступлениеТоваровУслуг
```

### Rule #4 - ALWAYS Copy Changes to Main Configuration
When working in a worktree (`.claude/worktrees/...`), **ALWAYS** copy modified files back to the main configuration directory `C:\Configuration_downloads\BASERP25\` after editing. Do NOT leave changes only in the worktree — the 1C Designer loads from the main directory.

**🔒 ЕДИНЫЙ ПОРЯДОК worktree (ТОЛЬКО одно расположение):**
- Все worktree — ТОЛЬКО в `C:\Configuration_downloads\BASERP25\.claude\worktrees\`.
- **НИКОГДА** не создавать/использовать `C:\Users\<user>\.claude-worktrees\` (глобальный путь = «разбериха» с двумя источниками). Если он появляется — причина во внешней env-переменной `CLAUDE_WORKTREE_DIR` или старой версии CLI; чинить у источника, не плодить worktree там.
- Для изоляции использовать ТОЛЬКО инструмент `EnterWorktree` (всегда кладёт worktree в локальный `.claude/worktrees/`).
- Рабочая ветка — `claude/main`. Worktree ветвить от `claude/main`; после подтверждения изменения переносить в `claude/main` (основной каталог, откуда грузит Designer/сервер).
- Папка `.claude/` целиком в `.gitignore` → worktree-контент на GitHub НЕ уходит.

**After every code edit, run `cp` to copy changed files:**
```bash
cp "<worktree_path>/path/to/file.bsl" "C:/Configuration_downloads/BASERP25/path/to/file.bsl"
```

This applies to all .bsl modules, .xml forms, and any other configuration files.

### Rule #5 - Python COM Testing First
BEFORE writing code into 1C processors/modules:
1. Create a Python test in `_Rarzrabotki/Python/test/` (examples already there)
2. Test queries and logic via COM connection (`V83.COMConnector`)
3. Use `mcp__python-runner__run_command` to execute tests
4. Only after successful test — write code into the 1C processor

**Database connections (for Python tests):**
```python
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")

# ERP (BaseERP)
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn_erp = v8.Connect(CONN_ERP)

# BAS Бухгалтерія (BuhBud)
CONN_BUH = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'
conn_buh = v8.Connect(CONN_BUH)
```

**Workflow:** Python test → verify results → write into .bsl module

### Rule #6 - Superpowers Workflow (plugin `superpowers@superpowers-marketplace`)

Для нетривіальних задач застосовуй skills з плагіна `superpowers`:

| Тип задачі | Обов'язковий skill |
|---|---|
| Новий об'єкт метаданих / масова фіча | `brainstorming` → `writing-plans` |
| Баг у BSL / Python | `systematic-debugging` (гіпотези + перевірка через `mcp__1c-workerp__execute_query`) |
| Python-пайплайн | `test-driven-development` (RED → GREEN → REFACTOR) |
| Паралельні правки | `using-git-worktrees` + `dispatching-parallel-agents` |
| Перед "готово" | `verification-before-completion` |
| Закриття гілки | `finishing-a-development-branch` |

**Синергія з існуючими правилами:**
- Rule #-1 + Rule #5 (Python COM тест ПЕРЕД BSL) = RED-фаза TDD
- Rule #0 (Golden Sequence) = обов'язковий крок у `systematic-debugging`
- Rule #1 (CRUD) = чек-лист у `verification-before-completion`
- Rule #4 (копіювати зміни в основну конфігурацію) = завершення `finishing-a-development-branch`

**Доменні 1С-skills** (meta-*, cf-*, cfe-*, epf-*, erf-*, form-*, db-*, skd-*, subsystem-*, role-*, mxl-*, web-*, 1c-bsl-*) викликаються superpowers зсередини як інструменти.

---

## PROJECT STRUCTURE

### Paths
- **Configuration:** `/home/user/BASERP25_INDUSTRIALBUD` (Linux) or `C:\Configuration_downloads\BASERP25` (Windows)
- **External data processors:** `_Rarzrabotki/Обработки/`
- **GitHub:** `Alex16111977/BASERP25_INDUSTRIALBUD`
- **Organization:** LLC INDUSTRIALBUD

### Statistics
| Object Type | Count |
|-------------|-------|
| Documents | 400 |
| Accumulation Registers | 185 |
| Information Registers | 586 |
| Catalogs | 564 |
| Common Modules | 1770+ |

### Key Documents (Procurement/Sales)
| Document | Description |
|----------|-------------|
| `ПоступлениеТоваровУслуг` | Receipt of goods and services |
| `РеализацияТоваровУслуг` | Sales of goods and services |
| `ЗаказКлиента` | Customer order |
| `ЗаказПоставщику` | Purchase order |
| `ПеремещениеТоваров` | Inventory transfer |
| `СписаниеТоваров` | Inventory write-off |
| `ВыпускПродукции` | Production output |
| `ПриходныйКассовыйОрдер` | Cash receipt |
| `РасходныйКассовыйОрдер` | Cash disbursement |
| `ПлатежноеПоручение` | Payment order |
| `АвансовыйОтчет` | Advance report |
| `АктВыполненныхРабот` | Completed work act |

### Key Documents (Payroll - Custom prefix A_)
| Document | Description |
|----------|-------------|
| `А_ОтражениеЗарплатыВУчете` | Salary reflection in accounting |
| `А_РаспределениеЗаработнойПлаты` | Salary distribution |
| `А_ТабельУчетаРабочегоВремени` | Timesheet |
| `А_АрендаТехники` | Equipment rental |
| `А_БюджетМесяц` | Monthly budget |
| `А_ОтгрузкаПродукции` | Product shipment |

### Key Accumulation Registers
| Register | Description |
|----------|-------------|
| `ТоварыНаСкладах` | Inventory on hand |
| `ТоварыОрганизаций` | Organization inventory |
| `СебестоимостьТоваров` | Inventory cost |
| `ВыпускПродукции` | Production output |
| `ВыручкаИСебестоимостьПродаж` | Revenue and COGS |
| `ДенежныеСредстваБезналичные` | Bank funds |
| `ДенежныеСредстваНаличные` | Cash funds |
| `ЗарплатаКВыплате` | Payroll payable |
| `ВзаиморасчетыССотрудниками` | Employee settlements |
| `РасчетыНалогоплательщиковСБюджетомПоНДФЛ` | PIT settlements |
| `РасчетыСКлиентами` | Customer settlements |
| `РасчетыСПоставщиками` | Supplier settlements |

### Key Catalogs
| Catalog | Description |
|---------|-------------|
| `Номенклатура` | Products/Items |
| `Контрагенты` | Counterparties |
| `Партнеры` | Partners |
| `Организации` | Organizations |
| `Сотрудники` | Employees |
| `ФизическиеЛица` | Individuals |
| `Склады` | Warehouses |
| `ДоговорыКонтрагентов` | Contracts |
| `Подразделения` | Departments |
| `СтатьиРасходов` | Expense items |
| `СтатьиДоходов` | Income items |

---

## MCP TOOLS

### Connection
MCP server: `1c-workerp`

### Available Tools
| Tool | Purpose | When to Use |
|------|---------|-------------|
| `list_metadata_objects` | List objects | ALWAYS before query |
| `get_metadata_structure` | Object structure | MANDATORY for queries |
| `execute_query` | Execute queries | ONLY after analysis |
| `find_document_ref` | Document UUID | BEFORE CRUD operations |
| `create_document` | Create document | After get_metadata_structure |
| `update_document` | Update document | After find_document_ref |
| `post_document` | Post document | After update |
| `unpost_document` | Unpost document | BEFORE modifying posted doc |
| `delete_document` | Delete document | Use with caution |
| `execute_code` | Run 1C code | For complex operations |

### Tool Usage Examples

#### Finding Metadata
```javascript
// List all documents
list_metadata_objects({metaType: "Documents", maxItems: 50})

// List accumulation registers
list_metadata_objects({metaType: "AccumulationRegisters", maxItems: 50})

// Get document structure
get_metadata_structure({metaType: "Document", metaName: "ПоступлениеТоваровУслуг"})
```

#### Document Operations
```javascript
// Find document by number
find_document_ref({
  documentType: "ПоступлениеТоваровУслуг",
  searchField: "Number",
  searchValue: "00000001"
})

// Execute query
execute_query({
  query: "SELECT Ref, Date, Number, Контрагент FROM Document.ПоступлениеТоваровУслуг WHERE Date >= &StartDate",
  params: {"StartDate": "2025-01-01"}
})
```

---

## EXTERNAL DATA PROCESSORS

### File Structure
```
ProcessorName.xml
ProcessorName/
  ├── Ext/ObjectModule.bsl
  └── Forms/FormName/
        ├── FormName.xml
        └── Ext/Form.xml + Form/Module.bsl
```

### Available Custom Processors (_Rarzrabotki/Обработки/)
| Processor | Purpose |
|-----------|---------|
| `ЗаполнениеНачисленийЗП` | Fill payroll accruals |
| `Загрузить зарплату из казны` | Import salary from Kazna |
| `Заполнить ОтражениеЗарплатыВФинансовомУчете` | Fill salary reflection |
| `Перенос остатков номенклатуры` | Transfer inventory balances |
| `Перенос остатков денег` | Transfer cash balances |
| `А_ОбновитьСтатьюДенегВЗатратах` | Update cash flow article |
| `РегистрацияИзмененийДляОбменаДанными` | Data exchange registration |

### XML Rules (CRITICAL!)
- Command MUST have `id="number"`
- Action WITHOUT `xsi:type`
- Button: `Type=CommandBarButton`
- All IDs must be unique within form

---

## TYPICAL TASKS

### Analyze Document Movement
1. `get_metadata_structure` -> document structure
2. `execute_query` -> document data
3. Find register movements in configuration
4. `execute_query` -> register records with Recorder filter

### Query Register Movements
```sql
SELECT
    RecordType,
    Period,
    Recorder,
    Номенклатура,
    Склад,
    Quantity
FROM AccumulationRegister.ТоварыНаСкладах
WHERE Recorder = &DocumentRef
```

### Create/Edit External Processor
1. Find a WORKING processor as template
2. Read its Form.xml structure
3. Copy structure
4. Modify only content

### Payroll Operations
1. Check `А_ОтражениеЗарплатыВУчете` for salary reflection
2. Use `А_НачисленнаяЗарплатаИВзносыПоФизлицам` register
3. Link with `ФизическиеЛица` and `Сотрудники` catalogs

---

## COMMON ERRORS

| Error | Cause | Solution |
|-------|-------|----------|
| Field not found | Invented field name | Use `get_metadata_structure` first |
| Query returns empty | Wrong date format | Use ISO format: "2025-01-01" |
| UUID not found | Used execute_query for UUID | Use `find_document_ref` instead |
| Cannot modify posted | Doc is posted | Use `unpost_document` first |
| Command without id | Missing id in XML | Add `id="N"` to Command element |
| xsi:type in Action | Extra attribute | Remove `xsi:type` from Action |

### What NEVER Works
- `execute_query` to get document UUID for modifications
- `update_document(post=false)` to unpost document
- Inventing field names without checking metadata
- Modifying posted document without unposting first

---

## CRITICAL: Cross-base UUID lookup (BuhKazn/BuhBud ↔ ERP via COM)

### Context
When reading registrators from another base via `V83.COMConnector` (e.g. regisTrator from `РегистрНакопления.БДДС` in `BuhKazn`) and you need to find the corresponding document in ERP — the правила обміну (`_Rarzrabotki/ConvertERP/`, `_Rarzrabotki/ExChange/Казна/`) preserve UUID **1:1** between bases. Direct UUID lookup is the only architecturally correct approach (not search by Номер+Дата — that's a workaround).

### Architecture proof (`РегистрСведений.СоответствияОбъектовИнформационныхБаз`)
Standard BSP register maintains the mapping. Узел `ПланОбмена.КазначействоBASERP` contains rows like:
- `УникальныйИдентификаторПриемника` (string) = UUID from BuhKazn
- `УникальныйИдентификаторИсточника` (composite ref) = ERP reference

Exchange rules confirmed 1:1 for these types:
- `РаспределениеЗаработнойПлаты` (Казна) → `А_РаспределениеЗаработнойПлаты` (ERP)
- `РаспределениеФ2` (Казна) → `РаспределениеФ2` (ERP)
- `ПриходныйКассовыйОрдер` → `ПриходныйКассовыйОрдер` (or `А_ПриходДенегОтФинАгента` if `_ФинАгент` filled)
- `ПлатежноеПоручениеВходящее` → `ПоступлениеБезналичныхДенежныхСредств`

### CORRECT pattern for UUID extraction from COM reference

```bsl
// V83 = ранее ОткрытеCOM-зʼєднання до Казни
// СтрТЗ.Регистратор — COM-reference to a document in Kazna

// STEP 1: Extract UUID as string via REMOTE context (V83)
УИДСтр = V83.string(СтрТЗ.Регистратор.УникальныйИдентификатор());

// STEP 2: Create LOCAL UUID object and look up document by type
УИД = Новый УникальныйИдентификатор(УИДСтр);
ДокСсылка = Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(УИД);

// STEP 3: Verify document physically exists (not a broken ref)
Если ДокСсылка.ПолучитьОбъект() <> Неопределено Тогда
    // Found — use ДокСсылка
КонецЕсли;
```

### What NEVER works
| Call | Result | Why |
|------|--------|-----|
| `Строка(V83_UUID)` | empty / `'None'` | ERP-side `Строка()` doesn't know how to serialize foreign COM-UUID |
| `XMLСтрока(V83_UUID)` | **EXCEPTION** "Ошибка преобразования данных XML" | ERP-side XML serializer rejects foreign COM object |
| `V83.XMLСтрока(V83_UUID)` | unreliable via COM dispatch (Cyrillic name issues) | COM dispatch may not resolve Cyrillic method names through V83.COMConnector proxy |
| `V83.String(V83_UUID)` | works in Python, unreliable in BSL | BSL COM proxy prefers lowercase |

### The ONLY reliable call: `V83.string(...)` — **English, lowercase**

This pattern is used in production processors (e.g. `Обработка_ПлатежВЕРП`, `ПометитьНаУдалениеБухАкты`).

### CRITICAL: `Ссылка` is a reserved variable name
**NEVER** declare a local variable named `Ссылка` in BSL — it conflicts with the system `.Ссылка` property on object refs and causes silent failures inside `Попытка/Исключение` blocks. Use `ДокСсылка`, `РезСсылка`, `НайденнаяСсылка` instead.

Example of the bug:
```bsl
// WRONG — Ссылка conflicts with system property
Ссылка = Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(УИД);  // throws silently
// CORRECT
ДокСсылка = Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(УИД);
```

### Alternative: via `СоответствияОбъектовИнформационныхБаз`
If direct UUID lookup doesn't work for your document type, query the BSP register in ERP:

```1c
ВЫБРАТЬ
    СО.УникальныйИдентификаторИсточника КАК СсылкаERP
ИЗ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК СО
ГДЕ СО.УникальныйИдентификаторПриемника = &UUIDКазна
    И ТИПЗНАЧЕНИЯ(СО.УникальныйИдентификаторИсточника) = ТИП(Документ.А_РаспределениеЗаработнойПлаты)
```

### Reference implementation
See `Documents/А_ОтражениеЗПпоКазне/Ext/ObjectModule.bsl` — procedure `ЗагрузитьРаспределениеКазна_ОтражениеЗП()` + function `А_НайтиДокументПоУИД_ОтражениеЗП()`.

### Python diagnostic pattern (MANDATORY before writing BSL)
Before writing UUID-based lookup in BSL, validate in Python:
```python
uuid_kazna = kazna.String(kazna_doc.Ссылка.УникальныйИдентификатор())  # works
uid = erp.NewObject("УникальныйИдентификатор", uuid_kazna)
ref = erp.Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(uid)
obj = ref.ПолучитьОбъект()
assert obj is not None, "UUID doesn't map to ERP document"
```
Test script: `_Rarzrabotki/Python/test/test_kazna_erp_uuid_mapping.py`

---

## CRITICAL: Programmatic Document Creation (Settlements / Взаиморасчёты)

### Problem
When creating documents programmatically (via external processors or form modules), settlement registers (`РасчетыСКлиентамиПоСрокам`, `РасчетыСКлиентами`) may have **Сумма = 0** or empty movements if required fields are not filled.

### Root Cause
The settlement mechanism `ВзаиморасчетыСервер.ПередЗаписью()` calculates `СуммаВзаиморасчетов` using formula: `Сумма * Курс / Кратность`. If `Курс = 0` or `Кратность = 0` (default for unfilled Number fields), the result is **0**.

### MANDATORY Fields When Creating Documents Programmatically

**For `РеализацияТоваровУслуг`, `ПриобретениеТоваровУслуг`, and similar settlement documents:**

```bsl
// === ОБЯЗАТЕЛЬНЫЕ реквизиты для взаиморасчётов ===
ДокОбъект.Курс = 1;                    // Без этого СуммаВзаиморасчетов = 0!
ДокОбъект.Кратность = 1;               // Без этого СуммаВзаиморасчетов = 0!
ДокОбъект.СуммаВзаиморасчетов = Сумма; // Сумма взаиморасчётов в шапке
ДокОбъект.ФормаОплаты = Перечисления.ФормыОплаты.Безналичная;
// Для Реализации также:
ДокОбъект.Статус = Перечисления.СтатусыРеализацийТоваровУслуг.Отгружено;
```

**В табличной части Товары:**
```bsl
СтрокаТовары.СуммаСНДС = Сумма;           // Используется для расчёта СуммаВзаиморасчетов
СтрокаТовары.СуммаВзаиморасчетов = Сумма;  // Сумма взаиморасчётов строки
```

### Settlement Architecture (НоваяАрхитектураВзаиморасчетов = TRUE)

| Concept | Details |
|---------|---------|
| **Report register** | `РасчетыСКлиентамиПоСрокам` (NOT `РасчетыСКлиентами`) |
| **Report fields** | `ДолгРегл`, `ПредоплатаРегл` (NOT `Сумма`, `КОплате`) |
| **Registrar** | `Документ.РегистраторРасчетов` (service document, NOT the source document) |
| **Entry point** | `ВзаиморасчетыСервер.ПередЗаписью()` → `ОперативныеВзаиморасчетыСервер.ЗаполнитьОперативныеВзаиморасчеты()` |
| **РасшифровкаПлатежа** | Auto-filled by `ВзаиморасчетыСервер`, but requires non-zero `СуммаДокумента` and `СуммаВзаиморасчетов` |
| **Settlement report** | `ВедомостьРасчетовСПартнерами` uses `РасчетыСКлиентамиПоСрокам` |

### Checklist for Programmatic Document Creation
- [ ] `Курс = 1` set? (CRITICAL!)
- [ ] `Кратность = 1` set? (CRITICAL!)
- [ ] `СуммаВзаиморасчетов` in header filled?
- [ ] `СуммаСНДС` in Товары rows filled?
- [ ] `СуммаВзаиморасчетов` in Товары rows filled?
- [ ] `Валюта` and `ВалютаВзаиморасчетов` set?
- [ ] `Договор` set? (needed for ОбъектРасчетов auto-creation)
- [ ] `ПорядокРасчетов` set?
- [ ] `ГруппаФинансовогоУчета` set?

---

## PROJECT-SPECIFIC NOTES

### Custom Prefix Convention
- `А_` prefix = custom/modified objects (e.g., `А_ОтражениеЗарплатыВУчете`)
- Standard BAS ERP objects have no prefix

### Integration Points
- **Kazna** integration (`_Rarzrabotki/Kazna/`)
- **BAS Accounting** exchange (`_Rarzrabotki/BASEBuh/`)
- **Data conversion** tools (`_Rarzrabotki/ConvertERP/`)

### Accounting Registers
- `Хозрасчетный` - Management accounting (main)
- `МеждународныйУчет` - IFRS accounting

---

## CHECKLIST BEFORE COMPLETION

### 1C Queries:
- [ ] `list_metadata_objects` called?
- [ ] `get_metadata_structure` called?
- [ ] Exact field names used (not invented)?
- [ ] Russian field names in queries?

### XML Processing:
- [ ] Working processor template read?
- [ ] Command has `id`?
- [ ] Action without `xsi:type`?
- [ ] All IDs unique?

### CRUD Operations:
- [ ] `find_document_ref` for UUID?
- [ ] `unpost_document` before modifying posted?
- [ ] `post_document` after update if needed?

### Git Operations:
- [ ] Branch: `claude/*` format?
- [ ] Commit message in Ukrainian/Russian?
- [ ] Push to correct branch?

---

## QUICK REFERENCE

### Query Templates

**Get document with tabular section:**
```sql
SELECT
    Doc.Ref AS DocumentRef,
    Doc.Date,
    Doc.Number,
    Doc.Контрагент,
    Товары.Номенклатура,
    Товары.Количество,
    Товары.Цена,
    Товары.Сумма
FROM Document.ПоступлениеТоваровУслуг AS Doc
LEFT JOIN Document.ПоступлениеТоваровУслуг.Товары AS Товары
ON Doc.Ref = Товары.Ref
WHERE Doc.Date BETWEEN &StartDate AND &EndDate
```

**Get register balance:**
```sql
SELECT
    Номенклатура,
    Склад,
    КоличествоBalance AS Остаток
FROM AccumulationRegister.ТоварыНаСкладах.Balance(&Date,)
```

**Get register turnover:**
```sql
SELECT
    Номенклатура,
    КоличествоTurnover AS Оборот
FROM AccumulationRegister.ТоварыНаСкладах.Turnovers(&StartDate, &EndDate,,)
```

---

## NOTEBOOKLM INTEGRATION

### Default Notebook
- **Notebook ID**: `3303acdb-2d7f-4879-9f13-78705ab3fb8c`
- **Name**: BAS ERP 2.5 INDUSTRIALBUD
- **Knowledge source**: `_Rarzrabotki/notebook/knowledge/baserp25_knowledge.md`

### Два блокноти NotebookLM (розділені за тематикою)

| Блокнот | ID | Навик | Коли питати |
|---------|-----|-------|-------------|
| **BAS ERP 2.5** (проект) | `3303acdb-2d7f-4879-9f13-78705ab3fb8c` | `/consult-notebooklm` | Архітектура конфігурації, невидимі реквізити, обмін ERP↔BuhBud↔Казна, взаєморозрахунки, зарплата, побічні ефекти ПередЗаписью/ПриЗаписи |
| **Skill 1C** (загальний) | `0e047e67-bf01-48b8-8a11-44d02b2461e8` | `/consult-notebooklm-skills` | XML-формати об'єктів 1С, як створити EPF/CFE, паттерни форм, СКД, пакетний режим, стандарти BSL, code review, оптимізація запитів, веб-тестування |

**Приклади:**
- "Які поля критично заповнити при створенні ПриобретениеТоваровУслуг?" → **BAS ERP 2.5**
- "Як устроєний XML управляемої форми? Які елементи обов'язкові?" → **Skill 1C**
- "Як працює свертка НачисленнаяЗарплатаИВзносы?" → **BAS ERP 2.5**
- "Який навик (слеш-команду) використати для створення EPF з формою?" → **Skill 1C**

### Knowledge Files
| Тема | Папка | Маніфест |
|------|-------|----------|
| Проект BAS ERP 2.5 | `_Rarzrabotki/notebook/knowledge/` | `KNOWLEDGE_MAP.md` (12 джерел) |
| Навики 1С (загальні) | `_Rarzrabotki/notebook/skill/` | 6 файлів (catalog, BSL, XML, forms, EPF/CFE, DB/web) |

Перевірка актуальності: `python _Rarzrabotki/Python/check_knowledge_freshness.py`

---

## НАВИКИ (SKILLS) — `.claude/skills/`

**73 навики** — 67 з [cc-1c-skills](https://github.com/Nikolay-Shirokov/cc-1c-skills) + 6 кастомних.

### Кастомні навики (наші)

| Навик | Тип | Для чого |
|-------|-----|----------|
| `/consult-notebooklm` | user | Консультація з NotebookLM по архітектурі BAS ERP 2.5 |
| `/consult-notebooklm-skills` | user | Консультація з NotebookLM по навикам 1С (XML, форми, EPF, СКД) |
| `/1c-bsl-review` | user | Code Review BSL по 17-пунктному чеклісту (3 рівні критичності) |
| `/1c-query-optimizer` | user | Аналіз і оптимізація запитів 1С (10 правил + антипаттерни) |
| `1c-bsl-standards` | auto | Стандарти BSL (завантажується автоматично при роботі з .bsl) |
| `1c-developer` | auto | Карта компетенцій 1С + DevOps + інструменти |

### Навики з cc-1c-skills (основні групи)

| Група | Кількість | Основні команди |
|-------|-----------|-----------------|
| Конфігурація | 4 | `/cf-info`, `/cf-init`, `/cf-edit`, `/cf-validate` |
| Метадані | 5 | `/meta-info`, `/meta-compile`, `/meta-edit`, `/meta-remove` |
| Розширення CFE | 5 | `/cfe-init`, `/cfe-borrow`, `/cfe-patch-method`, `/cfe-diff` |
| EPF/ERF | 8+3 | `/epf-init`, `/epf-build`, `/epf-dump`, `/epf-add-form` |
| Форми | 7 | `/form-info`, `/form-compile`, `/form-edit`, `/form-patterns` |
| СКД | 4 | `/skd-info`, `/skd-compile`, `/skd-edit`, `/skd-validate` |
| Бази даних | 9 | `/db-create`, `/db-run`, `/db-dump-xml`, `/db-load-git` |
| Веб | 5 | `/web-publish`, `/web-test`, `/web-info` |
| Ролі | 3 | `/role-info`, `/role-compile`, `/role-validate` |
| Підсистеми | 4+2 | `/subsystem-info`, `/subsystem-edit`, `/interface-edit` |

Повний каталог: `_Rarzrabotki/notebook/skill/skills_catalog.md`
Репозиторій: `_Rarzrabotki/cc-1c-skills/` (оновлення через `git pull`)

---

## CONTACT & SUPPORT

- **GitHub:** [Alex16111977/BASERP25_INDUSTRIALBUD](https://github.com/Alex16111977/BASERP25_INDUSTRIALBUD)
- **Organization:** LLC INDUSTRIALBUD
- **Platform:** 1C:Enterprise 8.3.20+
- **Configuration Version:** BAS ERP 2.5 (v2.13)

---

*Last updated: April 2026*
