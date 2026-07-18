# BAS ERP 2.5 INDUSTRIALBUD - Knowledge Base

## Overview
Configuration: BAS ERP 2.5 (v2.13) on 1C:Enterprise 8.3.20+
Organization: LLC INDUSTRIALBUD
GitHub: Alex16111977/BASERP25_INDUSTRIALBUD

### Configuration Scale
| Object Type | Count |
|-------------|-------|
| Documents | 400 |
| Accumulation Registers | 185 |
| Information Registers | 586 |
| Catalogs | 564 |
| Common Modules | 1770+ |

### Databases
- **ERP (BaseERP)**: Main ERP database, connection: `Srvr="localhost";Ref="BaseERP"`
- **BAS Accounting (BuhBud)**: Accounting database, connection: `Srvr="localhost";Ref="bas_industrialbud"`

---

## Golden Rules for Working with 1C

### Rule #0 - Metadata Discovery Sequence
BEFORE any query, ALWAYS follow this sequence:
1. `list_metadata_objects` - find the exact object name
2. `get_metadata_structure` - get EXACT field names, tabular sections, dimensions
3. `execute_query` - use ONLY real field names from metadata

NEVER invent field names, register names, or document names!

### Rule #1 - CRUD Operations Sequence
1. `find_document_ref` - get UUID BEFORE any modification
2. `unpost_document` - cancel posting BEFORE modification (if document is posted)
3. `update_document` - modify data
4. `post_document` - re-post after changes

### Rule #2 - Query Language
All 1C queries use Russian field names. The query language is specific to 1C platform.
Example: `SELECT Ref, Date, Number FROM Document.ПоступлениеТоваровУслуг`
Date format: ISO "2025-01-01"

### Rule #3 - Python COM Testing First
Before writing code into 1C processors/modules:
1. Create Python test in `_Rarzrabotki/Python/test/`
2. Test via COM connection (`V83.COMConnector`)
3. Only after successful test - write code into .bsl module

---

## Key Documents

### Procurement & Sales Documents
| Document Name | Description | Key Usage |
|---------------|-------------|-----------|
| `ПоступлениеТоваровУслуг` | Receipt of goods and services | Incoming goods, creates inventory movements |
| `РеализацияТоваровУслуг` | Sales of goods and services | Outgoing goods, creates revenue/COGS movements |
| `ПриобретениеТоваровУслуг` | Acquisition of goods (ERP-specific) | Similar to ПоступлениеТоваровУслуг but ERP variant |
| `ЗаказКлиента` | Customer order | Sales pipeline, reservation |
| `ЗаказПоставщику` | Purchase order | Procurement pipeline |
| `ПеремещениеТоваров` | Inventory transfer | Between warehouses |
| `СписаниеТоваров` | Inventory write-off | Damage, loss |
| `ВыпускПродукции` | Production output | Manufacturing |

### Financial Documents
| Document Name | Description |
|---------------|-------------|
| `ПриходныйКассовыйОрдер` (ПКО) | Cash receipt order |
| `РасходныйКассовыйОрдер` (РКО) | Cash disbursement order |
| `ПлатежноеПоручение` (ПП) | Payment order (bank transfer) |
| `АвансовыйОтчет` | Advance report (employee expenses) |
| `АктВыполненныхРабот` | Completed work act |

### Custom Documents (А_ prefix)
| Document Name | Description |
|---------------|-------------|
| `А_ОтражениеЗарплатыВУчете` | Salary reflection in accounting |
| `А_РаспределениеЗаработнойПлаты` | Salary distribution across departments/projects |
| `А_ТабельУчетаРабочегоВремени` | Timesheet |
| `А_АрендаТехники` | Equipment rental |
| `А_БюджетМесяц` | Monthly budget |
| `А_ОтгрузкаПродукции` | Product shipment |
| `А_ПриходДенегОтФинАгента` | Cash receipt from financial agent |

---

## Key Accumulation Registers

### Inventory Registers
| Register | Type | Key Dimensions | Resources |
|----------|------|----------------|-----------|
| `ТоварыНаСкладах` | Balance | Номенклатура, Склад | КоличествоBalance |
| `ТоварыОрганизаций` | Balance | Номенклатура, Организация | Количество |
| `СебестоимостьТоваров` | Balance | Номенклатура | Стоимость |
| `ВыпускПродукции` | Turnover | Номенклатура, Подразделение | Количество, Стоимость |

### Financial Registers
| Register | Description |
|----------|-------------|
| `ДенежныеСредстваБезналичные` | Bank funds (non-cash) |
| `ДенежныеСредстваНаличные` | Cash funds |
| `РасчетыСКлиентами` | Customer settlements |
| `РасчетыСПоставщиками` | Supplier settlements |
| `РасчетыСКлиентамиПоСрокам` | Customer settlements by terms (NEW architecture) |
| `ВыручкаИСебестоимостьПродаж` | Revenue and COGS |

### Payroll Registers
| Register | Description |
|----------|-------------|
| `ЗарплатаКВыплате` | Payroll payable |
| `ВзаиморасчетыССотрудниками` | Employee settlements |
| `РасчетыНалогоплательщиковСБюджетомПоНДФЛ` | PIT settlements |
| `А_НачисленнаяЗарплатаИВзносыПоФизлицам` | Accrued salary and contributions by individual |

---

## Key Catalogs

| Catalog | Description | Notes |
|---------|-------------|-------|
| `Номенклатура` | Products/Items | Hierarchical, has groups |
| `Контрагенты` | Counterparties | Legal entities |
| `Партнеры` | Partners | Business partners (linked to Контрагенты) |
| `Организации` | Organizations | Own legal entities |
| `Сотрудники` | Employees | Linked to ФизическиеЛица |
| `ФизическиеЛица` | Individuals | Physical persons |
| `Склады` | Warehouses | Storage locations |
| `ДоговорыКонтрагентов` | Contracts | Counterparty contracts |
| `Подразделения` | Departments | Organizational structure |
| `СтатьиРасходов` | Expense items | Cost classification |
| `СтатьиДоходов` | Income items | Revenue classification |
| `БанковскиеСчетаОрганизаций` | Bank accounts | Organization bank accounts (ERP) |
| `СтавкиНДС` | VAT rates | Catalog (NOT enum!), e.g., "Без НДС" |
| `Кассы` | Cash desks | Has ВалютаДенежныхСредств field |

---

## CRITICAL: Settlement Architecture (Взаиморасчёты)

### Problem
When creating documents programmatically, settlement registers may have Сумма = 0 or empty movements if required fields are not filled.

### Root Cause
The settlement mechanism `ВзаиморасчетыСервер.ПередЗаписью()` calculates:
```
СуммаВзаиморасчетов = Сумма * Курс / Кратность
```
If `Курс = 0` or `Кратность = 0` (default for unfilled Number fields), the result is 0.

### MANDATORY Fields for Programmatic Document Creation

For `РеализацияТоваровУслуг`, `ПриобретениеТоваровУслуг`, and similar:

**Header fields:**
```bsl
ДокОбъект.Курс = 1;                    // WITHOUT THIS: СуммаВзаиморасчетов = 0!
ДокОбъект.Кратность = 1;               // WITHOUT THIS: СуммаВзаиморасчетов = 0!
ДокОбъект.СуммаВзаиморасчетов = Сумма; // Settlement amount in header
ДокОбъект.ФормаОплаты = Перечисления.ФормыОплаты.Безналичная;
// For РеализацияТоваровУслуг also:
ДокОбъект.Статус = Перечисления.СтатусыРеализацийТоваровУслуг.Отгружено;
```

**Tabular section Товары:**
```bsl
СтрокаТовары.СуммаСНДС = Сумма;           // Used for settlement calculation
СтрокаТовары.СуммаВзаиморасчетов = Сумма;  // Row settlement amount
```

### Settlement Architecture (НоваяАрхитектураВзаиморасчетов = TRUE)

| Concept | Details |
|---------|---------|
| Report register | `РасчетыСКлиентамиПоСрокам` (NOT `РасчетыСКлиентами`!) |
| Report fields | `ДолгРегл`, `ПредоплатаРегл` (NOT `Сумма`, `КОплате`) |
| Registrar | `Документ.РегистраторРасчетов` (service document, NOT the source) |
| Entry point | `ВзаиморасчетыСервер.ПередЗаписью()` -> `ОперативныеВзаиморасчетыСервер.ЗаполнитьОперативныеВзаиморасчеты()` |
| РасшифровкаПлатежа | Auto-filled, but requires non-zero СуммаДокумента and СуммаВзаиморасчетов |
| Settlement report | `ВедомостьРасчетовСПартнерами` uses `РасчетыСКлиентамиПоСрокам` |

### Complete Checklist for Programmatic Document Creation
1. `Курс = 1` set? (CRITICAL!)
2. `Кратность = 1` set? (CRITICAL!)
3. `СуммаВзаиморасчетов` in header filled?
4. `СуммаСНДС` in Товары rows filled?
5. `СуммаВзаиморасчетов` in Товары rows filled?
6. `Валюта` and `ВалютаВзаиморасчетов` set?
7. `Договор` set? (needed for ОбъектРасчетов auto-creation)
8. `ПорядокРасчетов` set?
9. `ГруппаФинансовогоУчета` set?

---

## Critical Flags for Programmatic Document Creation

### А_ОбработанКазна flag
When creating ПКО/РКО programmatically, the `ПередЗаписью` event calls `ОбновитьПриходныйКассовыйОрдерПослеОбмена()`. If `А_ОбработанКазна = False`:
- Clears РасшифровкаПлатежа
- Overwrites ХозяйственнаяОперация (e.g., to ПрочееПоступлениеДС)
- **FIX**: Set `А_ОбработанКазна = True` to skip this processing

### А_ВведенВЕРП flag
When creating ПриобретениеТоваровУслуг programmatically, if `А_ВведенВЕРП = False` AND `А_Обработан = False`:
- Clears Товары, РасшифровкаПлатежа, ВидыЗапасов, ЭтапыГрафикаОплаты
- Rebuilds from source document data
- **FIX**: Set `А_ВведенВЕРП = True` to skip this processing

### ExternalConnection fix for А_СобытияОбъектов
Event subscriptions `А_ДоработкаДокументыПередЗаписью` and `А_ДоработкаДокументыПриЗаписи` use handler in CommonModule `А_СобытияОбъектов`.
Original: `<ExternalConnection>false</ExternalConnection>` - causes "Обработчик события не найден" in COM connections.
Fixed to: `<ExternalConnection>true</ExternalConnection>`.

---

## Accounting Register: Хозрасчетный (BAS Бухгалтерія)

### Key Facts
- The register is **correspondent** (Корреспонденция = True)
- `Количество` field does NOT exist in raw table - use `КоличествоДт` / `КоличествоКт`
- `Сумма` is NOT split (single field)
- `СчетДт`/`СчетКт`, `СубконтоДт1-3`/`СубконтоКт1-3` are split by side

### CRITICAL: СубконтоДт1/СубконтоКт1 NOT available in direct query
- Direct query `SELECT Д.СубконтоДт1 FROM РегистрБухгалтерии.Хозрасчетный` gives ERROR
- **Workaround**: Use virtual table `ОстаткиИОбороты` which returns `Субконто1` as a column
- Or iterate COM table individually

### Virtual Table ОстаткиИОбороты Parameters
Accepts 7 parameters: (НачалоПериода, КонецПериода, Периодичность, , УсловиеСчета, , )
Example: `ОстаткиИОбороты(&НачалоПериода, &КонецПериода, , , Счет В ИЕРАРХИИ(...), , )`

### В ИЕРАРХИИ inside aggregate functions - NOT ALLOWED
`В ИЕРАРХИИ` cannot be used inside `СУММА(ВЫБОР КОГДА ... В ИЕРАРХИИ ... ТОГДА ...)`.
**Workaround**: Use UNION ALL with `В ИЕРАРХИИ` in WHERE clause of each subquery.

### UNION ALL Pattern for Товары Analysis
- Дт part: `Д.КоличествоДт КАК Количество` (receipt/debit = positive)
- Кт part: `-Д.КоличествоКт` (expense/credit = negative)
- HAVING: Always use `СУММА(Т.Сумма) <> 0` (NOT Количество!) - otherwise docs with monetary-only movements (МБП) get filtered out
- Add `Д.Сумма КАК Сумма` / `-Д.Сумма` to inner queries for HAVING

---

## Cross-Database Operations (ERP <-> BAS Бухгалтерія)

### NEVER serialize COM tables across databases
BuhBud uses `Справочник.БанковскиеСчета`, ERP uses `Справочник.БанковскиеСчетаОрганизаций`.
After `ЗначениеИзСтрокиВнутр` in ERP, references become invalid.
**FIX**: Iterate directly over BuhBud COM table.

### Cross-base UUID Lookup (BuhKazn/BuhBud → ERP via COM)

**Architecture**: Правила обміну (`_Rarzrabotki/ConvertERP/`, `_Rarzrabotki/ExChange/Казна/`) зберігають UUID документа **1:1** між базами. Регістр БСП `РегистрСведений.СоответствияОбъектовИнформационныхБаз` також тримає маппінг (вузли `ПланОбмена.КазначействоBASERP`, `ПланОбмена.ОбменУправлениеПредприятиемБухгалтерия20`). Прямий пошук за UUID — єдиний архітектурно правильний шлях, а не Номер+Дата.

**Підтверджено правилами обміну** (приклади для Казни):
| Казна (BuhKazn) | ERP (BaseERP) |
|---|---|
| `РаспределениеЗаработнойПлаты` | `А_РаспределениеЗаработнойПлаты` |
| `РаспределениеФ2` | `РаспределениеФ2` |
| `ПриходныйКассовыйОрдер` | `ПриходныйКассовыйОрдер` (або `А_ПриходДенегОтФинАгента` якщо `_ФинАгент`) |
| `ПлатежноеПоручениеВходящее` | `ПоступлениеБезналичныхДенежныхСредств` |

#### КОРЕКТНИЙ патерн (перевірено на `Documents/А_ОтражениеЗПпоКазне`)

```bsl
// V83 — відкрите COM-зʼєднання до бази-джерела (Казни/BuhBud)
// СтрТЗ.Регистратор — COM-reference на документ у базі-джерелі

// КРОК 1: Витягнути UUID як рядок через ВІДДАЛЕНИЙ контекст (V83)
// КРИТИЧНО: V83.string(...) — англійське lowercase, єдиний надійний варіант
УИДСтр = V83.string(СтрТЗ.Регистратор.УникальныйИдентификатор());

// КРОК 2: Створити ЛОКАЛЬНИЙ (ERP) УникальныйИдентификатор + знайти документ
УИД = Новый УникальныйИдентификатор(УИДСтр);
ДокСсылка = Документы.А_РаспределениеЗаработнойПлаты.ПолучитьСсылку(УИД);

// КРОК 3: Перевірити, що документ фізично існує (захист від битого посилання)
Если ДокСсылка.ПолучитьОбъект() <> Неопределено Тогда
    // Документ знайдено — використовуйте ДокСсылка
КонецЕсли;
```

#### Таблиця відмов (що НЕ працює)

| Виклик | Результат | Причина |
|---|---|---|
| `Строка(V83_UUID)` | `'None'` / порожньо | ERP-local `Строка()` не вміє серіалізувати чужий COM-UUID |
| `XMLСтрока(V83_UUID)` | **EXCEPTION** `Ошибка преобразования данных XML` | ERP-local XML-серіалізатор відхиляє foreign COM-object |
| `V83.XMLСтрока(V83_UUID)` | ненадійно через COM dispatch | Cyrillic method names через V83.COMConnector proxy не завжди резолвляться |
| `V83.String(V83_UUID)` | ненадійно в BSL | BSL COM-proxy віддає перевагу lowercase |
| `V83.string(V83_UUID)` | **ПРАЦЮЄ** | стандартний патерн у робочих обробках (`Обработка_ПлатежВЕРП` та ін.) |

Python COM validation (усе в **kazna**-контексті працює, в **erp**-контексті падає):
```
kazna.String(uuid_obj)     → 'd4d9b293-dfcc-11f0-8104-00155dce3d04' ✓
erp.XMLСтрока(uuid_obj)    → EXCEPTION ✗
erp.String(uuid_obj)       → 'None' ✗
```

#### CRITICAL: `Ссылка` — зарезервоване ім'я змінної

**НЕ МОЖНА** оголошувати локальну змінну з ім'ям `Ссылка` в BSL — конфліктує з системною властивістю `.Ссылка` на об'єктах-посиланнях. Всередині `Попытка/Исключение` це призводить до мовчазного виключення: функція повертає `Неопределено` без жодної помилки, і неможливо зрозуміти причину.

```bsl
// ❌ WRONG — мовчки падає
Ссылка = Документы.X.ПолучитьСсылку(УИД);
// ✅ RIGHT
ДокСсылка = Документы.X.ПолучитьСсылку(УИД);
```

Інші небезпечні імена локальних змінних у модулях об'єктів: `Дата`, `Номер`, `Организация`, `ПометкаУдаления`, `Проведен` (усе — реквізити/властивості документа).

#### Альтернатива через `СоответствияОбъектовИнформационныхБаз`

```1c
ВЫБРАТЬ СО.УникальныйИдентификаторИсточника КАК СсылкаERP
ИЗ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК СО
ГДЕ СО.УникальныйИдентификаторПриемника = &UUIDКазна
    И ТИПЗНАЧЕНИЯ(СО.УникальныйИдентификаторИсточника) = ТИП(Документ.А_РаспределениеЗаработнойПлаты)
```

#### Довідкова реалізація

`Documents/А_ОтражениеЗПпоКазне/Ext/ObjectModule.bsl`:
- `ЗагрузитьРаспределениеКазна_ОтражениеЗП()` — COM-обхід BuhKazn через БДДС, витягнення UUID через `V83.string()`
- `А_НайтиДокументПоУИД_ОтражениеЗП(УИДСтрока, ВидДокумента)` — локальний ERP lookup через `ПолучитьСсылку + ПолучитьОбъект`

Python тести для валідації: `test_kazna_erp_uuid_mapping.py`, `test_xmlstr_context.py`, `test_lookup_single_uuid.py`.

### XMLСтрока(ВидДвижения) returns English
For accumulation registers: returns "Receipt"/"Expense" (English), NOT "Приход"/"Расход".

### BAS Бухгалтерія uses different document names
BuhBud: `ПоступлениеТоваровУслуг` (NOT `ПриобретениеТоваровУслуг` like ERP)

---

## Query Templates

### Document with Tabular Section
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

### Register Balance (Virtual Table)
```sql
SELECT
    Номенклатура,
    Склад,
    КоличествоBalance AS Остаток
FROM AccumulationRegister.ТоварыНаСкладах.Balance(&Date,)
```

### Register Turnover (Virtual Table)
```sql
SELECT
    Номенклатура,
    КоличествоTurnover AS Оборот
FROM AccumulationRegister.ТоварыНаСкладах.Turnovers(&StartDate, &EndDate,,)
```

### Register Movements by Document
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

---

## External Data Processors Structure

### File Layout
```
ProcessorName.xml                    <- Main descriptor
ProcessorName/
  ├── Ext/ObjectModule.bsl          <- Business logic
  └── Forms/FormName/
        ├── FormName.xml            <- Form descriptor
        └── Ext/Form.xml + Form/Module.bsl  <- Form layout + code
```

### XML Rules (CRITICAL!)
- Command element MUST have `id="number"` attribute
- Action element WITHOUT `xsi:type` attribute
- Button: `Type=CommandBarButton`
- All IDs must be unique within the form

### Available Custom Processors
| Processor | Purpose |
|-----------|---------|
| `СинхронизироватьТовары` | Sync inventory between ERP and BAS Accounting |
| `ЗаполнениеНачисленийЗП` | Fill payroll accruals |
| `Загрузить зарплату из казны` | Import salary from Kazna |
| `Перенос остатков номенклатуры` | Transfer inventory balances |
| `Перенос остатков денег` | Transfer cash balances |
| `А_ОбновитьСтатьюДенегВЗатратах` | Update cash flow article in expenses |
| `РегистрацияИзмененийДляОбменаДанными` | Register changes for data exchange |

---

## Integration Points

### Kazna (Treasury/Payroll System)
- Location: `_Rarzrabotki/Kazna/`
- Used for salary import and treasury operations
- Custom document: `А_ПриходДенегОтФинАгента`

### BAS Accounting (BuhBud)
- Location: `_Rarzrabotki/BASEBuh/`
- Data exchange rules: `_Rarzrabotki/ExChange/ПравилаОбменаДаннымиЕРПБух/`
- Sync processor: `СинхронизироватьТовары` (inventory, money, settlements)

### Data Conversion
- Location: `_Rarzrabotki/ConvertERP/`
- Migration tools for initial data import

---

## Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| Field not found | Invented field name | Use `get_metadata_structure` first |
| Query returns empty | Wrong date format | Use ISO format: "2025-01-01" |
| UUID not found | Used execute_query for UUID | Use `find_document_ref` instead |
| Cannot modify posted doc | Document is posted | Use `unpost_document` first |
| Command without id | Missing id in XML | Add `id="N"` to Command element |
| xsi:type in Action | Extra attribute in XML | Remove `xsi:type` from Action |
| СуммаВзаиморасчетов = 0 | Курс or Кратность = 0 | Set Курс=1, Кратность=1 |
| Обработчик не найден (COM) | ExternalConnection=false | Set ExternalConnection=true in CommonModule |
| COM Количество error | Name collision in Python | Use alias КолВо in queries |
| СубконтоДт1 not found | Not available in raw query | Use ОстаткиИОбороты virtual table |

### What NEVER Works
- `execute_query` to get document UUID for modifications (use `find_document_ref`)
- `update_document(post=false)` to unpost (use `unpost_document`)
- Inventing field names without checking metadata
- Modifying posted document without unposting first
- ЗначениеВСтрокуВнутр/ЗначениеИзСтрокиВнутр across different databases
- В ИЕРАРХИИ inside aggregate functions (СУММА, КОЛИЧЕСТВО)

---

## Custom Prefix Convention

- **`А_` prefix** = custom/modified objects created specifically for INDUSTRIALBUD
- Standard BAS ERP 2.5 objects have no prefix
- Examples: `А_ОтражениеЗарплатыВУчете`, `А_НаправлениеДеятельности`, `А_СобытияОбъектов`
- Custom field `А_НаправлениеДеятельности` in documents (note: custom prefix on the field too!)
- `СтавкиНДС` is a Catalog (Справочник), NOT an Enum (Перечисление)

---

## MCP Limitations

### Posting via MCP
`ПриобретениеТоваровУслуг` posting via MCP fails with "Не задано значение параметра РаздельныйУчетПостатейныхПроизводственныхЗатратПоНалогообложениюНДС". This is a session parameter not initialized in MCP HTTP service context. Works fine inside 1C Application Server (managed forms).

### Python COM Testing
- `result.Количество` may be interpreted as a method (COM name collision) - use alias `КолВо` in test queries
- BAS Бухгалтерія connection: `Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"`
- ERP connection: `Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"`

---

## СинхронизироватьТовары Processor - Key Details

- Location: `_Rarzrabotki/Обработки/СинхронизироватьТовары/Ext/ObjectModule.bsl`
- Key procedures: `АнализДокументов_Товары` (line 1082), `АнализДокументовBuhBud_Товары` (line 1127), `ОбробитиДокументиЕРП` (line 1512), `ПроверитьДокументВBuhBud` (line 1592)
- Money section: skip transfer operations, UNION ALL query for КоличествоБух, amount comparison
- BuhBud reverse check: `ПеревіритиДокументиBuhBud_Деньги()` procedure finds documents existing ONLY in BuhBud
- Edit tool struggles with file encoding (BOM UTF-8 + tabs) - use Python scripts for complex replacements

---

## Accounting Registers

| Register | Description | Type |
|----------|-------------|------|
| `Хозрасчетный` | Management accounting (main) | Correspondent |
| `МеждународныйУчет` | IFRS accounting | - |

The `Хозрасчетный` register is the primary accounting register used in BAS Бухгалтерія for all financial postings.
