# CLAUDE.md - AI ASSISTANT INSTRUCTION: BAS ERP 2.5 INDUSTRIALBUD

## CRITICAL RULES (READ AT THE START OF EVERY SESSION!)

### Rule #0 - Golden Sequence for Metadata
BEFORE ANY 1C query:
1. `list_metadata_objects` -> find objects
2. `get_metadata_structure` -> get EXACT structure
3. `execute_query` -> use ONLY real field names

**NEVER invent field names, registers, documents!**

### Rule #1 - CRUD Operations
- `find_document_ref` -> get UUID BEFORE modification
- `unpost_document` -> cancel posting BEFORE modification
- `update_document` -> modify data
- `post_document` -> post after changes

### Rule #2 - Step-by-Step Development
Break the task into steps. Test after each step!

### Rule #3 - Query Language
All 1C queries use Russian field names. Example:
```
SELECT Ref, Date, Number FROM Document.ПоступлениеТоваровУслуг
```

### Rule #4 - Python COM Testing First
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

## CONTACT & SUPPORT

- **GitHub:** [Alex16111977/BASERP25_INDUSTRIALBUD](https://github.com/Alex16111977/BASERP25_INDUSTRIALBUD)
- **Organization:** LLC INDUSTRIALBUD
- **Platform:** 1C:Enterprise 8.3.20+
- **Configuration Version:** BAS ERP 2.5 (v2.13)

---

*Last updated: January 2026*
