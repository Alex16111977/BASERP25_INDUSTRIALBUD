# Внешние обработки (EPF/ERF), Расширения конфигурации (CFE) и Пакетный режим конфигуратора

> Полные гайды по созданию/сборке EPF/ERF, работе с расширениями CFE и пакетному режиму 1С.
> Источник: cc-1c-skills (Nikolay-Shirokov/cc-1c-skills)

---

## Часть 1: Внешние обработки (EPF) — Гайд

# Внешние обработки и отчёты (EPF / ERF)

Навыки для создания, модификации и сборки внешних обработок (`.epf`) и внешних отчётов (`.erf`) 1С:Предприятия 8.3 из XML-исходников.

## Навыки обработок (EPF)

| Навык | Параметры | Описание |
|-------|-----------|----------|
| `/epf-init` | `<Name> [Synonym]` | Создать новую обработку (корневой XML + модуль объекта) |
| `/epf-add-form` | `<ProcessorName> <FormName> [Synonym]` | Добавить управляемую форму к обработке |
| `/epf-build` | `<ProcessorName>` | Собрать EPF из XML (через 1cv8.exe) |
| `/epf-dump` | `<EpfFile>` | Разобрать EPF в XML (через 1cv8.exe) |
| `/epf-bsp-init` | `<ProcessorName> <Вид>` | Добавить регистрацию БСП (СведенияОВнешнейОбработке) |
| `/epf-bsp-add-command` | `<ProcessorName> <Идентификатор>` | Добавить команду в дополнительную обработку БСП |
| `/epf-validate` | `<ObjectPath> [-MaxErrors 30]` | Валидация структурной корректности обработки (10 проверок) |

## Внешние отчёты (ERF)

| Навык | Параметры | Описание |
|-------|-----------|----------|
| `/erf-init` | `<ReportName> [Synonym] [--WithSKD]` | Создать новый отчёт (корневой XML + модуль объекта + опционально СКД) |
| `/erf-build` | `<ReportName>` | Собрать ERF из XML (через 1cv8.exe) |
| `/erf-dump` | `<ErfFile>` | Разобрать ERF в XML (через 1cv8.exe) |
| `/erf-validate` | `<ObjectPath> [-MaxErrors 30]` | Валидация структурной корректности отчёта (10 проверок) |

Флаг `--WithSKD` создаёт макет `ОсновнаяСхемаКомпоновкиДанных` и привязывает его к `MainDataCompositionSchema`.

## Универсальные навыки

Работают с любыми объектами — обработками, отчётами, справочниками, документами и др.

| Навык | Параметры | Описание |
|-------|-----------|----------|
| `/template-add` | `<ObjectName> <TemplateName> <TemplateType>` | Добавить макет (HTML, Text, SpreadsheetDocument, BinaryData, DataCompositionSchema) |
| `/template-remove` | `<ObjectName> <TemplateName>` | Удалить макет |
| `/help-add` | `<ObjectName>` | Добавить встроенную справку (Help.xml + HTML) |
| `/form-remove` | `<ObjectName> <FormName>` | Удалить форму |

Для отчётов: при добавлении макета типа DataCompositionSchema автоматически заполняется `MainDataCompositionSchema` (если пуст).

Навыки удаления (`template-remove`, `form-remove`) не вызываются Claude автоматически — только по явной команде пользователя.

## Сценарии использования

Не обязательно запоминать команды и параметры. Просто опишите задачу своими словами — Claude сам подберёт нужные навыки.

### Обработка с формой

Типичная обработка для манипуляций с данными — форма для пользователя, модуль объекта для логики.

```
> Создай обработку ЗагрузкаПрайса с формой
```

Claude выполнит `/epf-init` и `/epf-add-form` с правильными параметрами.

### Внешняя печатная форма

Обработка с макетом табличного документа, подключаемая через механизм дополнительных обработок БСП к конкретному документу.

```
> Создай внешнюю печатную форму для документа Реализация. Макет — табличный документ.
```

Claude создаст обработку, добавит макет SpreadsheetDocument, вызовет `/epf-bsp-init` с видом ПечатнаяФорма и назначением, сгенерирует `СведенияОВнешнейОбработке()` и процедуру `Печать()`.

### Внешний отчёт с СКД

```
> Создай внешний отчёт ОстаткиНаСкладе с СКД
```

Claude выполнит `/erf-init ОстаткиНаСкладе --WithSKD`, затем предложит заполнить схему компоновки через `/skd-compile`.

### Доработка существующей обработки

```
> Добавь справку с описанием как пользоваться обработкой
```

Claude вызовет `/help-add` и предложит отредактировать HTML.

```
> Добавь ещё одну команду печати — накладная
```

Claude вызовет `/epf-bsp-add-command`, добавит команду в `СведенияОВнешнейОбработке()` и блок обработки в процедуру `Печать()`.

```
> Собери
```

Claude вызовет `/epf-build` или `/erf-build` в зависимости от типа объекта.

### Примеры слеш-команд

Слеш-команды работают для случаев, когда хочется точного контроля:

```
> /epf-init МояОбработка "Моя обработка"
> /epf-add-form МояОбработка Форма
> /template-add МояОбработка Макет HTML
> /help-add МояОбработка
> /epf-build МояОбработка

> /erf-init МойОтчёт --WithSKD
> /template-add МойОтчёт ДопМакет SpreadsheetDocument
> /erf-build МойОтчёт
```

## Структура каталогов

После `/epf-init` создаётся структура:

```
src/
├── МояОбработка.xml                          # Корневой файл метаданных
└── МояОбработка/
    └── Ext/
        └── ObjectModule.bsl                  # Модуль объекта
```

После `/epf-add-form` и `/template-add`:

```
src/
├── МояОбработка.xml
└── МояОбработка/
    ├── Ext/
    │   └── ObjectModule.bsl
    ├── Forms/
    │   ├── Форма.xml                         # Метаданные формы
    │   └── Форма/
    │       └── Ext/
    │           ├── Form.xml                  # Описание формы
    │           └── Form/
    │               └── Module.bsl            # Модуль формы
    └── Templates/
        ├── Макет.xml                         # Метаданные макета
        └── Макет/
            └── Ext/
                └── Template.html             # Содержимое макета
```

После `/erf-init МойОтчёт --WithSKD`:

```
src/
├── МойОтчёт.xml                              # Корневой файл (ExternalReport)
└── МойОтчёт/
    ├── Ext/
    │   └── ObjectModule.bsl                  # Модуль объекта
    └── Templates/
        ├── ОсновнаяСхемаКомпоновкиДанных.xml
        └── ОсновнаяСхемаКомпоновкиДанных/
            └── Ext/
                └── Template.xml              # Пустая СКД
```

Первая добавленная форма автоматически становится основной (DefaultForm). Флаг `--main` нужен только для переназначения основной формы на другую.

## Сборка и разборка

### Сборка (`/epf-build`, `/erf-build`)

Если база не указана — автоматически создаётся временная база с заглушками метаданных для ссылочных типов. Явная база не обязательна, но предпочтительна — обеспечивает точное соответствие типов.

**Ограничение**: если на форме обработки выведены наборы записей регистров, в XML-исходниках нет информации о том, чем является поле регистра — измерением, ресурсом или реквизитом. Скрипт пытается угадать категорию по имени, но при ошибке 1С молча сбрасывает привязки колонок (DataPath). В таких случаях лучше использовать реальную базу с нужной конфигурацией.

### Разборка (`/epf-dump`, `/erf-dump`)

База с конфигурацией **обязательна**. Dump в пустой базе безвозвратно теряет ссылочные типы (`CatalogRef.XXX` → `xs:string`).

## Технические детали

- Все XML-файлы создаются в **UTF-8 с BOM** (как в реальных выгрузках 1С)
- PowerShell-скрипты используют `System.Xml.XmlDocument` для модификации корневого XML
- UUID генерируются через `[guid]::NewGuid()`
- ClassId обработки: `c3831ec8-d8d5-4f93-8a22-f9bfae07327f`
- ClassId отчёта: `e41aff26-25cf-4bb6-b6c1-3f478a75f374`
- Порядок элементов в `ChildObjects`: TabularSections → Forms → Templates
- Первая форма автоматически назначается основной (DefaultForm)
- Навыки БСП (`epf-bsp-*`) не используют скрипты — Claude модифицирует код напрямую через Read/Edit
- Для отчётов: `/template-add` с типом DataCompositionSchema автоматически заполняет `MainDataCompositionSchema`

## Спецификации

- [XML-формат выгрузки обработок](1c-epf-spec.md) — структура XML-файлов, namespace, элементы форм
- [XML-формат внешних отчётов](1c-erf-spec.md) — отличия ERF от EPF, Properties, MainDataCompositionSchema
- [Встроенная справка](1c-help-spec.md) — Help.xml, HTML-страницы, кнопка справки на форме
- [Сборка и разборка EPF/ERF](build-spec.md) — команды `1cv8.exe`, параметры, коды возврата


---

## Часть 2: Расширения конфигурации (CFE) — Гайд

# Расширения конфигурации (CFE)

Навыки группы `/cfe-*` позволяют создавать, заимствовать объекты, перехватывать методы, проверять и анализировать расширения конфигурации 1С.

## Навыки

| Навык | Параметры | Описание |
|-------|-----------|----------|
| `/cfe-init` | `<Name> [-Purpose Patch\|Customization\|AddOn] [-CompatibilityMode]` | Создание расширения (scaffold XML-исходников) |
| `/cfe-borrow` | `-ExtensionPath <path> -ConfigPath <path> -Object "Type.Name" [-BorrowMainAttribute]` | Заимствование объектов из конфигурации |
| `/cfe-patch-method` | `-ExtensionPath <path> -ModulePath "Type.Name.Module" -MethodName "X" -InterceptorType Before` | Генерация перехватчика метода |
| `/cfe-validate` | `<ExtensionPath> [-MaxErrors 30]` | Валидация структурной корректности (9 проверок) |
| `/cfe-diff` | `-ExtensionPath <path> -ConfigPath <path> [-Mode A\|B]` | Анализ расширения и проверка переноса |

## Рабочий цикл

```
cf-info (версия, совместимость)
    ↓
/cfe-init → scaffold расширения
    ↓
/cfe-borrow → заимствование объектов из конфигурации
    ↓
/cfe-patch-method → перехват методов
    ↓
/cfe-validate → проверка корректности
    ↓
/cfe-diff Mode A → обзор изменений
```

## Типичные сценарии

### Создание расширения для исправления бага

```
> Создай расширение для исправления бага в справочнике Контрагенты,
  конфигурация ERP в C:\cfsrc\erp
```

Claude выполнит:
1. `/cf-info C:\cfsrc\erp -Mode brief` — получить версию и режим совместимости
2. `/cfe-init` — создать расширение с нужным `CompatibilityMode`
3. `/cfe-borrow` — заимствовать `Catalog.Контрагенты`
4. `/cfe-patch-method` — создать перехватчик нужного метода
5. `/cfe-validate` — проверить результат

### Добавление реквизита в объект и вывод на форму

```
> Добавь реквизит "ОсновнойПоставщик" (тип СправочникСсылка.Партнеры)
  в справочник Номенклатура и выведи на форму элемента.
  Конфигурация ERP в C:\cfsrc\erp
```

Claude выполнит:
1. `/cfe-init` — создать расширение
2. `/cfe-borrow -Object "Catalog.Номенклатура.Form.ФормаЭлемента" -BorrowMainAttribute` — заимствовать форму с реквизитами объекта
3. `/meta-edit` — добавить новый реквизит `Расш1_ОсновнойПоставщик` в Номенклатура
4. `/form-edit` — вывести реквизит на форму
5. `/cfe-validate` — проверить результат

### Анализ существующего расширения

```
> Покажи что изменено в расширении src/
```

Claude вызовет `/cfe-diff -Mode A` и покажет: заимствованные объекты, перехватчики, собственные объекты.

### Проверка переноса изменений

```
> Проверь, все ли изменения из расширения перенесены в конфигурацию
```

Claude вызовет `/cfe-diff -Mode B` — найдёт блоки `#Вставка` и проверит их наличие в конфигурации.

## cfe-init — создание расширения

Параметры:

| Параметр | Описание | По умолчанию |
|----------|----------|--------------|
| `Name` | Имя расширения (обязат.) | — |
| `Synonym` | Синоним | = Name |
| `NamePrefix` | Префикс собственных объектов | = Name + "_" |
| `OutputDir` | Каталог | `src` |
| `Purpose` | Назначение | `Customization` |
| `Version` | Версия | — |
| `Vendor` | Поставщик | — |
| `CompatibilityMode` | Режим совместимости | `Version8_3_24` |
| `NoRole` | Без основной роли | false |

Создаёт:
```
<OutputDir>/
├── Configuration.xml         # Свойства расширения
├── Languages/
│   └── Русский.xml           # Язык (Adopted)
└── Roles/                    # Если не -NoRole
    └── <Prefix>ОсновнаяРоль.xml
```

Назначение расширения (`Purpose`):
- `Patch` — исправление ошибок (минимальные изменения, только перехватчики)
- `Customization` — доработка (реквизиты, формы, модули)
- `AddOn` — дополнение (полноценный функционал)

## cfe-borrow — заимствование объектов

Заимствует объекты из основной конфигурации в расширение. Создаёт минимальные XML-файлы с `ObjectBelonging=Adopted` и `ExtendedConfigurationObject`.

Формат `-Object`:
- `Catalog.Контрагенты` — справочник
- `CommonModule.РаботаСФайлами` — общий модуль
- `Enum.ВидыОплат` — перечисление
- `Document.Заказ ;; Catalog.Товары` — несколько объектов через `;;`

Поддерживаемые типы: Catalog, Document, Enum, CommonModule, Report, DataProcessor, ExchangePlan, InformationRegister, AccumulationRegister, AccountingRegister, CalculationRegister, ChartOfAccounts, ChartOfCharacteristicTypes, ChartOfCalculationTypes, BusinessProcess, Task, и другие (44 типа).

### Заимствование формы с реквизитами объекта (-BorrowMainAttribute)

При добавлении нового реквизита на заимствованную форму нужна опция `-BorrowMainAttribute`:
- Без неё форма заимствуется "пустой" — только визуальные элементы, без привязки к данным
- С ней — форма сохраняет `DataPath` привязки к реквизитам объекта (`Объект.XXX`)

Два режима:
- `Form` (по умолчанию) — заимствует только реквизиты, выведенные на форму
- `All` — заимствует все реквизиты объекта (включая не выведенные на форму)

Каскадно заимствует зависимые объекты по типам реквизитов (справочники, перечисления, определяемые типы) как оболочки. Если зависимый объект уже заимствован с содержимым — не перезаписывает его.

## cfe-patch-method — перехват методов

Генерирует `.bsl` файл с декоратором перехвата для заимствованного объекта.

Параметры:

| Параметр | Описание |
|----------|----------|
| `ModulePath` | `Catalog.X.ObjectModule`, `CommonModule.Y`, `Catalog.X.Form.Z` |
| `MethodName` | Имя перехватываемого метода |
| `InterceptorType` | `Before` / `After` / `ModificationAndControl` |
| `Context` | `НаСервере` / `НаКлиенте` / `НаСервереБезКонтекста` |
| `IsFunction` | Добавить `Возврат` |

Типы перехватчиков:

| Тип | Декоратор | Когда использовать |
|-----|-----------|-------------------|
| `Before` | `&Перед` | Выполнить код до вызова оригинального метода |
| `After` | `&После` | Выполнить код после вызова оригинального метода |
| `ModificationAndControl` | `&ИзменениеИКонтроль` | Полная замена тела метода с маркерами `#Вставка`/`#Удаление` |

Пример генерируемого кода (`Before`):
```bsl
&НаСервере
&Перед("ПриЗаписи")
Процедура Расш1_ПриЗаписи()
	// TODO: код перед вызовом оригинального метода
КонецПроцедуры
```

## cfe-validate — проверки

| # | Проверка | Уровень |
|---|----------|---------|
| 1 | XML well-formedness, MetaDataObject/Configuration, version | ERROR |
| 2 | InternalInfo: 7 ContainedObject, валидные ClassId | ERROR |
| 3 | Extension properties: ObjectBelonging=Adopted, Name, Purpose, NamePrefix, KeepMapping | ERROR |
| 4 | Enum-значения (4 свойства) | ERROR |
| 5 | ChildObjects: валидные типы, нет дубликатов, порядок | ERROR/WARN |
| 6 | DefaultLanguage ссылается на существующий Language | ERROR |
| 7 | Файлы языков существуют | WARN |
| 8 | Каталоги объектов существуют | WARN |
| 9 | Заимствованные объекты: ObjectBelonging=Adopted, ExtendedConfigurationObject UUID | ERROR/WARN |

## cfe-diff — режимы

### Mode A — обзор расширения

Для каждого объекта показывает:
- `[BORROWED]` — заимствованный: перехватчики, собственные реквизиты/формы
- `[OWN]` — собственный: количество реквизитов, ТЧ, форм

### Mode B — проверка переноса

Для каждого `&ИзменениеИКонтроль` проверяет, перенесены ли блоки `#Вставка` в конфигурацию:
- `[TRANSFERRED]` — код найден в конфигурации
- `[NOT_TRANSFERRED]` — код не найден
- `[NEEDS_REVIEW]` — нет блоков `#Вставка` или модуль конфигурации не найден

## Связь с другими навыками

- `/cf-info` — получение версии и совместимости конфигурации перед `cfe-init`
- `/meta-compile` — создание собственных объектов расширения (реквизиты, ТЧ)
- `/form-compile`, `/form-edit` — создание и модификация форм расширения
- `/cfe-validate` — всегда проверяйте расширение после изменений

## Спецификации

- [1c-extension-spec.md](1c-extension-spec.md) — XML-формат выгрузки расширений конфигурации (CFE)


---

## Часть 3: Пакетный режим конфигуратора

# Пакетный режим конфигуратора 1С

## Общие сведения

Конфигуратор 1С:Предприятия 8.3 поддерживает пакетный (безоконный) режим для автоматизации операций с конфигурациями, информационными базами и внешними обработками. Все операции выполняются через командную строку `1cv8.exe`.

**Два режима запуска:**

| Режим | Назначение |
|-------|-----------|
| `DESIGNER` | Конфигуратор — работа с конфигурацией, сборка EPF, обновление БД |
| `ENTERPRISE` | Предприятие — запуск обработок, навигация по ссылкам |
| `CREATEINFOBASE` | Создание новой информационной базы |

**Путь к 1cv8.exe** зависит от версии платформы: `C:\Program Files\1cv8\8.3.27.1859\bin\1cv8.exe`.

## Подключение к информационной базе

| Параметр | Описание |
|----------|----------|
| `/F <каталог>` | Файловая база — каталог с файлом `1Cv8.1CD` |
| `/S <адрес>` | Серверная база — формат `server/ibname` |
| `/IBName <имя>` | По имени из списка баз (в кавычках если содержит пробелы) |
| `/IBConnectionString` | Полная строка соединения |

Примеры:
```
1cv8.exe DESIGNER /F "C:\Bases\MyBase" ...
1cv8.exe DESIGNER /S server-pc/accounting ...
1cv8.exe DESIGNER /IBName "Бухгалтерия предприятия" ...
```

### Аутентификация

| Параметр | Описание |
|----------|----------|
| `/N<имя>` | Имя пользователя (**без пробела** после `/N`) |
| `/P<пароль>` | Пароль (**без пробела** после `/P`). Можно опустить если пароля нет |
| `/WA-` | Запретить аутентификацию ОС |
| `/WA+` | Обязательная аутентификация ОС (по умолчанию) |

> **Важно**: между `/N` и именем, а также между `/P` и паролем пробела нет: `/NАдмин /PSecret123`.

## Общие параметры пакетного режима

| Параметр | Описание |
|----------|----------|
| `/DisableStartupDialogs` | Подавляет интерактивные диалоги. **Обязательно** для пакетного режима — без него конфигуратор может зависнуть в ожидании ввода |
| `/DisableStartupMessages` | Подавляет стартовые предупреждения (несоответствие конфигурации БД и т.п.) |
| `/Out <файл> [-NoTruncate]` | Файл для вывода служебных сообщений (UTF-8). `-NoTruncate` — не очищать файл перед записью |
| `/DumpResult <файл>` | Записать числовой код результата в файл (0 — успех, 1 — ошибка, 101 — ошибки проверки) |
| `/Visible` | Показать окно конфигуратора (по умолчанию скрыто в пакетном режиме) |

## Создание информационной базы

```
1cv8.exe CREATEINFOBASE <строка_соединения> [/AddToList [<имя>]] [/UseTemplate <файл>] [/DumpResult <файл>]
```

### Файловая база

```
1cv8.exe CREATEINFOBASE File="C:\Bases\EmptyDB"
```

### Серверная база

```
1cv8.exe CREATEINFOBASE Srvr="server-pc";Ref="new_db"
```

### Параметры

| Параметр | Описание |
|----------|----------|
| `File="<путь>"` | Строка соединения для файловой базы |
| `Srvr="<сервер>";Ref="<имя>"` | Строка соединения для серверной базы |
| `/AddToList [<имя>]` | Добавить в список баз. Имя — необязательно |
| `/UseTemplate <файл>` | Создать по шаблону (.cf или .dt) |
| `/DumpResult <файл>` | Записать результат (0 — успех) |

## Работа с конфигурацией — бинарные файлы (CF)

### Выгрузка конфигурации в CF-файл

```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpCfg config.cf /Out log.txt
```

**`/DumpCfg <файл> [-Extension <имя>]`** — сохранить конфигурацию в .cf-файл.

### Загрузка конфигурации из CF-файла

```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /LoadCfg config.cf /Out log.txt
```

**`/LoadCfg <файл> [-Extension <имя>] [-AllExtensions]`** — загрузить конфигурацию из .cf-файла.

| Параметр | Описание |
|----------|----------|
| `-Extension <имя>` | Работа с расширением (указать имя) |
| `-AllExtensions` | Работа со всеми расширениями (файл — архив расширений) |

> После `/LoadCfg` конфигурация загружается в «основную» конфигурацию конфигуратора. Для применения к БД необходим `/UpdateDBCfg`.

## Работа с конфигурацией — XML-исходники

### Выгрузка `/DumpConfigToFiles`

```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpConfigToFiles <каталог> [параметры] /Out log.txt
```

Полная сигнатура:
```
/DumpConfigToFiles <каталог> [-Extension <имя>] [-AllExtensions]
    [-update] [-force] [-getChanges <файл>]
    [-configDumpInfoForChanges <файл>] [-listFile <файл>]
    [-configDumpInfoOnly] [-Server] [-Format <формат>]
    [-Archive <файл>] [-ignoreUnresolvedReferences]
```

#### Режимы выгрузки

**Полная выгрузка** — все объекты конфигурации:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpConfigToFiles "C:\src\config" /Out log.txt
```

**Инкрементальная выгрузка** — только изменённые объекты:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpConfigToFiles "C:\src\config" -update -force /Out log.txt
```

Инкрементальная выгрузка с отслеживанием изменений:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpConfigToFiles "C:\src\config" -update -getChanges "changes.txt" -configDumpInfoForChanges "old\ConfigDumpInfo.xml" /Out log.txt
```

**Частичная выгрузка** — выбранные объекты по списку:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpConfigToFiles "C:\src\config" -listFile "dump_objects.txt" /Out log.txt
```

**Обновление ConfigDumpInfo.xml** — без выгрузки файлов:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /DumpConfigToFiles "C:\src\config" -configDumpInfoOnly /Out log.txt
```

#### Параметры выгрузки

| Параметр | Описание |
|----------|----------|
| `-update` | Обновляющая (инкрементальная) выгрузка — только изменённые объекты |
| `-force` | Принудительная полная выгрузка. Используется с `-update` при несовпадении версий |
| `-getChanges <файл>` | Записать список изменённых файлов |
| `-configDumpInfoForChanges <файл>` | Файл ConfigDumpInfo.xml для определения изменений |
| `-listFile <файл>` | Файл со списком выгружаемых объектов (по одному на строку) |
| `-configDumpInfoOnly` | Выгрузить только ConfigDumpInfo.xml |
| `-Extension <имя>` | Выгрузить расширение |
| `-AllExtensions` | Выгрузить все расширения |
| `-Server` | Выгрузка на стороне сервера |
| `-Format <формат>` | Формат файлов (Hierarchical / Plain) |
| `-Archive <файл>` | Выгрузка в архивный файл |
| `-ignoreUnresolvedReferences` | Игнорировать неразрешённые ссылки |

#### Формат listFile для выгрузки

Файл содержит **имена объектов метаданных** (одно на строку):
```
Справочник.Номенклатура
Справочник.Валюты
Документ.РеализацияТоваровУслуг
Отчет.АнализПродаж
```

Кодировка: UTF-8 с BOM.

### Загрузка `/LoadConfigFromFiles`

```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /LoadConfigFromFiles <каталог> [параметры] /Out log.txt
```

Полная сигнатура:
```
/LoadConfigFromFiles <каталог> [-Extension <имя>] [-AllExtensions]
    [-updateConfigDumpInfo] [-listFile <файл>]
    [-Server] [-Archive <файл>] [-Format <формат>]
```

#### Режимы загрузки

**Полная загрузка** — замена всей конфигурации:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /LoadConfigFromFiles "C:\src\config" /Out log.txt
```

**Частичная загрузка** — выбранные файлы по списку:
```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /LoadConfigFromFiles "C:\src\config" -listFile "load_list.txt" -Format Hierarchical -partial -updateConfigDumpInfo /Out log.txt
```

#### Параметры загрузки

| Параметр | Описание |
|----------|----------|
| `-listFile <файл>` | Файл со списком загружаемых файлов (по одному на строку) |
| `-partial` | Частичная загрузка — **не заменять** всю конфигурацию, а внести точечные изменения. Недокументированный, но рабочий параметр |
| `-updateConfigDumpInfo` | Обновить ConfigDumpInfo.xml после загрузки |
| `-Extension <имя>` | Загрузить в расширение |
| `-AllExtensions` | Загрузить все расширения |
| `-Server` | Загрузка на стороне сервера |
| `-Archive <файл>` | Загрузка из архивного файла |
| `-Format <формат>` | Формат файлов (Hierarchical / Plain) |

#### Формат listFile для загрузки

Файл содержит **относительные пути к файлам** в каталоге выгрузки (один на строку):
```
Catalogs/Валюты.xml
Catalogs/Валюты/Ext/ObjectModule.bsl
Documents/РеализацияТоваровУслуг.xml
Documents/РеализацияТоваровУслуг/Forms/ФормаДокумента.xml
```

Кодировка: UTF-8 с BOM.

> **Важно: различие форматов listFile для dump и load:**
> - **Выгрузка** (`/DumpConfigToFiles -listFile`): имена объектов метаданных — `Справочник.Номенклатура`
> - **Загрузка** (`/LoadConfigFromFiles -listFile`): относительные пути файлов — `Catalogs/Валюты.xml`

## Обновление конфигурации БД

```
1cv8.exe DESIGNER /F <база> /DisableStartupDialogs /UpdateDBCfg /Out log.txt
```

Полная сигнатура:
```
/UpdateDBCfg [-Dynamic<режим>] [-Server]
    [-WarningsAsErrors]
    [-BackgroundStart] [-BackgroundFinish]
    [-BackgroundCancel] [-BackgroundSuspend] [-BackgroundResume]
    [-Extension <имя>] [-AllExtensions]
```

| Параметр | Описание |
|----------|----------|
| `-Dynamic+` | Использовать динамическое обновление |
| `-Dynamic-` | Не использовать динамическое обновление |
| `-Server` | Обновление на стороне сервера |
| `-WarningsAsErrors` | Предупреждения считать ошибками |
| `-Extension <имя>` | Обновить расширение |
| `-AllExtensions` | Обновить все расширения |

### Фоновое обновление

| Параметр | Описание |
|----------|----------|
| `-BackgroundStart` | Начать фоновое обновление |
| `-BackgroundFinish` | Дождаться окончания и завершить |
| `-BackgroundCancel` | Отменить фоновое обновление |
| `-BackgroundSuspend` | Приостановить |
| `-BackgroundResume` | Возобновить |

> После `/LoadCfg` или `/LoadConfigFromFiles` необходимо выполнить `/UpdateDBCfg` чтобы изменения применились к базе данных.

## Сборка и разборка внешних обработок (EPF/ERF)

### Сборка (XML → EPF)

```
1cv8.exe DESIGNER /F <путь_к_базе> /DisableStartupDialogs /LoadExternalDataProcessorOrReportFromFiles <корневой_xml> <путь_к_epf> /Out <лог_файл>
```

| Параметр | Описание |
|----------|----------|
| `<корневой_xml>` | Путь к корневому XML-файлу обработки (например, `src\МояОбработка.xml`) |
| `<путь_к_epf>` | Путь к выходному файлу `.epf` или `.erf` |

> **Важно**: первый аргумент — путь к **корневому XML-файлу** (не к каталогу). Если указать каталог, конфигуратор вернёт ошибку.

### Разборка (EPF → XML)

```
1cv8.exe DESIGNER /F <путь_к_базе> /DisableStartupDialogs /DumpExternalDataProcessorOrReportToFiles <каталог_выгрузки> <путь_к_epf> [-Format Hierarchical] /Out <лог_файл>
```

| Параметр | Описание |
|----------|----------|
| `<каталог_выгрузки>` | Каталог для XML-файлов |
| `<путь_к_epf>` | Исходный файл `.epf` или `.erf` |
| `-Format Hierarchical` | Иерархическая структура каталогов (по умолчанию) |
| `-Format Plain` | Плоская структура |

### Примечания по сборке

- Если база не указана — скрипт `epf-build.ps1` автоматически создаёт временную базу. Для обработок со ссылочными типами (`CatalogRef.*`, `DocumentRef.*` и т.п.) генерируются заглушки метаданных. Временная база удаляется после сборки.
- Категории колонок регистров (Dimension/Resource/Attribute) угадываются по Form.xml — при round-trip через реальную базу привязки полей формы могут не сохраниться.

### Примечания по разборке

- Разборка **обязательно** требует базу с конфигурацией, содержащей используемые типы.
- Dump в пустой базе **безвозвратно** теряет ссылочные типы — `CatalogRef.XXX` превращается в `xs:string`.

## Запуск в режиме предприятия

```
1cv8.exe ENTERPRISE /F <база> [/N<имя> /P<пароль>] /DisableStartupDialogs [параметры]
```

| Параметр | Описание |
|----------|----------|
| `/Execute <файл.epf>` | Запуск внешней обработки сразу после старта. При указании `/Execute` параметр `/URL` игнорируется |
| `/URL <ссылка>` | Навигационная ссылка (формат `e1cib/...`) |
| `/C <строка>` | Передача параметра в прикладное решение |

Примеры:
```
1cv8.exe ENTERPRISE /F "C:\Bases\MyBase" /NАдмин /PSecret /DisableStartupDialogs /Execute "C:\scripts\process.epf"
```

```
1cv8.exe ENTERPRISE /IBName "Бухгалтерия" /NАдмин /DisableStartupDialogs /URL "e1cib/data/Справочник.Номенклатура"
```

## Коды возврата

| Код | Значение |
|-----|----------|
| `0` | Успешно |
| `1` | Ошибка |
| `101` | Ошибки при проверке конфигурации |

Числовой код можно записать в файл через `/DumpResult <файл>`.

При работе с расширениями (`-Extension`, `-AllExtensions`): 0 — успех, 1 — ошибка.

## ConfigDumpInfo.xml

`ConfigDumpInfo.xml` — служебный файл, создаваемый при выгрузке конфигурации в файлы (`/DumpConfigToFiles`). Содержит информацию о составе и версиях объектов конфигурации на момент выгрузки.

**Назначение:**
- Определение изменений при инкрементальной выгрузке (`-update`, `-configDumpInfoForChanges`)
- Синхронизация состояния выгрузки с конфигурацией ИБ

**Использование:**
- `-configDumpInfoForChanges <файл>` — передать предыдущий ConfigDumpInfo.xml для определения изменений
- `-configDumpInfoOnly` — обновить только этот файл без выгрузки объектов
- `-updateConfigDumpInfo` — обновить файл после частичной загрузки (`/LoadConfigFromFiles`)

**Расположение:** корень каталога выгрузки (рядом с `Configuration.xml`).

## Переменные окружения

| Переменная | Описание |
|-----------|----------|
| `V8_PATH` | Каталог `bin` платформы 1С (например, `C:\Program Files\1cv8\8.3.27.1859\bin`) |
| `V8_BASE` | Путь к пустой ИБ для EPF-сборки (создаётся автоматически при первом запуске) |



---

## Часть 4: Спецификация расширений конфигурации (XML)

# Спецификация формата выгрузки расширений конфигурации 1С (CFE)

Формат: XML-выгрузка расширения конфигурации 1С:Предприятие 8.3 (Конфигуратор → Конфигурация → Расширения → Выгрузить расширение в файлы).
Версия формата: `2.17` (платформа 8.3.17–8.3.24).

> **Связанные спецификации:**
> - Корневая структура конфигурации — [1c-configuration-spec.md](1c-configuration-spec.md)
> - Объекты метаданных — [1c-config-objects-spec.md](1c-config-objects-spec.md)
> - Подсистемы — [1c-subsystem-spec.md](1c-subsystem-spec.md)
> - Управляемые формы — [1c-form-spec.md](1c-form-spec.md)
> - Роли — [1c-role-spec.md](1c-role-spec.md)
> - Сводный индекс — [1c-specs-index.md](1c-specs-index.md)

---

## 1. Общая структура выгрузки расширения

```
Configuration.xml                  # Корневой файл — свойства и состав расширения
ConfigDumpInfo.xml                 # Служебный файл — версии объектов
Languages/                         # Языки (всегда заимствованные)
Roles/                             # Роли (собственные)
Subsystems/                        # Подсистемы (собственные или заимствованные)
CommonModules/                     # Общие модули
CommonPictures/                    # Общие картинки
CommonCommands/                    # Общие команды
Catalogs/                          # Справочники
Documents/                         # Документы
Enums/                             # Перечисления
...                                # Другие типы объектов
```

### Ключевые отличия от конфигурации

| Аспект | Конфигурация | Расширение |
|--------|-------------|------------|
| Корневой `Ext/` | Есть (модули, интерфейс, справка) | **Нет** |
| `ObjectBelonging` в Properties | Нет | `Adopted` (всегда) |
| `ConfigurationExtensionPurpose` | Нет | `Patch` / `Customization` / `AddOn` |
| `KeepMappingToExtendedConfigurationObjectsByIDs` | Нет | `true` / `false` |
| `NamePrefix` | Пустой или нет | Префикс для собственных объектов |
| `CompatibilityMode` | Да | Нет (используется `ConfigurationExtensionCompatibilityMode`) |
| Свойства режимов работы | Полный набор | Сокращённый набор |
| Объекты в ChildObjects | Только собственные | Собственные **и заимствованные** |

---

## 2. Configuration.xml — корневой файл расширения

### 2.1. Общая структура

```xml
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses"
    xmlns:v8="http://v8.1c.ru/8.1/data/core"
    xmlns:xr="http://v8.1c.ru/8.3/xcf/readable"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:app="http://v8.1c.ru/8.2/managed-application/core"
    ... version="2.17">
  <Configuration uuid="...">
    <InternalInfo>...</InternalInfo>
    <Properties>...</Properties>
    <ChildObjects>...</ChildObjects>
  </Configuration>
</MetaDataObject>
```

Пространства имён и корневой элемент идентичны конфигурации. Атрибут `version` соответствует версии формата выгрузки.

### 2.2. InternalInfo

Содержит 7 записей `xr:ContainedObject` — аналогично конфигурации. ClassId фиксированные, ObjectId уникальны для каждого расширения.

### 2.3. Properties — свойства расширения

Свойства идут в фиксированном порядке. Набор свойств **отличается** от конфигурации — часть свойств специфична для расширений, часть свойств конфигурации отсутствует.

#### Специфичные свойства расширения

| Свойство | Тип | Описание |
|----------|-----|----------|
| `ObjectBelonging` | enum | Всегда `Adopted` — расширение «принято» к основной конфигурации |
| `ConfigurationExtensionPurpose` | enum | Назначение расширения: `Patch` (исправление), `Customization` (адаптация), `AddOn` (дополнение) |
| `KeepMappingToExtendedConfigurationObjectsByIDs` | `xs:boolean` | Сохранять привязку к объектам по идентификаторам |
| `NamePrefix` | `xs:string` | Префикс имён собственных объектов (напр. `Расш1_`, `МоёРасш_`) |
| `ConfigurationExtensionCompatibilityMode` | enum | Режим совместимости расширения (`Version8_3_17`, `Version8_3_24`, ...) |

#### Общие свойства (совпадают с конфигурацией)

| Свойство | Тип | Описание |
|----------|-----|----------|
| `Name` | `xs:string` | Имя расширения (идентификатор) |
| `Synonym` | `LocalString` | Отображаемое имя |
| `Comment` | `xs:string` | Комментарий |
| `DefaultRunMode` | enum | Режим запуска (`ManagedApplication`) |
| `UsePurposes` | list | Назначения (`PlatformApplication`) |
| `ScriptVariant` | enum | Язык скриптов (`Russian` / `English`) |
| `DefaultRoles` | list | Роли по умолчанию |
| `Vendor` | `xs:string` | Поставщик |
| `Version` | `xs:string` | Версия расширения |
| `DefaultLanguage` | ref | Язык по умолчанию (`Language.Русский`) |
| `BriefInformation` | `LocalString` | Краткая информация |
| `DetailedInformation` | `LocalString` | Подробная информация |
| `Copyright` | `LocalString` | Авторские права |
| `VendorInformationAddress` | `LocalString` | Адрес поставщика |
| `ConfigurationInformationAddress` | `LocalString` | Адрес информации |
| `InterfaceCompatibilityMode` | enum | Совместимость интерфейса |

> **Примечание:** Свойства `DefaultRunMode`, `UsePurposes`, `DefaultRoles`, `DefaultLanguage`, `InterfaceCompatibilityMode` **опциональны** — могут отсутствовать в расширении (в отличие от конфигурации, где они обязательны).

#### Свойства конфигурации, отсутствующие в расширении

В расширениях **нет** следующих свойств:
- `CompatibilityMode` (заменено на `ConfigurationExtensionCompatibilityMode`)
- `DataLockControlMode`
- `ObjectAutonumerationMode`
- `ModalityUseMode`
- `SynchronousPlatformExtensionAndAddInCallUseMode`
- `DatabaseTablespacesUseMode`
- `MainClientApplicationWindowMode`
- `UpdateCatalogAddress`
- `IncludeHelpInContents`
- `UseManagedFormInOrdinaryApplication`
- `UseOrdinaryFormInManagedApplication`
- `Content`
- `StandaloneConfigurationRestrictionRoles`

### 2.4. Порядок свойств

```xml
<Properties>
  <ObjectBelonging>Adopted</ObjectBelonging>
  <Name>ИмяРасширения</Name>
  <Synonym>...</Synonym>
  <Comment/>
  <ConfigurationExtensionPurpose>Patch</ConfigurationExtensionPurpose>
  <KeepMappingToExtendedConfigurationObjectsByIDs>true</KeepMappingToExtendedConfigurationObjectsByIDs>
  <NamePrefix>Расш1_</NamePrefix>
  <ConfigurationExtensionCompatibilityMode>Version8_3_17</ConfigurationExtensionCompatibilityMode>
  <DefaultRunMode>ManagedApplication</DefaultRunMode>          <!-- опционально -->
  <UsePurposes>...</UsePurposes>                                <!-- опционально -->
  <ScriptVariant>Russian</ScriptVariant>
  <DefaultRoles>...</DefaultRoles>                              <!-- опционально -->
  <Vendor/>
  <Version/>
  <DefaultLanguage>Language.Русский</DefaultLanguage>           <!-- опционально -->
  <BriefInformation/>
  <DetailedInformation/>
  <Copyright/>
  <VendorInformationAddress/>
  <ConfigurationInformationAddress/>
  <InterfaceCompatibilityMode>TaxiEnableVersion8_2</InterfaceCompatibilityMode>  <!-- опционально -->
</Properties>
```

### 2.5. ChildObjects — состав расширения

Содержит как **собственные** объекты расширения, так и **заимствованные** из основной конфигурации. Порядок типов аналогичен конфигурации.

```xml
<ChildObjects>
  <Language>Русский</Language>                          <!-- заимствованный -->
  <Subsystem>Расш1_МояПодсистема</Subsystem>              <!-- собственный -->
  <CommonPicture>Расш1_МояКартинка</CommonPicture>       <!-- собственный -->
  <Role>Расш1_ОсновнаяРоль</Role>                        <!-- собственный -->
  <CommonModule>Расш1_МодульСервер</CommonModule>         <!-- собственный -->
  <CommonModule>ОбщийМодульКонфигурации</CommonModule>    <!-- заимствованный -->
  <Catalog>Контрагенты</Catalog>                         <!-- заимствованный -->
  <Catalog>Расш1_Проекты</Catalog>                       <!-- собственный -->
  <Enum>Расш1_ВидыДокументов</Enum>                      <!-- собственный -->
  <InformationRegister>Расш1_ДатыРабот</InformationRegister> <!-- собственный -->
</ChildObjects>
```

В `ChildObjects` не видно различие между собственными и заимствованными — оно определяется по содержимому XML-файла объекта (свойство `ObjectBelonging` и наличие `ExtendedConfigurationObject`).

**Правило именования:** собственные объекты расширения обычно имеют `NamePrefix` в начале имени (напр. `Расш1_Справочник1`, `Расш1_Проекты`), заимствованные — имя объекта из основной конфигурации без префикса (напр. `Контрагенты`, `Валюты`).

---

## 3. ConfigDumpInfo.xml

Формат идентичен конфигурации:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ConfigDumpInfo xmlns="http://v8.1c.ru/8.3/xcf/dumpinfo"
    xmlns:xen="http://v8.1c.ru/8.3/xcf/enums"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    format="Hierarchical" version="2.17">
  <ConfigVersions>
    <Metadata name="Configuration.ИмяРасширения" id="uuid" configVersion="sha1"/>
    <Metadata name="Language.Русский" id="uuid" configVersion="sha1"/>
    <Metadata name="Role.Расш1_ОсновнаяРоль" id="uuid" configVersion="sha1"/>
    <!-- ... все объекты расширения ... -->
  </ConfigVersions>
</ConfigDumpInfo>
```

Включает записи для **всех** объектов расширения (и собственных, и заимствованных). Атрибут `configVersion` — 40-символьный SHA1-хеш версии объекта.

---

## 4. Заимствованные и собственные объекты

Расширение может содержать два типа объектов:

### 4.1. Заимствованные объекты (Adopted)

Объекты, существующие в основной конфигурации, которые расширение модифицирует или дополняет.

**Признаки:**
- `<ObjectBelonging>Adopted</ObjectBelonging>` в Properties
- `<ExtendedConfigurationObject>uuid</ExtendedConfigurationObject>` — UUID объекта в основной конфигурации
- Минимальный набор свойств (только те, что изменяются)

```xml
<Catalog uuid="81de7e56-...">
  <InternalInfo>
    <xr:GeneratedType name="CatalogObject.Валюты" category="Object">...</xr:GeneratedType>
    <!-- ... стандартные GeneratedType ... -->
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>Валюты</Name>
    <Comment/>
    <ExtendedConfigurationObject>7aadbb67-...</ExtendedConfigurationObject>
  </Properties>
  <ChildObjects>
    <!-- заимствованные и собственные реквизиты, формы -->
  </ChildObjects>
</Catalog>
```

Заимствованные объекты **не содержат** полного набора свойств — только `ObjectBelonging`, `Name`, `Comment`, `ExtendedConfigurationObject` и те свойства, которые расширение изменяет (напр. `CodeLength`, `DefaultListForm`).

### 4.2. Собственные объекты (Own)

Объекты, созданные непосредственно в расширении.

**Признаки:**
- **Нет** элемента `ObjectBelonging`
- **Нет** элемента `ExtendedConfigurationObject`
- Полный набор свойств (как в объектах конфигурации)
- Имя обычно начинается с `NamePrefix` расширения

```xml
<Catalog uuid="7dcd4d14-...">
  <InternalInfo>
    <xr:GeneratedType name="CatalogObject.Расш5_Справочник1" category="Object">...</xr:GeneratedType>
    <!-- ... -->
  </InternalInfo>
  <Properties>
    <Name>Расш5_Справочник1</Name>
    <Synonym/>
    <Comment/>
    <Hierarchical>false</Hierarchical>
    <CodeLength>9</CodeLength>
    <!-- ... полный набор свойств как в конфигурации ... -->
  </Properties>
  <ChildObjects/>
</Catalog>
```

Формат полностью совпадает с форматом объектов конфигурации (см. [1c-config-objects-spec.md](1c-config-objects-spec.md)).

---

## 5. Заимствованные дочерние элементы

Дочерние элементы заимствованных объектов (реквизиты, табличные части, значения перечислений) также маркируются как заимствованные или собственные.

### 5.1. Заимствованные реквизиты

```xml
<Attribute uuid="259e5f94-...">
  <InternalInfo/>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ОсновнаяВалюта</Name>
    <Comment/>
    <ExtendedConfigurationObject>206abcd3-...</ExtendedConfigurationObject>
    <Type>
      <v8:Type>cfg:CatalogRef.Валюты</v8:Type>
    </Type>
  </Properties>
</Attribute>
```

Заимствованный реквизит содержит `ObjectBelonging: Adopted` и `ExtendedConfigurationObject`. Набор свойств минимальный (Name, Comment, ExtendedConfigurationObject, Type).

### 5.2. Собственные реквизиты в заимствованном объекте

```xml
<Attribute uuid="7fabdcb4-...">
  <Properties>
    <Name>Расш5_Реквизит1</Name>
    <Synonym/>
    <Comment/>
    <Type>
      <v8:Type>cfg:CatalogRef.Расш5_Справочник1</v8:Type>
    </Type>
    <PasswordMode>false</PasswordMode>
    <!-- ... полный набор свойств ... -->
  </Properties>
</Attribute>
```

Собственный реквизит **не имеет** `ObjectBelonging` и `ExtendedConfigurationObject`. Содержит полный набор свойств.

### 5.3. Заимствование реквизитов объекта для работы с формой

Когда в расширении заимствуется форма объекта и требуется сохранить привязки к данным (`DataPath`), необходимо заимствовать реквизиты объекта, на которые форма ссылается. Это позволяет добавлять новые элементы на форму через расширение.

**Что заимствуется**:
- Реквизиты объекта, на которые ссылается форма через `Объект.XXX` в `<DataPath>`
- Табличные части целиком (все колонки), если форма обращается к `Объект.ТЧ`
- Зависимые объекты по типам реквизитов (`CatalogRef.X` → заимствуется `Catalog.X` как оболочка)
- Глубокие ссылки: `Объект.A.B` → реквизит B заимствуется внутри каталога A

**Обогащение Properties**: заимствованный объект с реквизитами содержит дополнительные свойства из источника (`Hierarchical`, `CodeLength`, `DescriptionLength`, `CodeType`, `CodeAllowedLength`), которые отсутствуют у простой оболочки.

**Заимствованная табличная часть** содержит `InternalInfo` с `GeneratedType` (категории `TabularSection` и `TabularSectionRow`):

```xml
<TabularSection uuid="...">
  <InternalInfo>
    <xr:GeneratedType name="CatalogTabularSection.Номенклатура.ДрагоценныеМатериалы" category="TabularSection">
      <xr:TypeId>...</xr:TypeId>
      <xr:ValueId>...</xr:ValueId>
    </xr:GeneratedType>
    <xr:GeneratedType name="CatalogTabularSectionRow.Номенклатура.ДрагоценныеМатериалы" category="TabularSectionRow">
      <xr:TypeId>...</xr:TypeId>
      <xr:ValueId>...</xr:ValueId>
    </xr:GeneratedType>
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ДрагоценныеМатериалы</Name>
    <Comment/>
    <ExtendedConfigurationObject>...</ExtendedConfigurationObject>
  </Properties>
  <ChildObjects>
    <!-- все реквизиты ТЧ — adopted -->
  </ChildObjects>
</TabularSection>
```

**Типы в реквизитах**: `v8:Type` для обычных типов (`cfg:CatalogRef.X`, `cfg:EnumRef.X`), `v8:TypeSet` для определяемых типов (`cfg:DefinedType.X`).

### 5.4. Заимствованные значения перечислений

```xml
<EnumValue uuid="9bc7380f-...">
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>НаценкаНаКурсДругойВалюты</Name>
    <ExtendedConfigurationObject>c9ab3890-...</ExtendedConfigurationObject>
    <Synonym>...</Synonym>
    <Comment/>
  </Properties>
</EnumValue>
```

### 5.5. Формы в расширениях

В расширении существуют **два принципиально разных сценария** работы с формами:

| Сценарий | Описание | `<BaseForm>` | ID элементов | `callType` |
|----------|----------|:------------:|:------------:|:----------:|
| **Собственная форма** на заимствованном объекте | Новая форма, не существующая в базовой конфигурации | Нет | Обычные (1+) | Нет |
| **Заимствованная форма** | Расширение существующей формы базовой конфигурации | Есть | Базовые + 1000000+ | Есть |

> **Как отличить:** Если файл метаданных формы (`.xml`) содержит `<ObjectBelonging>Adopted</ObjectBelonging>` — это заимствованная форма. Собственные формы не имеют `ObjectBelonging`.

#### 5.5.1. Метаданные заимствованной формы

Файл `.xml` в каталоге `Forms/`:

```xml
<Form uuid="8fcebcc1-...">
  <InternalInfo/>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ФормаСписка</Name>
    <Comment/>
    <ExtendedConfigurationObject>5f91b00f-...</ExtendedConfigurationObject>
    <FormType>Managed</FormType>
  </Properties>
</Form>
```

Содержимое формы хранится в `Forms/ФормаСписка/Ext/Form.xml`, модуль формы — в `Forms/ФормаСписка/Ext/Form/Module.bsl`.

#### 5.5.2. Структура Form.xml заимствованной формы

Form.xml заимствованной формы — **двухчастный файл**: Part 1 (результирующая форма) и BaseForm (исходная форма). Существуют **два варианта** в зависимости от наличия модификаций модуля формы.

##### Вариант A — Минимальная форма (чистое заимствование)

Когда форма заимствована без модификации модуля — Form.xml содержит **только свойства формы**, AutoCommandBar без кнопок и пустые Attributes. **Нет ChildItems**. Обе секции (Part 1 и BaseForm) идентичны.

```xml
<Form xmlns="http://v8.1c.ru/8.3/xcf/logform" ... version="2.17">
  <AutoTitle>false</AutoTitle>
  <AutoTime>CurrentOrLast</AutoTime>
  <!-- ... другие свойства формы ... -->
  <AutoCommandBar name="ФормаКоманднаяПанель" id="-1">
    <Autofill>false</Autofill>
  </AutoCommandBar>
  <Attributes/>
  <!-- Events, Commands — только расширения (если есть) -->

  <BaseForm version="2.17">
    <AutoTitle>false</AutoTitle>
    <AutoTime>CurrentOrLast</AutoTime>
    <!-- те же свойства -->
    <AutoCommandBar name="ФормаКоманднаяПанель" id="-1">
      <Autofill>false</Autofill>
    </AutoCommandBar>
    <Attributes/>
  </BaseForm>
</Form>
```

##### Вариант B — Полная форма (заимствование процедуры модуля через `ИзменениеИКонтроль`)

Когда в расширении заимствуется процедура из модуля формы, Конфигуратор выгружает **полное дерево ChildItems** с применением правил очистки.

```xml
<Form xmlns="http://v8.1c.ru/8.3/xcf/logform" ... version="2.17">
  <AutoTitle>false</AutoTitle>
  <!-- свойства формы -->
  <AutoCommandBar name="ФормаКоманднаяПанель" id="-1">
    <Autofill>false</Autofill>
    <!-- Без ChildItems (кнопки удалены) -->
  </AutoCommandBar>
  <ChildItems>
    <!-- Полное дерево визуальных элементов -->
  </ChildItems>
  <Attributes/>
  <!-- Events, Commands — только расширения -->

  <BaseForm version="2.17">
    <!-- Идентичная копия (свойства + AutoCommandBar + ChildItems + Attributes) -->
  </BaseForm>
</Form>
```

**Ключевые правила (для обоих вариантов):**

1. **Свойства формы** — элементы между `<Form>` и `<AutoCommandBar>` (напр. `AutoTitle`, `AutoTime`, `UsePostingMode`, `RepostOnWrite`, `WindowOpeningMode`, `Customizable`, `CommandBarLocation`) копируются из исходной формы в обе секции.

2. **AutoCommandBar** — присутствует всегда с `id="-1"`, но без `<ChildItems>` (кнопки удаляются). `<Autofill>` = `false`.

3. **Attributes** — по умолчанию пустой `<Attributes/>` в обеих секциях. Атрибуты базовой конфигурации **не включаются**. Реквизиты расширения (id ≥ 1000000) добавляются только в Part 1.

   **Вариант с заимствованным основным реквизитом**: если реквизиты объекта заимствованы (см. раздел 5.3), секция `<Attributes>` содержит основной реквизит формы:
   ```xml
   <Attributes>
     <Attribute name="Объект" id="1000001">
       <Type><v8:Type>cfg:CatalogObject.Номенклатура</v8:Type></Type>
       <MainAttribute>true</MainAttribute>
       <SavedData>true</SavedData>
     </Attribute>
   </Attributes>
   ```
   Присутствует в обеих секциях (Part 1 и BaseForm). Тип зависит от родительского объекта: `CatalogObject`, `DocumentObject` и т.д.

4. **BaseForm** — последний элемент в `<Form>`, атрибут `version`. В BaseForm **нет** Events, Commands, Parameters.

5. **DataPath** — два варианта в зависимости от наличия заимствованного основного реквизита:
   - **Без основного реквизита**: все `<DataPath>` удаляются в обеих секциях (ссылаются на реквизиты, не включённые в расширение).
   - **С основным реквизитом**: `<DataPath>`, начинающиеся с `Объект.` **сохраняются** (реквизиты заимствованы и привязки валидны). DataPath формовых реквизитов (не начинающиеся с `Объект.`) — удаляются.

6. **TitleDataPath** — аналогично DataPath: удаляются без основного реквизита, `Объект.*` сохраняются с ним.

7. **TypeLink: удаление** (вариант B) — блоки `<TypeLink>` с `<xr:DataPath>Items.*</xr:DataPath>` удаляются (человекочитаемые пути, которые нельзя преобразовать в UUID-формат Конфигуратора).

8. **Events элементов: удаление** (вариант B) — все `<Events>` внутри визуальных элементов удаляются в обеих секциях. Обработчики расширения добавляются через `elementEvents` в Part 1 с `callType`.

9. **Picture stripping** (вариант B) — блоки `<Picture>` с `<xr:Ref>CommonPicture.XXX</xr:Ref>` удаляются, если `CommonPicture.XXX` **не заимствован** в расширение. Сам элемент PictureDecoration остаётся, только `<Picture>` убирается. `StdPicture.Print` сохраняется, остальные StdPicture удаляются.

10. **Авто-заимствование CommonPictures** — при заимствовании формы автоматически заимствуются все CommonPictures, на которые ссылаются элементы формы.

11. **Авто-заимствование StyleItems** — элементы формы ссылаются на StyleItems через `<Font ref="style:XXX" kind="StyleItem"/>` и `<BackColor>style:XXX</BackColor>`. Все такие StyleItems должны быть заимствованы. Стандартные стили (NormalTextFont, AccentColor, FormBackColor и др.) не имеют файлов и автоматически пропускаются.

12. **Авто-заимствование Enums + EnumValues** — `<ChoiceParameters>` могут содержать `<Value xsi:type="xr:DesignTimeRef">Enum.XXX.EnumValue.YYY</Value>`. Перечисление `Enum.XXX` заимствуется вместе с конкретными `EnumValue` (borrowed с `ExtendedConfigurationObject` указывающим на UUID оригинального значения).

#### 5.5.3. Нумерация ID элементов

| Диапазон | Принадлежность |
|----------|---------------|
| `-1` | Авто-командная панель (`AutoCommandBar`) — фиксированный ID |
| `1` – `999999` | Элементы базовой формы (сохраняют оригинальные ID) |
| `1000000`+ | Реквизиты (`Attributes`) и команды (`Commands`), добавленные расширением |

> **Важно:** Визуальные элементы форм (элементы в `ChildItems`), добавленные расширением в тело базовой формы, могут использовать ID из обычного диапазона (продолжая нумерацию базовой формы). Диапазон 1000000+ гарантирован для `Attributes` и `Commands`.

#### 5.5.4. Атрибут callType — перехват событий и команд

В заимствованных формах события и действия команд используют атрибут `callType` для определения момента перехвата:

| Значение | Описание |
|----------|----------|
| `Before` | Обработчик расширения вызывается **до** оригинального обработчика |
| `After` | Обработчик расширения вызывается **после** оригинального обработчика |
| `Override` | Обработчик расширения **заменяет** оригинальный обработчик |

##### События формы (form-level)

```xml
<Events>
  <Event name="OnCreateAtServer" callType="After">Расш1_ПриСозданииНаСервереПосле</Event>
  <Event name="OnOpen" callType="Before">Расш1_ПриОткрытииПеред</Event>
  <Event name="BeforeWriteAtServer" callType="After">Расш1_ПередЗаписьюНаСервереПосле</Event>
  <Event name="NotificationProcessing" callType="After">Расш1_ОбработкаОповещенияПосле</Event>
</Events>
```

##### События элементов формы (element-level)

```xml
<InputField name="Банк" id="37">
  <Events>
    <Event name="OnChange" callType="Before">Расш1_БанкПриИзменении</Event>
    <Event name="Clearing" callType="Before">Расш1_БанкОчистка</Event>
  </Events>
  ...
</InputField>

<Table name="СписокСпецификаций" id="102">
  <Events>
    <Event name="Selection" callType="Before">Расш1_СписокВыборПеред</Event>
  </Events>
  ...
</Table>
```

##### Действия команд (Command Action)

Команда может иметь **несколько элементов `<Action>`** с разными `callType`:

```xml
<!-- Перехват существующей команды базовой формы: до + после -->
<Command name="ПодборИзКлассификатора" id="1000002">
  <Action callType="Before">Расш1_ПодборИзКлассификатораПеред</Action>
  <Action callType="After">Расш1_ПодборИзКлассификатораПосле</Action>
</Command>

<!-- Полная замена обработчика команды -->
<Command name="НоваяКоманда" id="1000000">
  <Action callType="Override">Расш1_НоваяКомандаВместо</Action>
</Command>

<!-- Один перехват (только после) -->
<Command name="ЗапросКорректировки" id="1000005">
  <Action callType="After">Расш1_ЗапросКорректировкиПосле</Action>
</Command>
```

> **Отличие от обычной формы:** В обычной форме (конфигурации или собственной форме расширения) у `<Event>` и `<Action>` **нет** атрибута `callType` — обработчик вызывается напрямую.

#### 5.5.5. Собственная форма на заимствованном объекте

Расширение может добавить к заимствованному объекту **собственную форму**, не существующую в базовой конфигурации. Такая форма:

- **Не имеет** `ObjectBelonging` и `ExtendedConfigurationObject` в метаданных формы
- **Не содержит** `<BaseForm>` в Form.xml
- **Не использует** атрибут `callType`
- Использует обычную нумерацию ID (1+)
- Формат полностью совпадает с форматом форм конфигурации (см. [1c-form-spec.md](1c-form-spec.md))

```xml
<!-- Метаданные: Forms/МояФорма.xml — без ObjectBelonging -->
<Form uuid="...">
  <Properties>
    <Name>Расш1_МояФорма</Name>
    <Synonym>...</Synonym>
    <Comment/>
    <FormType>Managed</FormType>
    <UsePurposes>...</UsePurposes>
  </Properties>
</Form>
```

```xml
<!-- Содержимое: Forms/МояФорма/Ext/Form.xml — обычная форма без BaseForm -->
<Form ... version="2.17">
  <Events>
    <Event name="OnCreateAtServer">ПриСозданииНаСервере</Event>
  </Events>
  <ChildItems>...</ChildItems>
  <Attributes>...</Attributes>
</Form>
```

#### 5.5.6. Модуль заимствованной формы

Модуль формы (`Forms/Имя/Ext/Form/Module.bsl`) в заимствованной форме использует те же декораторы перехвата, что и другие модули расширений (см. раздел 7.2):

```bsl
&НаСервере
&Вместо("ЗаполнитьПодменюПараметры")
Процедура Расш1_ЗаполнитьПодменюПараметры()
    ПродолжитьВызов();
КонецПроцедуры

&НаКлиенте
&ИзменениеИКонтроль("ПараметрыНаЯзыке")
Функция Расш1_ПараметрыНаЯзыке(КодЯзыка)
    // ... тело с #Вставка / #Удаление маркерами ...
КонецФункции

// Обработчик собственной команды расширения (без декоратора)
&НаКлиенте
Процедура Расш1_НоваяКомандаВместо(Команда)
    // ...
КонецПроцедуры
```

> **Обработчики событий с `callType`** (определённые в Form.xml секции Events/Action) реализуются в модуле как обычные процедуры **без** аннотаций-декораторов — привязка к событию уже задана в XML через `callType`.

---

## 6. Расширение свойств (xr:PropertyState и xr:ExtendedProperty)

Расширения могут изменять свойства заимствованных реквизитов. Для этого используются специальные XML-конструкции.

### 6.1. PropertyState — уведомление об изменении

Элемент `xr:PropertyState` в `InternalInfo` реквизита указывает, что свойство было изменено расширением.

```xml
<Attribute uuid="a1752169-...">
  <InternalInfo>
    <xr:PropertyState>
      <xr:Property>Type</xr:Property>
      <xr:State>Notify</xr:State>
    </xr:PropertyState>
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>Наценка</Name>
    <ExtendedConfigurationObject>87429f11-...</ExtendedConfigurationObject>
    <Type>
      <v8:Type>xs:decimal</v8:Type>
      <v8:NumberQualifiers>
        <v8:Digits>10</v8:Digits>
        <v8:FractionDigits>2</v8:FractionDigits>
        <v8:AllowedSign>Any</v8:AllowedSign>
      </v8:NumberQualifiers>
    </Type>
  </Properties>
</Attribute>
```

Значения `xr:State`:
| Значение | Описание |
|----------|----------|
| `Notify` | Свойство изменено расширением, платформа выводит предупреждение |
| `MultiState` | Свойство расширено (тип отличается от основной конфигурации) |

### 6.2. ExtendedProperty — расширение типа

Когда расширение изменяет тип реквизита, используется конструкция `xr:ExtendedProperty`:

```xml
<Type xsi:type="xr:ExtendedProperty">
  <xr:ExtendValue xsi:type="v8:TypeDescription">
    <v8:Type>xs:string</v8:Type>
    <v8:StringQualifiers>
      <v8:Length>60</v8:Length>
      <v8:AllowedLength>Variable</v8:AllowedLength>
    </v8:StringQualifiers>
  </xr:ExtendValue>
</Type>
```

Свойство `Type` получает атрибут `xsi:type="xr:ExtendedProperty"`, а значение оборачивается в `xr:ExtendValue`. При этом в `InternalInfo` указывается `<xr:State>MultiState</xr:State>`.

---

## 7. Модули в расширениях

### 7.1. Типы модулей

Расширения поддерживают те же типы модулей, что и конфигурация:

| Модуль | Файл | Для каких объектов |
|--------|------|--------------------|
| Модуль объекта | `Ext/ObjectModule.bsl` | Справочники, документы, обработки |
| Модуль менеджера | `Ext/ManagerModule.bsl` | Справочники, документы, регистры |
| Модуль набора записей | `Ext/RecordSetModule.bsl` | Регистры |
| Модуль формы | `Forms/Имя/Ext/Form/Module.bsl` | Формы |
| Общий модуль | `CommonModules/Имя/Ext/Module.bsl` | Общие модули |
| Модуль команды | `Commands/Имя/Ext/CommandModule.bsl` | Команды |

### 7.2. Декораторы перехвата

Модули расширений используют специальные **аннотации-декораторы** для перехвата вызовов процедур основной конфигурации:

| Декоратор | Описание |
|-----------|----------|
| `&Перед("ИмяПроцедуры")` | Выполняется **до** оригинальной процедуры |
| `&После("ИмяПроцедуры")` | Выполняется **после** оригинальной процедуры |
| `&Вместо("ИмяПроцедуры")` | **Заменяет** оригинальную процедуру |
| `&ИзменениеИКонтроль("ИмяПроцедуры")` | Копия с контролем изменений (diff-маркеры) |

#### Пример &Перед / &После

```bsl
&НаКлиенте
&Перед("ПодборИзКлассификатора")
Процедура Расш5_ПодборИзКлассификатораПеред(Команда)
    // Код выполняется ДО оригинальной процедуры
КонецПроцедуры

&НаКлиенте
&После("ПодборИзКлассификатора")
Процедура Расш5_ПодборИзКлассификатораПосле(Команда)
    // Код выполняется ПОСЛЕ оригинальной процедуры
КонецПроцедуры
```

#### Пример &Вместо

```bsl
&НаСервере
&Вместо("ЗаполнитьПодменюПараметрыПрописиВалюты")
Процедура Расш5_ЗаполнитьПодменюПараметрыПрописиВалюты()
    // Полная замена оригинальной процедуры
    // ПродолжитьВызов() — вызов оригинальной реализации
    ПродолжитьВызов();
КонецПроцедуры
```

#### Пример &ИзменениеИКонтроль

```bsl
&ИзменениеИКонтроль("РеквизитыРедактируемыеВГрупповойОбработке")
Функция Расш5_РеквизитыРедактируемыеВГрупповойОбработке()
    Результат = Новый Массив;
    Результат.Добавить("СпособУстановкиКурса");
#Удаление
    Результат.Добавить("ФормулаРасчетаКурса");
#КонецУдаления
#Вставка
    Результат.Добавить("НоваяФормулаРасчетаКурса");
#КонецВставки
    Возврат Результат;
КонецФункции
```

### 7.3. Diff-маркеры в коде

Внутри процедур с декоратором `&ИзменениеИКонтроль` используются diff-маркеры для отслеживания изменений:

| Маркер | Описание |
|--------|----------|
| `#Удаление` | Начало блока удалённого кода |
| `#КонецУдаления` | Конец блока удалённого кода |
| `#Вставка` | Начало блока вставленного кода |
| `#КонецВставки` | Конец блока вставленного кода |

Маркеры могут быть вложенными и чередующимися — типичный паттерн «удалить → вставить»:

```bsl
#Удаление
    ИначеЕсли Выборка.Дата > Дата Тогда
#КонецУдаления
#Вставка
    ИначеЕсли Выборка.Дата < Дата Тогда
#КонецВставки
```

### 7.4. Именование процедур расширения

Собственные процедуры расширения именуются с `NamePrefix`:
- `Расш5_ПрочитатьАрхивВнутр()` — при `NamePrefix = Расш5_`
- `МоёРасш_ПроверитьДату()` — при `NamePrefix = МоёРасш_`
- `Расш1_ПриОпределенииНастроек()` — при `NamePrefix = Расш1_`

---

## 8. Предопределённые элементы в расширениях

Расширения могут добавлять предопределённые элементы к заимствованным справочникам. Файл: `Каталог/Ext/Predefined.xml`.

```xml
<?xml version="1.0" encoding="UTF-8"?>
<PredefinedData xmlns="http://v8.1c.ru/8.3/xcf/predef"
    xmlns:v8="http://v8.1c.ru/8.1/data/core"
    xmlns:xr="http://v8.1c.ru/8.3/xcf/readable"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="CatalogPredefinedItems" version="2.17">
  <Item id="9b751d8b-...">
    <Name>НовыйЭлемент</Name>
    <Code>000000001</Code>
    <Description>Новый элемент</Description>
    <IsFolder>false</IsFolder>
    <ExtensionState>Native</ExtensionState>
  </Item>
</PredefinedData>
```

### Отличие от Predefined.xml конфигурации

| Аспект | Конфигурация | Расширение |
|--------|-------------|------------|
| Пространство имён | `http://v8.1c.ru/8.3/xcf/predef` (то же) | То же |
| `ExtensionState` | Нет | `Native` — элемент создан расширением |

---

## 9. Языки (Languages/)

Язык в расширении **всегда** заимствованный:

```xml
<Language uuid="9453bb96-...">
  <InternalInfo/>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>Русский</Name>
    <Comment/>
    <ExtendedConfigurationObject>0663bf5b-...</ExtendedConfigurationObject>
    <LanguageCode>ru</LanguageCode>
  </Properties>
</Language>
```

UUID в `ExtendedConfigurationObject` одинаков для всех расширений одной конфигурации — это UUID языка в основной конфигурации.

---

## 10. Роли (Roles/)

### 10.1. Собственная роль расширения (без прав)

Минимальная роль без `Ext/Rights.xml`:

```xml
<Role uuid="c630865b-...">
  <Properties>
    <Name>Расш1_ОсновнаяРоль</Name>
    <Synonym/>
    <Comment/>
  </Properties>
</Role>
```

Собственные роли расширений в изученных примерах **не имеют** каталога `Ext/` с `Rights.xml`. Права могут задаваться через конфигуратор.

### 10.2. DefaultRoles

Ссылки на роли по умолчанию в Configuration.xml:

```xml
<DefaultRoles>
  <xr:Item xsi:type="xr:MDObjectRef">Role.Расш1_ОсновнаяРоль</xr:Item>
</DefaultRoles>
```

---

## 11. Подсистемы (Subsystems/)

### 11.1. Собственная подсистема расширения

```xml
<Subsystem uuid="...">
  <Properties>
    <Name>Расш1_МояПодсистема</Name>
    <Synonym>...</Synonym>
    <Comment/>
    <Picture>
      <xr:Ref>CommonPicture.Расш1_МояКартинка</xr:Ref>
    </Picture>
    <IncludeHelpInContents>false</IncludeHelpInContents>
    <IncludeInCommandInterface>true</IncludeInCommandInterface>
    <Content>
      <xr:Item xsi:type="xr:MDObjectRef">Catalog.Расш1_Проекты</xr:Item>
      <xr:Item xsi:type="xr:MDObjectRef">Catalog.Расш1_Задачи</xr:Item>
      <xr:Item xsi:type="xr:MDObjectRef">DataProcessor.Расш1_Обработка1</xr:Item>
      <!-- ... -->
    </Content>
  </Properties>
  <ChildObjects/>
</Subsystem>
```

### 11.2. Заимствованная подсистема

```xml
<Subsystem uuid="...">
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>Закупки</Name>
    <ExtendedConfigurationObject>f230c0c7-...</ExtendedConfigurationObject>
    <Content>
      <!-- пустой или с добавленными объектами -->
    </Content>
  </Properties>
  <ChildObjects>
    <Subsystem>Расш1_Документы</Subsystem>   <!-- дочерняя подсистема -->
  </ChildObjects>
</Subsystem>
```

Заимствованная подсистема может содержать элементы `Content` (добавленные расширением команды/объекты) и дочерние подсистемы в `ChildObjects`.

### 11.3. CommandInterface в расширении

Командный интерфейс подсистемы в расширении: `Subsystems/Имя/Ext/CommandInterface.xml`.

```xml
<CommandInterface version="2.17">
  <CommandsVisibility>
    <xr:Command>
      <xr:CommandID>Document.ЗаказНаПеремещение.StandardCommand.OpenList</xr:CommandID>
      <xr:Visibility>
        <xr:Common>false</xr:Common>
        <xr:Value name="Role.Расш1_ОсновнаяРоль">true</xr:Value>
      </xr:Visibility>
    </xr:Command>
  </CommandsVisibility>
  <CommandsOrder>
    <xr:Group name="NavigationPanelOrdinary">
      <xr:CommandID>...</xr:CommandID>
    </xr:Group>
  </CommandsOrder>
</CommandInterface>
```

---

## 12. Общие модули (CommonModules/)

### 12.1. Заимствованный общий модуль

```xml
<CommonModule uuid="a32b77fa-...">
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ZipАрхивы</Name>
    <ExtendedConfigurationObject>b92e2bb8-...</ExtendedConfigurationObject>
  </Properties>
</CommonModule>
```

Модуль расширения: `CommonModules/ZipАрхивы/Ext/Module.bsl` — содержит процедуры с декораторами перехвата.

### 12.2. Собственный общий модуль

Формат аналогичен конфигурации (без `ObjectBelonging` и `ExtendedConfigurationObject`), со всеми свойствами (Server, ExternalConnection, ClientManagedApplication и т.д.).

---

## 13. Другие типы заимствованных объектов

### 13.1. Константы

```xml
<Constant uuid="...">
  <InternalInfo>
    <xr:GeneratedType name="ConstantManager.ИмяКонстанты" category="ConstantManager">...</xr:GeneratedType>
    <xr:GeneratedType name="ConstantValueManager.ИмяКонстанты" category="ConstantValueManager">...</xr:GeneratedType>
    <xr:GeneratedType name="ConstantValueKey.ИмяКонстанты" category="ConstantValueKey">...</xr:GeneratedType>
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяКонстанты</Name>
    <ExtendedConfigurationObject>81d26b82-...</ExtendedConfigurationObject>
    <Type>
      <v8:Type>xs:boolean</v8:Type>
    </Type>
  </Properties>
</Constant>
```

### 13.2. Функциональные опции

```xml
<FunctionalOption uuid="...">
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяФункциональнойОпции</Name>
    <ExtendedConfigurationObject>d2699502-...</ExtendedConfigurationObject>
    <Location>Constant.ИмяКонстанты</Location>
  </Properties>
</FunctionalOption>
```

### 13.3. Определяемые типы

Заимствованный (минимальный):
```xml
<DefinedType uuid="...">
  <InternalInfo>
    <xr:GeneratedType name="DefinedType.ИмяОпределяемогоТипа" category="DefinedType">...</xr:GeneratedType>
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяОпределяемогоТипа</Name>
  </Properties>
</DefinedType>
```

Собственный (с полным описанием типа):
```xml
<DefinedType uuid="...">
  <Properties>
    <Name>Расш1_Координата</Name>
    <Synonym>...</Synonym>
    <Type>
      <v8:Type>xs:decimal</v8:Type>
      <v8:NumberQualifiers>
        <v8:Digits>15</v8:Digits>
        <v8:FractionDigits>10</v8:FractionDigits>
        <v8:AllowedSign>Any</v8:AllowedSign>
      </v8:NumberQualifiers>
    </Type>
  </Properties>
</DefinedType>
```

### 13.4. Элементы стиля

```xml
<StyleItem uuid="...">
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяЭлементаСтиля</Name>
    <ExtendedConfigurationObject>3d428bdf-...</ExtendedConfigurationObject>
    <Type>Font</Type>
  </Properties>
</StyleItem>
```

Тип: `Color`, `Font`, `Border`.

### 13.5. Общие картинки

```xml
<CommonPicture uuid="...">
  <InternalInfo/>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяКартинки</Name>
  </Properties>
</CommonPicture>
```

Собственные картинки могут иметь каталог `Ext/` с `Picture.xml` и файлом изображения:

```
CommonPictures/Расш1_МояКартинка/
  Ext/
    Picture.xml          # <Picture><xr:Abs>Picture.png</xr:Abs>...</Picture>
    Picture/
      Picture.png        # Файл изображения
```

### 13.6. Общие команды

```xml
<CommonCommand uuid="...">
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяКоманды</Name>
    <Group>FormCommandBarImportant</Group>
  </Properties>
</CommonCommand>
```

### 13.7. Планы обмена

```xml
<ExchangePlan uuid="...">
  <InternalInfo>
    <xr:ThisNode>c335c2b8-...</xr:ThisNode>
    <xr:GeneratedType .../>
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяПланаОбмена</Name>
    <ExtendedConfigurationObject>0c01b26a-...</ExtendedConfigurationObject>
  </Properties>
  <ChildObjects/>
</ExchangePlan>
```

### 13.8. Планы счетов

```xml
<ChartOfAccounts uuid="...">
  <InternalInfo>
    <xr:GeneratedType .../>
  </InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>Хозрасчетный</Name>
    <ExtendedConfigurationObject>3796bdf5-...</ExtendedConfigurationObject>
  </Properties>
  <ChildObjects/>
</ChartOfAccounts>
```

### 13.9. Регистры сведений

Заимствованный (минимальный):
```xml
<InformationRegister uuid="...">
  <InternalInfo>...</InternalInfo>
  <Properties>
    <ObjectBelonging>Adopted</ObjectBelonging>
    <Name>ИмяРегистра</Name>
    <ExtendedConfigurationObject>...</ExtendedConfigurationObject>
    <InformationRegisterPeriodicity>Quarter</InformationRegisterPeriodicity>
    <WriteMode>Independent</WriteMode>
  </Properties>
  <ChildObjects>
    <Resource uuid="...">
      <Properties>
        <ObjectBelonging>Adopted</ObjectBelonging>
        <Name>ИмяРесурса</Name>
        <ExtendedConfigurationObject>...</ExtendedConfigurationObject>
        <Type>...</Type>
      </Properties>
    </Resource>
  </ChildObjects>
</InformationRegister>
```

Собственный регистр — полный набор свойств, аналогично конфигурации.

---

## 14. Назначения расширений (ConfigurationExtensionPurpose)

| Значение | Русское название | Описание |
|----------|-----------------|----------|
| `Patch` | Исправление | Минимальные исправления ошибок. Наибольшие ограничения |
| `Customization` | Адаптация | Доработка под требования заказчика. Средний уровень ограничений |
| `AddOn` | Дополнение | Добавление новой функциональности. Минимальные ограничения |

Назначение влияет на то, какие модификации допускает платформа при подключении расширения.

---

## 15. Структура каталогов (сводка)

### 15.1. Собственный объект (полная структура)

```
Catalogs/Расш1_Проекты.xml                       # Метаданные (полные)
Catalogs/Расш1_Проекты/
  Ext/
    ObjectModule.bsl                             # Модуль объекта
    ManagerModule.bsl                            # Модуль менеджера
    Predefined.xml                               # Предопределённые (опц.)
  Forms/
    ФормаЭлемента.xml                            # Метаданные формы
    ФормаЭлемента/
      Ext/
        Form.xml                                 # Содержимое формы
        Form/
          Module.bsl                             # Модуль формы
    ФормаСписка.xml
    ФормаСписка/
      Ext/
        Form.xml
        Form/
          Module.bsl
  Templates/
    ПФ_MXL_Акт.xml                               # Метаданные макета
    ПФ_MXL_Акт/
      Ext/
        Template.xml                             # Содержимое макета
  Commands/
    ИмяКоманды.xml                               # Метаданные команды
    ИмяКоманды/
      Ext/
        CommandModule.bsl                        # Модуль команды
```

### 15.2. Заимствованный объект (минимальная структура)

```
Catalogs/Валюты.xml                              # Метаданные (сокращённые)
Catalogs/Валюты/
  Ext/
    ManagerModule.bsl                            # Расширение модуля менеджера
    Predefined.xml                               # Предопред. элементы (опц.)
  Forms/
    ФормаСписка.xml                              # Метаданные (сокращённые)
    ФормаСписка/
      Ext/
        Form.xml                                 # Расширение формы
        Form/
          Module.bsl                             # Расширение модуля формы
```

### 15.3. Минимальное расширение (пустое)

```
Configuration.xml                                # Корневой файл
ConfigDumpInfo.xml                               # Версии объектов
Languages/
  Русский.xml                                    # Язык (заимствованный)
```

### 15.4. Типичное расширение с ролью

```
Configuration.xml
ConfigDumpInfo.xml
Languages/
  Русский.xml
Roles/
  Расш1_ОсновнаяРоль.xml
```

---

## 16. Отличия заимствованного объекта от обычного (сводная таблица)

| Аспект | Обычный (конфигурация) | Заимствованный (расширение) | Собственный (расширение) |
|--------|----------------------|--------------------------|------------------------|
| `ObjectBelonging` | Нет | `Adopted` | Нет |
| `ExtendedConfigurationObject` | Нет | UUID объекта конфигурации | Нет |
| Набор Properties | Полный | Минимальный + изменённые | Полный |
| InternalInfo | GeneratedType | GeneratedType + PropertyState | GeneratedType |
| Реквизиты в ChildObjects | Полные | Заимствованные + собственные | Полные |
| Модули | Полные | С декораторами перехвата | Полные |
| Формы | Полные | С расширениями (Ext/) | Полные |
