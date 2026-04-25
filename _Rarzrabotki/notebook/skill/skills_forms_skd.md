# Управляемые формы и Схема компоновки данных — спецификации XML и паттерны

> Полные спецификации управляемых форм (Form.xml), паттерны компоновки, и СКД.
> Источник: cc-1c-skills (Nikolay-Shirokov/cc-1c-skills)

---

## Часть 1: Спецификация управляемых форм

# 1C Form.xml Format Specification

Спецификация формата управляемых форм 1С:Предприятие 8.3 (version 2.17).
Составлена на основе анализа 7723 форм конфигурации «Бухгалтерия предприятия 3.0.180».

---

## 0. Файловая структура и регистрация

### Файлы формы

Каждая форма объекта конфигурации состоит из 3 файлов:

```
<Объект>/Forms/
  ИмяФормы.xml                  ← метаданные (UUID, имя, синоним, FormType)
  ИмяФормы/
    Ext/
      Form.xml                   ← определение формы (описано в разделах 1–17)
      Form/
        Module.bsl               ← модуль формы (1С-код)
```

Общие формы (CommonForm) — аналогично, но на верхнем уровне конфигурации:

```
CommonForms/
  ИмяФормы.xml                  ← метаданные (тег <CommonForm>)
  ИмяФормы/
    Ext/
      Form.xml
      Form/
        Module.bsl
```

### Метаданные формы — шаблон

#### Форма объекта (Document, Catalog, DataProcessor, Report, ...)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses"
		xmlns:app="http://v8.1c.ru/8.2/managed-application/core"
		xmlns:v8="http://v8.1c.ru/8.1/data/core"
		xmlns:xr="http://v8.1c.ru/8.3/xcf/readable"
		xmlns:xs="http://www.w3.org/2001/XMLSchema"
		xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
		version="2.17">
	<Form uuid="XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX">
		<Properties>
			<Name>ИмяФормы</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Отображаемое имя</v8:content>
				</v8:item>
			</Synonym>
			<Comment/>
			<FormType>Managed</FormType>
			<IncludeHelpInContents>false</IncludeHelpInContents>
			<UsePurposes>
				<v8:Value xsi:type="app:ApplicationUsePurpose">PlatformApplication</v8:Value>
				<v8:Value xsi:type="app:ApplicationUsePurpose">MobilePlatformApplication</v8:Value>
			</UsePurposes>
		</Properties>
	</Form>
</MetaDataObject>
```

#### CommonForm

```xml
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses"
		xmlns:app="http://v8.1c.ru/8.2/managed-application/core"
		xmlns:v8="http://v8.1c.ru/8.1/data/core"
		xmlns:xr="http://v8.1c.ru/8.3/xcf/readable"
		xmlns:xs="http://www.w3.org/2001/XMLSchema"
		xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
		version="2.17">
	<CommonForm uuid="XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX">
		<Properties>
			<Name>ИмяФормы</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Отображаемое имя</v8:content>
				</v8:item>
			</Synonym>
			<Comment/>
			<FormType>Managed</FormType>
			<IncludeHelpInContents>false</IncludeHelpInContents>
			<UseStandardCommands>false</UseStandardCommands>
			<ExtendedPresentation/>
			<Explanation/>
			<UsePurposes>
				<v8:Value xsi:type="app:ApplicationUsePurpose">PlatformApplication</v8:Value>
				<v8:Value xsi:type="app:ApplicationUsePurpose">MobilePlatformApplication</v8:Value>
			</UsePurposes>
		</Properties>
	</CommonForm>
</MetaDataObject>
```

### Регистрация формы

#### В ChildObjects родительского объекта

```xml
<!-- В файле Documents/АвансовыйОтчет.xml (или Catalogs/Контрагенты.xml и т.д.) -->
<ChildObjects>
	<Form>ФормаДокумента</Form>
	<Form>ФормаСписка</Form>
	...
</ChildObjects>
```

CommonForms регистрируются в `Configuration.xml`:

```xml
<ChildObjects>
	<CommonForm>ИмяФормы</CommonForm>
	...
</ChildObjects>
```

#### DefaultForm в Properties родительского объекта

Формат значения: `ТипОбъекта.ИмяОбъекта.Form.ИмяФормы`

```xml
<Properties>
	<DefaultObjectForm>Document.АвансовыйОтчет.Form.ФормаДокумента</DefaultObjectForm>
	<DefaultListForm>Document.АвансовыйОтчет.Form.ФормаСписка</DefaultListForm>
	<DefaultChoiceForm>Document.АвансовыйОтчет.Form.ФормаВыбора</DefaultChoiceForm>
</Properties>
```

#### Свойства DefaultForm по типам объектов

| Тип объекта | Свойства DefaultForm |
|-------------|---------------------|
| Document | DefaultObjectForm, DefaultListForm, DefaultChoiceForm |
| Catalog | DefaultObjectForm, DefaultFolderForm, DefaultListForm, DefaultChoiceForm, DefaultFolderChoiceForm |
| ChartOfCharacteristicTypes | DefaultObjectForm, DefaultFolderForm, DefaultListForm, DefaultChoiceForm, DefaultFolderChoiceForm |
| ChartOfAccounts | DefaultObjectForm, DefaultListForm, DefaultChoiceForm |
| DataProcessor | DefaultForm |
| Report | DefaultForm |
| InformationRegister | DefaultRecordForm, DefaultListForm |
| ExchangePlan | DefaultObjectForm, DefaultListForm, DefaultChoiceForm |
| BusinessProcess | DefaultObjectForm, DefaultListForm, DefaultChoiceForm |
| Task | DefaultObjectForm, DefaultListForm, DefaultChoiceForm |
| CommonForm | — (регистрируется в Configuration.xml, нет DefaultForm) |

> Report.DefaultForm может указывать на общую форму: `CommonForm.ФормаОтчета`.

---

## 1. Корневой элемент

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Form xmlns="http://v8.1c.ru/8.3/xcf/logform"
      xmlns:app="http://v8.1c.ru/8.2/managed-application/core"
      xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config"
      xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core"
      xmlns:dcssch="http://v8.1c.ru/8.1/data-composition-system/schema"
      xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings"
      xmlns:ent="http://v8.1c.ru/8.1/data/enterprise"
      xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform"
      xmlns:style="http://v8.1c.ru/8.1/data/ui/style"
      xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system"
      xmlns:v8="http://v8.1c.ru/8.1/data/core"
      xmlns:v8ui="http://v8.1c.ru/8.1/data/ui"
      xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web"
      xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows"
      xmlns:xr="http://v8.1c.ru/8.3/xcf/readable"
      xmlns:xs="http://www.w3.org/2001/XMLSchema"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
      version="2.17">
  ...
</Form>
```

Все 17 namespace-деклараций **идентичны** во всех формах конфигурации. Атрибут `version` всегда `"2.17"`.

### Назначение namespace-префиксов

| Префикс | URI | Назначение |
|---------|-----|------------|
| _(default)_ | `http://v8.1c.ru/8.3/xcf/logform` | Основная схема формы |
| `v8` | `http://v8.1c.ru/8.1/data/core` | Базовые типы данных (Type, item, lang, content) |
| `v8ui` | `http://v8.1c.ru/8.1/data/ui` | UI-типы (Color, Font, Border, FormattedString) |
| `cfg` | `http://v8.1c.ru/8.1/data/enterprise/current-config` | Ссылки на объекты конфигурации (CatalogRef, DocumentRef) |
| `xr` | `http://v8.1c.ru/8.3/xcf/readable` | Читаемый формат (Ref, Item, LoadTransparent) |
| `style` | `http://v8.1c.ru/8.1/data/ui/style` | Стили оформления (FormBackColor и т.д.) |
| `web` | `http://v8.1c.ru/8.1/data/ui/colors/web` | Web-цвета |
| `win` | `http://v8.1c.ru/8.1/data/ui/colors/windows` | Windows-цвета |
| `sys` | `http://v8.1c.ru/8.1/data/ui/fonts/system` | Системные шрифты |
| `xs` | `http://www.w3.org/2001/XMLSchema` | XML Schema |
| `xsi` | `http://www.w3.org/2001/XMLSchema-instance` | XML Schema Instance |
| `app` | `http://v8.1c.ru/8.2/managed-application/core` | Ядро управляемого приложения |
| `lf` | `http://v8.1c.ru/8.2/managed-application/logform` | Формы управляемого приложения |
| `dcscor` | `http://v8.1c.ru/8.1/data-composition-system/core` | СКД — ядро |
| `dcssch` | `http://v8.1c.ru/8.1/data-composition-system/schema` | СКД — схема |
| `dcsset` | `http://v8.1c.ru/8.1/data-composition-system/settings` | СКД — настройки |
| `ent` | `http://v8.1c.ru/8.1/data/enterprise` | Данные предприятия |

---

## 2. Структура Form — порядок дочерних элементов

```
<Form>
  ┌─ Свойства формы (необязательные, в произвольном порядке)
  ├─ <CommandSet>           — исключённые стандартные команды
  ├─ <AutoCommandBar>       — главная командная панель (обязательный, id="-1")
  ├─ <Events>               — обработчики событий формы
  ├─ <ChildItems>           — дерево UI-элементов
  ├─ <Attributes>           — реквизиты формы
  ├─ <Parameters>           — параметры открытия формы
  └─ <Commands>             — пользовательские команды
</Form>
```

---

## 3. Свойства формы

Прямые дочерние элементы `<Form>` (все необязательные, указываются до `<CommandSet>`/`<AutoCommandBar>`):

### Общие свойства (все типы форм)

| Элемент | Тип | Значения | Описание |
|---------|-----|----------|----------|
| `<Title>` | multilang | — | Заголовок формы |
| `<Width>` | int | 60, 67... | Ширина формы в символах |
| `<Height>` | int | — | Высота формы в символах |
| `<Group>` | enum | `Vertical`, `Horizontal`, `AlwaysHorizontal`, `AlwaysVertical` | Направление размещения |
| `<WindowOpeningMode>` | enum | `LockOwnerWindow`, `Modeless` | Режим открытия окна |
| `<EnterKeyBehavior>` | enum | `DefaultButton`, `NewLine` | Действие по Enter |
| `<AutoTitle>` | bool | `true`/`false` | Автозаголовок |
| `<AutoURL>` | bool | `true`/`false` | Авто-URL |
| `<AutoFillCheck>` | bool | `true`/`false` | Автопроверка заполнения |
| `<Customizable>` | bool | `true`/`false` | Разрешить настройку |
| `<CommandBarLocation>` | enum | `Top`, `Bottom`, `None` | Расположение панели команд |
| `<VerticalScroll>` | enum | `useIfNecessary`, `Auto`, `AlwaysShow`, `Never` | Вертикальная прокрутка |
| `<ScalingMode>` | enum | — | Режим масштабирования |

### Свойства сохранения данных (DataProcessors)

| Элемент | Значения | Описание |
|---------|----------|----------|
| `<SaveDataInSettings>` | `UseList`, `Use`, `DontUse` | Сохранять данные в настройках |
| `<AutoSaveDataInSettings>` | `Use`, `DontUse` | Автосохранение |

### Свойства документов (Documents)

| Элемент | Значения | Описание |
|---------|----------|----------|
| `<AutoTime>` | `CurrentOrLast`, `Current`, `Last` | Управление временем документа |
| `<UsePostingMode>` | `Auto`, `Postings`, `Movements` | Режим проведения |
| `<RepostOnWrite>` | `true`/`false` | Перепроведение при записи |

### Свойства справочников (Catalogs, ChartsOfAccounts)

| Элемент | Значения | Описание |
|---------|----------|----------|
| `<UseForFoldersAndItems>` | `Folders`, `Items`, `FoldersAndItems` | Назначение формы |

### Свойства отчётов (Reports)

| Элемент | Значения | Описание |
|---------|----------|----------|
| `<ReportResult>` | string | Имя реквизита результата (`Результат`) |
| `<DetailsData>` | string | Имя реквизита расшифровки (`ДанныеРасшифровки`) |
| `<ReportFormType>` | `Main`, `Settings`, `Choice` | Тип формы отчёта |
| `<AutoShowState>` | `Auto`, `Show`, `Hide` | Автоотображение состояния |
| `<ReportResultViewMode>` | `Auto`, `Table`, `Spreadsheet` | Режим отображения результата |
| `<ViewModeApplicationOnSetReportResult>` | `Auto`, `Always`, `Never` | Применение режима |

### Мобильные свойства

| Элемент | Описание |
|---------|----------|
| `<MobileDeviceCommandBarContent>` | Конфигурация панели команд мобильного устройства |

### Матрица свойств по типам форм

| Свойство | CommonForm | Document | Catalog | Report | DataProcessor | InfoRegister |
|----------|:---:|:---:|:---:|:---:|:---:|:---:|
| Title | + | + | + | + | + | + |
| Width | — | — | + | — | + | — |
| WindowOpeningMode | + | + | + | — | + | — |
| AutoTitle | + | + | + | + | + | + |
| CommandBarLocation | + | + | + | + | + | + |
| AutoTime | — | + | — | — | — | — |
| UsePostingMode | — | + | — | — | — | — |
| UseForFoldersAndItems | — | — | + | — | — | — |
| ReportResult | — | — | — | + | — | — |
| SaveDataInSettings | — | — | — | — | + | — |

---

## 4. CommandSet — исключённые команды

```xml
<CommandSet>
  <ExcludedCommand>CommandName</ExcludedCommand>
  ...
</CommandSet>
```

### Стандартные команды диалогов

`OK`, `Cancel`, `Yes`, `No`, `Abort`, `Retry`, `Ignore`, `Help`, `SaveValues`, `RestoreValues`

### Стандартные команды объектов

`Copy`, `Delete`, `SetDeletionMark`, `CreateInitialImage`, `ReadChanges`, `WriteChanges`

### Команды отчётов

`CustomizeForm`

---

## 5. AutoCommandBar — главная панель команд

Всегда присутствует. Фиксированные `name="ФормаКоманднаяПанель"` и `id="-1"`.

```xml
<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">
  <HorizontalAlign>Right</HorizontalAlign>    <!-- Left | Center | Right -->
  <Autofill>false</Autofill>                  <!-- true | false -->
  <EnableContentChange>true</EnableContentChange>  <!-- optional -->
  <ChildItems>
    <!-- Button, ButtonGroup, Popup -->
  </ChildItems>
</AutoCommandBar>
```

Может быть пустым (самозакрывающийся тег) или содержать `<ChildItems>` с кнопками.

---

## 6. Events — обработчики событий формы

```xml
<Events>
  <Event name="EventName">ИмяОбработчика</Event>
  ...
</Events>
```

### Все события формы

| Имя события | Контекст | Описание |
|-------------|----------|----------|
| `OnCreateAtServer` | Сервер | Создание формы на сервере (инициализация) |
| `OnOpen` | Клиент | Открытие формы на клиенте |
| `BeforeClose` | Клиент | Перед закрытием формы |
| `OnClose` | Клиент | При закрытии формы |
| `AfterWrite` | Клиент | После записи объекта |
| `BeforeWrite` | Клиент | Перед записью объекта |
| `BeforeWriteAtServer` | Сервер | Перед записью на сервере |
| `OnWriteAtServer` | Сервер | При записи на сервере |
| `AfterWriteAtServer` | Сервер | После записи на сервере |
| `OnReadAtServer` | Сервер | При чтении объекта |
| `NotificationProcessing` | Клиент | Обработка межформенных оповещений |
| `ChoiceProcessing` | Клиент | Обработка результата выбора |
| `NewWriteProcessing` | Сервер | Создание нового объекта |
| `FillCheckProcessingAtServer` | Сервер | Проверка заполнения |
| `OnLoadUserSettingsAtServer` | Сервер | Загрузка пользовательских настроек (отчёты) |
| `OnSaveUserSettingsAtServer` | Сервер | Сохранение пользовательских настроек (отчёты) |
| `URLProcessing` | Клиент | Обработка навигационных ссылок |

### Типичные комбинации по типам форм

**Диалог:** `OnCreateAtServer` + `OnOpen`

**Документ:** `OnCreateAtServer` + `OnOpen` + `BeforeWriteAtServer` + `OnWriteAtServer` + `AfterWrite`

**Справочник:** `OnCreateAtServer` + `OnOpen` + `OnReadAtServer` + `BeforeWriteAtServer` + `AfterWrite` + `NotificationProcessing`

**Отчёт:** `OnCreateAtServer` + `OnOpen` + `BeforeClose` + `OnClose` + `OnLoadUserSettingsAtServer` + `OnSaveUserSettingsAtServer` + `NotificationProcessing` + `ChoiceProcessing` + `URLProcessing`

---

## 7. ChildItems — дерево UI-элементов

### 7.1. Иерархия вложенности

```
ChildItems
├── UsualGroup          → содержит любые элементы
│   └── ChildItems
├── Pages               → содержит только Page
│   └── ChildItems
│       └── Page        → содержит любые элементы
│           └── ChildItems
├── Table               → содержит колонки (InputField, LabelField, CheckBoxField, PictureField)
│   └── ChildItems
├── CommandBar          → содержит Button, ButtonGroup, Popup
│   └── ChildItems
├── InputField          (лист)
├── LabelField          (лист)
├── CheckBoxField       (лист)
├── LabelDecoration     (лист)
├── PictureDecoration   (лист)
├── PictureField        (лист)
├── CalendarField       (лист)
└── Button              (лист)
```

### 7.2. Общие свойства всех элементов

Каждый UI-элемент имеет атрибуты `name` (string) и `id` (int). Кроме того, большинство элементов поддерживают:

| Свойство | Тип | Описание |
|----------|-----|----------|
| `<Title>` | multilang | Заголовок |
| `<ToolTip>` | multilang | Подсказка |
| `<Visible>` | bool | Видимость |
| `<Enabled>` | bool | Доступность |
| `<ReadOnly>` | bool | Только чтение |
| `<Width>` | int | Ширина |
| `<Height>` | int | Высота |
| `<HorizontalStretch>` | bool | Растягивание по горизонтали |
| `<VerticalStretch>` | bool | Растягивание по вертикали |
| `<HorizontalAlign>` | enum | `Left` / `Center` / `Right` |
| `<VerticalAlign>` | enum | `Top` / `Center` / `Bottom` |
| `<GroupHorizontalAlign>` | enum | Горизонтальное выравнивание в группе |
| `<GroupVerticalAlign>` | enum | Вертикальное выравнивание в группе |
| `<SkipOnInput>` | bool | Пропускать при вводе |
| `<ContextMenu>` | ref | Контекстное меню (name + id) |
| `<ExtendedTooltip>` | ref | Расширенная подсказка (name + id) |
| `<Events>` | block | Обработчики событий элемента |

### 7.3. Мультиязычный формат (multilang)

```xml
<Title>
  <v8:item>
    <v8:lang>ru</v8:lang>
    <v8:content>Текст на русском</v8:content>
  </v8:item>
</Title>
```

Атрибут `formatted="true"` на `<Title>` означает форматированную строку.

---

## 8. Типы UI-элементов — полное описание

### 8.1. UsualGroup — группа элементов

Основной контейнер для компоновки. Используется в ~90% форм.

```xml
<UsualGroup name="..." id="...">
  <!-- Компоновка -->
  <Group>Vertical | Horizontal | AlwaysHorizontal | AlwaysVertical</Group>
  <Behavior>Usual | Collapsible | CommandBar</Behavior>
  <Representation>None | NormalSeparation | WeakSeparation | StrongSeparation</Representation>
  <ShowTitle>true | false</ShowTitle>
  <United>true | false</United>

  <!-- Расположение дочерних -->
  <ChildItemsWidth>LeftWidest | RightWidest | Equal</ChildItemsWidth>
  <HorizontalSpacing>Single | Half | Double</HorizontalSpacing>
  <VerticalSpacing>Single | Half | Double</VerticalSpacing>
  <ThroughAlign>Use | DontUse</ThroughAlign>

  <!-- Внешний вид -->
  <BackColor>style:... | web:... | win:...</BackColor>
  <TextColor>style:...</TextColor>
  <TitleTextColor>style:...</TitleTextColor>
  <EnableContentChange>true | false</EnableContentChange>
  <ControlRepresentation>Picture | Text</ControlRepresentation>

  <ChildItems>...</ChildItems>
</UsualGroup>
```

### 8.2. InputField — поле ввода

Основной элемент ввода данных. Используется в ~80% форм.

```xml
<InputField name="..." id="...">
  <DataPath>Объект.Организация</DataPath>

  <!-- Заголовок -->
  <TitleLocation>Left | Right | Top | Bottom | None</TitleLocation>
  <TitleHeight>N</TitleHeight>
  <TitleWidth>N</TitleWidth>

  <!-- Размеры -->
  <AutoMaxWidth>true | false</AutoMaxWidth>
  <AutoMaxHeight>true | false</AutoMaxHeight>

  <!-- Режим редактирования -->
  <EditMode>Enter | EnterOnInput</EditMode>
  <MultiLine>true | false</MultiLine>
  <Wrap>true | false</Wrap>
  <ExtendedEdit>true | false</ExtendedEdit>
  <PasswordMode>true | false</PasswordMode>
  <DefaultItem>true | false</DefaultItem>

  <!-- Кнопки -->
  <ChoiceButton>true | false</ChoiceButton>
  <ChoiceButtonRepresentation>ShowInInputField | ShowInToolbar | Auto</ChoiceButtonRepresentation>
  <OpenButton>true | false</OpenButton>
  <ClearButton>true | false</ClearButton>
  <SpinButton>true | false</SpinButton>
  <CreateButton>true | false</CreateButton>
  <DropListButton>true | false</DropListButton>
  <TextEdit>true | false</TextEdit>
  <ListChoiceMode>true | false</ListChoiceMode>

  <!-- Автозаполнение и проверка -->
  <AutoMarkIncomplete>true | false</AutoMarkIncomplete>
  <MarkIncomplete>true | false</MarkIncomplete>
  <AutoComplete>true | false</AutoComplete>
  <QuickChoice>true | false</QuickChoice>
  <ChoiceHistoryOnInput>Auto | Never | Always</ChoiceHistoryOnInput>

  <!-- Подсказка ввода -->
  <InputHint>
    <v8:item>
      <v8:lang>ru</v8:lang>
      <v8:content>Placeholder text</v8:content>
    </v8:item>
  </InputHint>
  <Mask>маска ввода</Mask>

  <!-- Параметры выбора -->
  <ChoiceFoldersAndItems>Items | Folders | FoldersAndItems</ChoiceFoldersAndItems>
  <ChoiceParameters>
    <v8:Parameter>
      <v8:Name>name</v8:Name>
      <v8:Value>value</v8:Value>
    </v8:Parameter>
  </ChoiceParameters>

  <!-- Стилизация -->
  <TextColor>style:... | web:... | win:...</TextColor>
  <BackColor>style:... | web:... | win:...</BackColor>
  <BorderColor>style:... | web:... | win:...</BorderColor>
  <Font>...</Font>

  <!-- События -->
  <Events>
    <Event name="OnChange">...</Event>
    <Event name="StartChoice">...</Event>
    <Event name="ChoiceProcessing">...</Event>
    <Event name="Clearing">...</Event>
    <Event name="AutoComplete">...</Event>
    <Event name="TextEditEnd">...</Event>
    <Event name="Opening">...</Event>
    <Event name="OnEditEnd">...</Event>
    <Event name="DragCheck">...</Event>
    <Event name="Drag">...</Event>
    <Event name="DragStart">...</Event>
  </Events>
</InputField>
```

### 8.3. Button — кнопка

```xml
<Button name="..." id="...">
  <Type>CommandBarButton | UsualButton | Hyperlink</Type>
  <CommandName>Form.Command.Name | Form.StandardCommand.Name | CommonCommand.Name</CommandName>
  <DataPath>Attribute</DataPath>

  <Picture>
    <xr:Ref>StdPicture.Name | CommonPicture.Name</xr:Ref>
    <xr:LoadTransparent>true | false</xr:LoadTransparent>
  </Picture>
  <Representation>Auto | Picture | Text | PictureAndText</Representation>
  <ShapeRepresentation>Auto | None | Button</ShapeRepresentation>

  <DefaultButton>true | false</DefaultButton>
  <LocationInCommandBar>InCommandBar | InAdditionalSubmenu | InCommandBarAndInAdditionalSubmenu | Auto</LocationInCommandBar>
  <OnlyInAllActions>true | false</OnlyInAllActions>

  <Events>
    <Event name="Click">...</Event>
  </Events>
</Button>
```

### 8.4. Table — таблица

```xml
<Table name="..." id="...">
  <DataPath>ТабличныйРеквизит</DataPath>
  <RowPictureDataPath>ТабличныйРеквизит.Иконка</RowPictureDataPath>
  <RowsPicture>
    <xr:Ref>CommonPicture.Name</xr:Ref>
  </RowsPicture>

  <!-- Отображение -->
  <Representation>List | Tree | HierarchicalList</Representation>
  <TitleLocation>Top | None</TitleLocation>
  <HeightInTableRows>N</HeightInTableRows>
  <Header>true | false</Header>
  <Footer>true | false</Footer>
  <HorizontalLines>true | false</HorizontalLines>
  <VerticalLines>true | false</VerticalLines>
  <UseAlternationRowColor>true | false</UseAlternationRowColor>

  <!-- Редактирование -->
  <SelectionMode>SingleRow | MultiRow</SelectionMode>
  <ChangeRowSet>true | false</ChangeRowSet>
  <ChangeRowOrder>true | false</ChangeRowOrder>
  <AutoInsertNewRow>true | false</AutoInsertNewRow>

  <!-- Панели -->
  <CommandBarLocation>None | Top | Bottom | Auto</CommandBarLocation>
  <SearchStringLocation>None | Top | Bottom | CommandBar | Auto</SearchStringLocation>
  <ViewStatusLocation>Top | Bottom | None</ViewStatusLocation>
  <SearchControlLocation>Top | Bottom | Auto</SearchControlLocation>

  <!-- D&D -->
  <EnableStartDrag>true | false</EnableStartDrag>
  <EnableDrag>true | false</EnableDrag>
  <FileDragMode>AsFile | AsFileRef</FileDragMode>

  <!-- Дерево -->
  <TopLevelParent xsi:nil="true"/>
  <ShowRoot>true | false</ShowRoot>
  <AllowRootChoice>true | false</AllowRootChoice>
  <ChoiceFoldersAndItems>Items | Folders | FoldersAndItems</ChoiceFoldersAndItems>

  <!-- Обновление -->
  <AutoRefresh>true | false</AutoRefresh>
  <AutoRefreshPeriod>seconds</AutoRefreshPeriod>
  <UpdateOnDataChange>Auto | DontUpdate</UpdateOnDataChange>

  <!-- Исключённые команды таблицы -->
  <CommandSet>
    <ExcludedCommand>...</ExcludedCommand>
  </CommandSet>

  <!-- Служебные элементы -->
  <ContextMenu name="..." id="..."/>
  <AutoCommandBar name="..." id="..."/>
  <SearchStringAddition name="..." id="..."/>
  <ViewStatusAddition name="..." id="..."/>
  <SearchControlAddition name="..." id="..."/>

  <!-- Колонки -->
  <ChildItems>
    <!-- InputField, LabelField, CheckBoxField, PictureField -->
  </ChildItems>

  <!-- События -->
  <Events>
    <Event name="Selection">...</Event>
    <Event name="OnActivateRow">...</Event>
    <Event name="BeforeRowChange">...</Event>
    <Event name="BeforeAddRow">...</Event>
    <Event name="BeforeDeleteRow">...</Event>
    <Event name="AfterDeleteRow">...</Event>
    <Event name="DragStart">...</Event>
    <Event name="Drag">...</Event>
    <Event name="DragCheck">...</Event>
    <Event name="Drop">...</Event>
  </Events>
</Table>
```

### 8.5. Pages / Page — вкладки

```xml
<Pages name="..." id="...">
  <PagesRepresentation>None | TabsOnTop | TabsOnBottom | TabsOnLeft | TabsOnRight</PagesRepresentation>

  <Events>
    <Event name="OnCurrentPageChange">...</Event>
  </Events>

  <ChildItems>
    <Page name="..." id="...">
      <Title>...</Title>
      <Picture>
        <xr:Ref>StdPicture.Name</xr:Ref>
      </Picture>
      <ShowTitle>true | false</ShowTitle>
      <ChildItemsWidth>LeftWidest | RightWidest | Equal</ChildItemsWidth>

      <ChildItems>
        <!-- Любые UI-элементы -->
      </ChildItems>
    </Page>
  </ChildItems>
</Pages>
```

### 8.6. CommandBar — командная панель

```xml
<CommandBar name="..." id="...">
  <CommandSource>Form | FormCommandPanelGlobalCommands</CommandSource>
  <Autofill>true | false</Autofill>
  <EnableContentChange>true | false</EnableContentChange>
  <HorizontalLocation>Left | Right</HorizontalLocation>
  <VerticalLocation>Top | Bottom</VerticalLocation>

  <ChildItems>
    <!-- Button, ButtonGroup, Popup -->
  </ChildItems>
</CommandBar>
```

### 8.7. ButtonGroup — группа кнопок

```xml
<ButtonGroup name="..." id="...">
  <Representation>Auto | Compact | Separate</Representation>
  <CommandSource>Form | FormCommandPanelGlobalCommands | CommandPanel</CommandSource>

  <ChildItems>
    <!-- Button, ButtonGroup -->
  </ChildItems>
</ButtonGroup>
```

### 8.8. Popup — выпадающее меню

```xml
<Popup name="..." id="...">
  <Picture>
    <xr:Ref>StdPicture.Print</xr:Ref>
    <xr:LoadTransparent>true</xr:LoadTransparent>
  </Picture>
  <Representation>Auto | Picture | Text | PictureAndText</Representation>
  <ShapeRepresentation>Auto | None | Button</ShapeRepresentation>
  <LocationInCommandBar>InCommandBar | InAdditionalSubmenu | Auto</LocationInCommandBar>

  <ChildItems>
    <!-- Button, ButtonGroup, Popup -->
  </ChildItems>
</Popup>
```

### 8.9. LabelDecoration — декоративная надпись

```xml
<LabelDecoration name="..." id="...">
  <Title formatted="true">...</Title>
  <AutoMaxWidth>true | false</AutoMaxWidth>
  <AutoMaxHeight>true | false</AutoMaxHeight>
  <Hyperlink>true | false</Hyperlink>
  <ToolTipRepresentation>Auto | Button | None</ToolTipRepresentation>

  <TextColor>style:... | web:... | win:...</TextColor>
  <BackColor>style:... | web:... | win:...</BackColor>
  <Font>...</Font>
  <Border width="N">
    <v8ui:style xsi:type="v8ui:ControlBorderType">WithoutBorder | WithBorder</v8ui:style>
  </Border>

  <Events>
    <Event name="Click">...</Event>
  </Events>
</LabelDecoration>
```

### 8.10. LabelField — поле надписи (привязанное к данным)

```xml
<LabelField name="..." id="...">
  <DataPath>Реквизит.Свойство</DataPath>
  <TitleLocation>Left | Right | Top | Bottom | None</TitleLocation>
  <TitleTextColor>style:...</TitleTextColor>
  <PasswordMode>true | false</PasswordMode>
  <Hyperlink>true | false</Hyperlink>

  <TextColor>style:... | web:... | win:...</TextColor>
  <BackColor>style:... | web:... | win:...</BackColor>
  <Font>...</Font>

  <Events>
    <Event name="Click">...</Event>
    <Event name="URLProcessing">...</Event>
  </Events>
</LabelField>
```

### 8.11. CheckBoxField — флажок

```xml
<CheckBoxField name="..." id="...">
  <DataPath>Реквизит.Свойство</DataPath>
  <TitleLocation>Left | Right | Top | Bottom | None</TitleLocation>
  <CheckBoxType>Auto | Checkbox | Tumbler</CheckBoxType>
  <EditMode>Enter | EnterOnInput</EditMode>

  <Events>
    <Event name="OnChange">...</Event>
  </Events>
</CheckBoxField>
```

### 8.12. PictureDecoration — декоративная картинка

```xml
<PictureDecoration name="..." id="...">
  <Picture>
    <xr:Ref>StdPicture.Name | CommonPicture.Name</xr:Ref>
    <xr:LoadTransparent>true | false</xr:LoadTransparent>
  </Picture>
  <Zoomable>true | false</Zoomable>
  <NonselectedPictureText>текст</NonselectedPictureText>
  <Hyperlink>true | false</Hyperlink>
  <FileDragMode>AsFile | AsFileRef</FileDragMode>
  <DisplayImportance>Auto | VeryLow | Low | Normal | High | VeryHigh</DisplayImportance>

  <Border width="N">
    <v8ui:style xsi:type="v8ui:ControlBorderType">WithoutBorder | WithBorder</v8ui:style>
  </Border>

  <Events>
    <Event name="Click">...</Event>
  </Events>
</PictureDecoration>
```

### 8.13. PictureField — поле картинки (привязанное к данным)

```xml
<PictureField name="..." id="...">
  <DataPath>Реквизит.Свойство</DataPath>
  <TitleLocation>Left | Right | Top | Bottom | None</TitleLocation>
  <ValuesPicture>
    <xr:Ref>CommonPicture.Name</xr:Ref>
  </ValuesPicture>
  <Zoomable>true | false</Zoomable>
  <NonselectedPictureText>текст</NonselectedPictureText>
  <FileDragMode>AsFile | AsFileRef</FileDragMode>

  <Border width="N">
    <v8ui:style xsi:type="v8ui:ControlBorderType">WithoutBorder | WithBorder</v8ui:style>
  </Border>

  <Events>
    <Event name="Click">...</Event>
    <Event name="StartDrag">...</Event>
    <Event name="DragCheck">...</Event>
    <Event name="Drag">...</Event>
  </Events>
</PictureField>
```

### 8.14. CalendarField — календарь

```xml
<CalendarField name="..." id="...">
  <DataPath>Реквизит.Дата</DataPath>
  <TitleLocation>Left | Right | Top | Bottom | None</TitleLocation>
  <WidthInMonths>N</WidthInMonths>
  <HeightInWeeks>N</HeightInWeeks>
  <ShowCurrentDate>true | false</ShowCurrentDate>
  <BeginOfRepresentationPeriod>date</BeginOfRepresentationPeriod>
  <EndOfRepresentationPeriod>date</EndOfRepresentationPeriod>

  <Events>
    <Event name="Selection">...</Event>
    <Event name="OnPeriodOutput">...</Event>
  </Events>
</CalendarField>
```

---

## 9. Attributes — реквизиты формы

```xml
<Attributes>
  <Attribute name="ИмяРеквизита" id="N">
    <Title>...</Title>                          <!-- multilang, необязательный -->
    <ToolTip>...</ToolTip>                      <!-- multilang, необязательный -->
    <Type>...</Type>                            <!-- определение типа -->
    <MainAttribute>true</MainAttribute>         <!-- основной реквизит формы -->
    <SavedData>true</SavedData>                 <!-- сохраняемые данные -->
    <FillChecking>Show | DontShow</FillChecking>  <!-- проверка заполнения -->
    <UseAlwaysAttributes>true</UseAlwaysAttributes>
    <Columns>...</Columns>                      <!-- для ValueTable/ValueTree -->
  </Attribute>
</Attributes>
```

### 9.1. Система типов

#### Примитивные типы (xs:*)

```xml
<!-- Строка -->
<Type>
  <v8:Type>xs:string</v8:Type>
  <v8:StringQualifiers>
    <v8:Length>100</v8:Length>              <!-- 0 = неограниченная -->
    <v8:AllowedLength>Variable</v8:AllowedLength>  <!-- Variable | Fixed -->
  </v8:StringQualifiers>
</Type>

<!-- Число -->
<Type>
  <v8:Type>xs:decimal</v8:Type>
  <v8:NumberQualifiers>
    <v8:Digits>15</v8:Digits>              <!-- всего цифр -->
    <v8:FractionDigits>2</v8:FractionDigits>  <!-- дробная часть -->
    <v8:AllowedSign>Any</v8:AllowedSign>   <!-- Any | Nonnegative -->
  </v8:NumberQualifiers>
</Type>

<!-- Булево -->
<Type>
  <v8:Type>xs:boolean</v8:Type>
</Type>

<!-- Дата -->
<Type>
  <v8:Type>xs:dateTime</v8:Type>
  <v8:DateQualifiers>
    <v8:DateFractions>Date</v8:DateFractions>  <!-- Date | Time | DateTime -->
  </v8:DateQualifiers>
</Type>

<!-- Двоичные данные -->
<Type>
  <v8:Type>xs:binary</v8:Type>
  <v8:BinaryDataQualifiers>
    <v8:Length>0</v8:Length>
    <v8:AllowedLength>Variable</v8:AllowedLength>
  </v8:BinaryDataQualifiers>
</Type>
```

#### Ссылочные типы (cfg:*)

| Шаблон | Пример | Описание |
|--------|--------|----------|
| `cfg:CatalogRef.<Имя>` | `cfg:CatalogRef.Организации` | Ссылка на элемент справочника |
| `cfg:CatalogObject.<Имя>` | `cfg:CatalogObject.Контрагенты` | Объект справочника |
| `cfg:DocumentRef.<Имя>` | `cfg:DocumentRef.СчетФактура` | Ссылка на документ |
| `cfg:DocumentObject.<Имя>` | `cfg:DocumentObject.ПроцессПокупки` | Объект документа |
| `cfg:EnumRef.<Имя>` | `cfg:EnumRef.СпособДоставки` | Ссылка на перечисление |
| `cfg:ChartOfAccountsRef.<Имя>` | `cfg:ChartOfAccountsRef.Хозрасчетный` | Ссылка на план счетов |
| `cfg:ChartOfCalculationTypesRef.<Имя>` | — | Ссылка на план видов расчёта |
| `cfg:ChartOfCharacteristicTypesRef.<Имя>` | — | Ссылка на план видов характеристик |
| `cfg:ExchangePlanRef.<Имя>` | `cfg:ExchangePlanRef.АвтономнаяРабота` | Ссылка на план обмена |
| `cfg:BusinessProcessRef.<Имя>` | — | Ссылка на бизнес-процесс |
| `cfg:TaskRef.<Имя>` | — | Ссылка на задачу |
| `cfg:InformationRegisterRecordSet.<Имя>` | — | Набор записей регистра сведений |
| `cfg:InformationRegisterRecordManager.<Имя>` | — | Менеджер записи регистра сведений |
| `cfg:AccumulationRegisterRecordSet.<Имя>` | — | Набор записей регистра накопления |
| `cfg:AccountingRegisterRecordSet.<Имя>` | — | Набор записей регистра бухгалтерии |
| `cfg:ChartOfAccountsObject.<Имя>` | — | Объект плана счетов |
| `cfg:ChartOfCharacteristicTypesObject.<Имя>` | — | Объект ПВХ |
| `cfg:ChartOfCalculationTypesObject.<Имя>` | — | Объект плана видов расчёта |
| `cfg:ExchangePlanObject.<Имя>` | — | Объект плана обмена |
| `cfg:BusinessProcessObject.<Имя>` | — | Объект бизнес-процесса |
| `cfg:TaskObject.<Имя>` | — | Объект задачи |
| `cfg:ConstantsSet` | — | Набор констант |
| `cfg:DataProcessorObject.<Имя>` | — | Объект обработки |
| `cfg:ReportObject.<Имя>` | — | Объект отчёта |
| `cfg:DynamicList` | — | Динамический список |

#### Платформенные типы (v8:*)

| Тип | Описание |
|-----|----------|
| `v8:ValueListType` | Список значений |
| `v8:ValueTable` | Таблица значений |
| `v8:ValueTree` | Дерево значений |
| `v8:TypeDescription` | Описание типов |
| `v8:Universal` | Произвольный тип |
| `v8:FixedArray` | Фиксированный массив |
| `v8:FixedStructure` | Фиксированная структура |
| `v8:FillChecking` | Проверка заполнения |
| `v8:Null` | Null |
| `v8:StandardPeriod` | Стандартный период |
| `v8:StandardBeginningDate` | Стандартная начальная дата |
| `v8:Type` | Тип |
| `v8:UUID` | Уникальный идентификатор |

#### UI-типы (v8ui:*)

| Тип | Описание |
|-----|----------|
| `v8ui:FormattedString` | Форматированная строка |
| `v8ui:Picture` | Картинка |
| `v8ui:Color` | Цвет |
| `v8ui:Font` | Шрифт |
| `v8ui:SizeChangeMode` | Режим изменения размера |
| `v8ui:VerticalAlign` | Вертикальное выравнивание |
| `v8ui:HorizontalAlign` | Горизонтальное выравнивание |

#### Типы СКД (dcs*:*)

| Тип | Описание |
|-----|----------|
| `dcsset:DataCompositionSettings` | Настройки СКД |
| `dcssch:DataCompositionSchema` | Схема СКД |
| `dcscor:DataCompositionComparisonType` | Тип сравнения СКД |
| `dcsset:Filter` | Отбор СКД |
| `dcsset:SettingsComposer` | Компоновщик настроек |
| `dcsset:DataCompositionFieldPlacement` | Размещение поля СКД |
| `dcscor:DataCompositionGroupType` | Тип группировки |
| `dcscor:DataCompositionPeriodAdditionType` | Тип дополнения периода |
| `dcscor:DataCompositionSortDirection` | Направление сортировки |
| `dcscor:Field` | Поле СКД |

#### Типы предприятия (ent:*)

| Тип | Описание |
|-----|----------|
| `ent:AccountType` | Тип счёта (Активный/Пассивный/АктивноПассивный) |
| `ent:AccumulationRecordType` | Тип движения регистра накопления (Приход/Расход) |
| `ent:AccountingRecordType` | Тип бухгалтерской записи |

#### Пустой тип

```xml
<Type/>  <!-- нетипизированный / произвольный -->
```

### 9.2. Составные типы

Несколько типов в одном реквизите:

```xml
<Type>
  <v8:Type>cfg:CatalogRef.Организации</v8:Type>
  <v8:Type>cfg:CatalogRef.ИндивидуальныеПредприниматели</v8:Type>
  <v8:Type>cfg:CatalogRef.Контрагенты</v8:Type>
</Type>
```

### 9.3. ValueTable / ValueTree с колонками

```xml
<Attribute name="Строки" id="5">
  <Type>
    <v8:Type>v8:ValueTable</v8:Type>
  </Type>
  <Columns>
    <Column name="Номенклатура" id="1">
      <Title>...</Title>
      <Type>
        <v8:Type>cfg:CatalogRef.Номенклатура</v8:Type>
      </Type>
    </Column>
    <Column name="Количество" id="2">
      <Type>
        <v8:Type>xs:decimal</v8:Type>
        <v8:NumberQualifiers>
          <v8:Digits>10</v8:Digits>
          <v8:FractionDigits>3</v8:FractionDigits>
          <v8:AllowedSign>Nonnegative</v8:AllowedSign>
        </v8:NumberQualifiers>
      </Type>
    </Column>
  </Columns>
</Attribute>
```

---

## 10. Parameters — параметры формы

```xml
<Parameters>
  <Parameter name="ИмяПараметра">
    <Type>...</Type>                        <!-- идентично типам Attributes -->
    <KeyParameter>true</KeyParameter>       <!-- ключевой параметр -->
  </Parameter>
</Parameters>
```

Параметры **не имеют** атрибута `id`. Типы — те же, что для Attributes.

---

## 11. Commands — команды формы

```xml
<Commands>
  <Command name="ИмяКоманды" id="N">
    <Title>...</Title>                    <!-- multilang -->
    <ToolTip>...</ToolTip>                <!-- multilang -->
    <Picture>
      <xr:Ref>StdPicture.Refresh</xr:Ref>
      <xr:LoadTransparent>true</xr:LoadTransparent>
    </Picture>
    <Action>ИмяОбработчика</Action>       <!-- имя процедуры обработки -->
    <Shortcut>Ctrl+S</Shortcut>           <!-- клавиатурное сочетание -->
    <Representation>Auto | Picture | Text | PictureAndText | TextPicture | None | Compact</Representation>
    <CurrentRowUse>DontUse | Use | Auto</CurrentRowUse>
    <ModifiesData>true | false</ModifiesData>
    <ModifiesSavedData>true | false</ModifiesSavedData>
    <ChangedStateSavedData>true | false</ChangedStateSavedData>
    <Use>Auto</Use>
    <Mark>true | false</Mark>
    <ParameterUse>...</ParameterUse>
  </Command>
</Commands>
```

---

## 12. Ссылки на картинки

Два вида ссылок:

```xml
<!-- Стандартная картинка платформы -->
<xr:Ref>StdPicture.Refresh</xr:Ref>

<!-- Общая картинка конфигурации -->
<xr:Ref>CommonPicture.ЗаполнитьФорму</xr:Ref>
```

С прозрачностью:

```xml
<Picture>
  <xr:Ref>StdPicture.Print</xr:Ref>
  <xr:LoadTransparent>true</xr:LoadTransparent>
</Picture>
```

---

## 13. Ссылки на стили, цвета, шрифты

```xml
<!-- Стиль -->
<BackColor>style:FormBackColor</BackColor>

<!-- Web-цвет -->
<TextColor>web:Red</TextColor>

<!-- Windows-цвет -->
<BackColor>win:ButtonFace</BackColor>

<!-- Системный шрифт -->
<Font>sys:DefaultGUIFont</Font>
```

---

## 14. Рамки (Border)

```xml
<Border width="1">
  <v8ui:style xsi:type="v8ui:ControlBorderType">WithoutBorder</v8ui:style>
</Border>
```

Значения `ControlBorderType`: `WithoutBorder`, `WithBorder`.

---

## 15. DataPath — привязка к данным

Формат пути:

| Пример | Описание |
|--------|----------|
| `Объект.Организация` | Реквизит основного объекта формы |
| `Объект.Товары.Номенклатура` | Колонка табличной части объекта |
| `Отчет.НачалоПериода` | Параметр отчёта |
| `Запись.ОКОФ` | Поле записи регистра |
| `ТекстСообщения` | Реквизит формы верхнего уровня |

---

## 16. Статистика использования элементов

| Тип элемента | Частота |
|--------------|---------|
| UsualGroup | ~90% форм |
| Button | ~85% |
| InputField | ~80% |
| LabelDecoration | ~75% |
| CommandBar | ~70% |
| Table | ~60% |
| LabelField | ~60% |
| Pages / Page | ~55% |
| ButtonGroup | ~50% |
| CheckBoxField | ~45% |
| Popup | ~40% |
| PictureDecoration | ~40% |
| PictureField | ~15% |
| CalendarField | ~5% |

---

## 17. Элементы, не встреченные в конфигурации

Следующие элементы управления существуют в платформе, но не использованы в БП 3.0:

- `RadioButtonField`
- `TrackBarField`
- `ProgressBarField`
- `TextDocumentField`
- `SpreadSheetDocumentField`
- `HTMLDocumentField`
- `ChartField`
- `GanttChartField`
- `PlannerField`
- `GraphicalSchemaField`
- `FormattedDocumentField`


---

## Часть 2: Паттерны компоновки форм

# Паттерны компоновки управляемых форм

Рекомендации по дизайну форм, извлечённые из типовых конфигураций 1С. Используйте при создании форм через `/form-compile`, когда требования пользователя не детализируют расположение элементов.

## Архетипы форм

### Форма документа

```
Шапка (horizontal, 2 колонки)
├─ Левая (vertical): НомерДата (H: Номер + Дата "от"), Контрагент, Договор
├─ Правая (vertical): Организация, Подразделение, ЦеныИВалюта (надпись-ссылка)
Страницы (pages)
├─ Товары: таблица Объект.Товары
├─ Услуги: таблица Объект.Услуги (опционально)
└─ Дополнительно: прочие реквизиты
Подвал (vertical)
├─ Итоги (horizontal): Всего, НДС, Скидка
└─ КомментарийОтветственный (horizontal): Комментарий + Ответственный
```

**Типичные события:** OnCreateAtServer, OnReadAtServer, OnOpen, BeforeWriteAtServer, AfterWriteAtServer, AfterWrite, NotificationProcessing

**Свойства:** autoTitle=false, командная панель со стандартными + глобальными командами

### Форма обработки (DataProcessor)

```
Параметры (vertical)
├─ Группа полей ввода (Организация, Период, режимы работы)
├─ Информационные надписи (label, hyperlink)
Рабочая область
├─ Таблица данных или Pages с вкладками
Кнопки действий
├─ Выполнить / Применить (defaultButton)
├─ Закрыть (stdCommand: Close)
```

**Типичные события:** OnCreateAtServer, OnOpen, NotificationProcessing

**Свойства:** windowOpeningMode=LockOwnerWindow (если диалог), autoTitle=false

### Форма списка

```
Отборы (group: alwaysHorizontal)
├─ ГруппаОтбор[Поле] (H): Флажок + Поле ввода (для каждого фильтра)
Список (table, DynamicList)
├─ Колонки: labelField (не input — данные только для чтения)
```

**Типичные события:** OnCreateAtServer, OnOpen, NotificationProcessing, OnLoadDataFromSettingsAtServer

**Свойства:** autoSaveDataInSettings=Use (запомнить отборы)

**Фильтры:** пара реквизитов на каждый фильтр — `Отбор[Поле]` (значение) + `Отбор[Поле]Использование` (boolean, флажок вкл/выкл)

### Форма элемента справочника

**Простая:**
```
ГруппаРеквизитов (horizontal)
├─ Наименование -> Объект.Description
└─ Код -> Объект.Code (если нужен)
```

**Сложная:**
```
Главное (vertical)
├─ Наименование -> Объект.Description
├─ Параметры (horizontal, 2 колонки)
│  ├─ Левая: основные реквизиты
│  └─ Правая: дополнительные реквизиты
└─ КонтактныеДанные / Дополнительно (vertical)
```

**Типичные события:** OnCreateAtServer, OnReadAtServer, BeforeWriteAtServer, NotificationProcessing

### Мастер (Wizard)

```
Страницы (pages, OnCurrentPageChange)
├─ Шаг1: описание + параметры
├─ Шаг2: основная работа
└─ Шаг3: результат
Кнопки (horizontal)
├─ Назад (command), Далее (command, defaultButton), Выполнить (command)
└─ Закрыть (stdCommand: Close)
```

**Свойства:** windowOpeningMode=LockOwnerWindow

## Конвенции именования

### Группы

| Назначение | Имя | Тип |
|-----------|-----|-----|
| Шапка | `ГруппаШапка` | horizontal |
| Левая колонка | `ГруппаШапкаЛевая` | vertical |
| Правая колонка | `ГруппаШапкаПравая` | vertical |
| Номер+Дата | `ГруппаНомерДата` | horizontal |
| Подвал | `ГруппаПодвал` | vertical |
| Итоги | `ГруппаИтоги` | horizontal |
| Кнопки | `ГруппаКнопок` | horizontal |
| Страницы | `ГруппаСтраницы` / `Страницы` | pages |
| Предупреждение | `ГруппаПредупреждение` | horizontal, visible:false |
| Доп. секция | `ГруппаДополнительно` / `ГруппаПрочее` | vertical, collapse |

### Элементы

| Назначение | Имя | Суффикс |
|-----------|-----|---------|
| Поле в таблице | `[Таблица][Поле]` | — |
| Итог | `Итоги[Поле]` | — |
| Надпись-ссылка | `[Поле]Надпись` | — |
| Фильтр | `Отбор[Поле]` | — |
| Флажок фильтра | `Отбор[Поле]Использование` | — |
| Кнопка команды | `[Команда]Кнопка` | — |
| Баннер-картинка | `[Баннер]Картинка` | — |
| Баннер-надпись | `[Баннер]Надпись` | — |
| Подменю | `Подменю[Действие]` | — |

### Обработчики событий

Имя обработчика = имя элемента + суффикс события на русском:

| Событие | Суффикс | Пример |
|---------|---------|--------|
| OnChange | ПриИзменении | `ОрганизацияПриИзменении` |
| StartChoice | НачалоВыбора | `КонтрагентНачалоВыбора` |
| Click | Нажатие | `ЦеныИВалютаНажатие` |
| OnEditEnd | ПриОкончанииРедактирования | `ТоварыПриОкончанииРедактирования` |
| OnStartEdit | ПриНачалеРедактирования | `ТоварыПриНачалеРедактирования` |

Обработчики формы — стандартные имена: `ПриСозданииНаСервере`, `ПриОткрытии`, `ПередЗакрытием`, `ОбработкаОповещения`.

## Принципы компоновки

1. **Порядок чтения.** Сверху вниз, слева направо. Самое важное — вверху.
2. **Двухколоночная шапка.** Основные реквизиты слева (контрагент, склад), организационные справа (организация, подразделение).
3. **Кнопки действий внизу.** Главная кнопка — `defaultButton: true`. Закрыть — всегда последняя.
4. **Таблицы — основная область.** Табличные части занимают большую часть формы, обычно на Pages.
5. **Итоги рядом с таблицей.** В подвале, горизонтальная группа, все поля readOnly.
6. **Фильтры — отдельная зона.** Над списком, горизонтальная группа (alwaysHorizontal), пара "флажок + поле" на каждый фильтр.
7. **Скрытые элементы для состояний.** Баннеры, предупреждения — `visible: false` по умолчанию, показываются программно.
8. **Надписи-ссылки для диалогов.** `labelField` с `hyperlink: true` и событием Click — для открытия подформ (ЦеныИВалюта, УчётнаяПолитика).

## Продвинутые паттерны (ERP)

Извлечены из конфигурации «Управление предприятием» (ERP 8.3.24). Применяйте в сложных формах.

### Сворачиваемые группы (Collapsible)

Для необязательных секций — «Подписи», «Дополнительно», «Прочее». Сворачиваются по умолчанию, экономят место.

```
ГруппаПодписи (vertical, collapse, collapsed)
├─ Руководитель -> Объект.Руководитель
└─ ГлавныйБухгалтер -> Объект.ГлавныйБухгалтер
```

DSL:
```json
{ "group": "vertical", "name": "ГруппаПодписи", "title": "Подписи",
  "behavior": "Collapsible", "collapsed": true, "children": [
    { "input": "Руководитель", "path": "Объект.Руководитель" },
    { "input": "ГлавныйБухгалтер", "path": "Объект.ГлавныйБухгалтер" }
]}
```

### Баннер-предупреждение (Status Banner)

Группа «картинка + надпись» без заголовка, скрыта по умолчанию. Показывается программно при определённых условиях (просрочка, блокировка, информация).

```
ГруппаПредупреждение (horizontal, showTitle:false, visible:false)
├─ [Picture] ПредупреждениеКартинка -> StdPicture.Information
└─ [Label] ПредупреждениеНадпись (maxWidth:76, textColor:style:ПоясняющийТекст)
```

DSL:
```json
{ "group": "horizontal", "name": "ГруппаПредупреждение", "showTitle": false,
  "visible": false, "children": [
    { "picture": "ПредупреждениеКартинка" },
    { "label": "ПредупреждениеНадпись", "title": "Текст предупреждения",
      "maxWidth": 76, "autoMaxWidth": false }
]}
```

### Выпадающее меню в командной панели (Popup)

Группировка связанных команд (печать, отправка, выгрузка) в одну кнопку-меню с иконкой.

```
[CmdBar] КоманднаяПанель
├─ [Popup] ПодменюПечать (picture: StdPicture.Print, representation: Picture)
│  ├─ [Button] ПечатьНакладная -> Печать [cmd]
│  └─ [Button] ПечатьСчёт -> ПечатьСчёт [cmd]
└─ [Popup] ПодменюОтправить (picture: StdPicture.SendByEmail)
   └─ [Button] ОтправитьПоПочте -> Отправить [cmd]
```

DSL:
```json
{ "cmdBar": "КоманднаяПанель", "children": [
    { "popup": "ПодменюПечать", "title": "Печать",
      "picture": "StdPicture.Print", "representation": "Picture", "children": [
        { "button": "ПечатьНакладная", "command": "Печать" },
        { "button": "ПечатьСчёт", "command": "ПечатьСчёт" }
    ]},
    { "popup": "ПодменюОтправить", "title": "Отправить",
      "picture": "StdPicture.SendByEmail", "representation": "Picture", "children": [
        { "button": "ОтправитьПоПочте", "command": "Отправить" }
    ]}
]}
```

### Форма без стандартной командной панели

Для модальных диалогов и мастеров — отключение стандартной командной панели, полностью ручное управление кнопками.

```
properties: commandBarLocation=None, windowOpeningMode=LockWholeInterface
Содержимое (vertical)
├─ ... рабочая область ...
ГруппаКнопок (horizontal)
├─ Назад (command), Далее (command, defaultButton)
└─ Закрыть (stdCommand: Close)
```

DSL:
```json
{
  "properties": { "commandBarLocation": "None", "windowOpeningMode": "LockWholeInterface" },
  "elements": [
    { "group": "vertical", "name": "Содержимое", "children": [ "..." ] },
    { "group": "horizontal", "name": "ГруппаКнопок", "children": [
      { "button": "Назад", "command": "Назад" },
      { "button": "Далее", "command": "Далее", "defaultButton": true },
      { "button": "Закрыть", "stdCommand": "Close" }
    ]}
  ]
}
```

### Надпись-гиперссылка для открытия подформ

`labelField` с `hyperlink: true` и событием Click — вместо кнопки. Типичный приём для «ЦеныИВалюта», «УчётнаяПолитика» и подобных.

```
[LabelField] ЦеныИВалютаНадпись -> ЦеныИВалюта (hyperlink) {Click}
```

DSL:
```json
{ "labelField": "ЦеныИВалютаНадпись", "path": "ЦеныИВалюта",
  "hyperlink": true, "on": ["Click"] }
```

## Примеры DSL

### Типичная форма обработки

```json
{
  "title": "Загрузка данных из CSV",
  "properties": { "autoTitle": false, "windowOpeningMode": "LockOwnerWindow" },
  "events": { "OnCreateAtServer": "ПриСозданииНаСервере" },
  "elements": [
    { "group": "vertical", "name": "ГруппаПараметры", "children": [
      { "input": "ФайлЗагрузки", "path": "ФайлЗагрузки", "title": "Файл", "clearButton": true, "horizontalStretch": true, "on": ["StartChoice"] },
      { "input": "Кодировка", "path": "Кодировка", "title": "Кодировка" },
      { "input": "Разделитель", "path": "Разделитель", "title": "Разделитель колонок" }
    ]},
    { "table": "Данные", "path": "Объект.Данные", "on": ["OnStartEdit"], "columns": [
      { "input": "ДанныеНомерСтроки", "path": "Объект.Данные.LineNumber", "readOnly": true, "title": "№" },
      { "input": "ДанныеНаименование", "path": "Объект.Данные.Наименование" },
      { "input": "ДанныеКоличество", "path": "Объект.Данные.Количество", "on": ["OnChange"] },
      { "input": "ДанныеСумма", "path": "Объект.Данные.Сумма", "readOnly": true }
    ]},
    { "group": "horizontal", "name": "ГруппаКнопок", "children": [
      { "button": "Загрузить", "command": "Загрузить", "title": "Загрузить из файла", "defaultButton": true },
      { "button": "Очистить", "command": "Очистить", "title": "Очистить таблицу" },
      { "button": "Закрыть", "stdCommand": "Close" }
    ]}
  ],
  "attributes": [
    { "name": "Объект", "type": "ExternalDataProcessorObject.ЗагрузкаИзCSV", "main": true },
    { "name": "ФайлЗагрузки", "type": "string" },
    { "name": "Кодировка", "type": "string(20)" },
    { "name": "Разделитель", "type": "string(5)" }
  ],
  "commands": [
    { "name": "Загрузить", "action": "ЗагрузитьОбработка" },
    { "name": "Очистить", "action": "ОчиститьОбработка" }
  ]
}
```

### Типичная форма со списком и фильтрами

```json
{
  "properties": { "autoTitle": false, "autoSaveDataInSettings": "Use" },
  "events": {
    "OnCreateAtServer": "ПриСозданииНаСервере",
    "NotificationProcessing": "ОбработкаОповещения"
  },
  "elements": [
    { "group": "alwaysHorizontal", "name": "ГруппаОтборы", "children": [
      { "group": "horizontal", "name": "ГруппаОтборОрганизация", "children": [
        { "check": "ОтборОрганизацияИспользование", "path": "ОтборОрганизацияИспользование", "titleLocation": "none", "on": ["OnChange"] },
        { "input": "ОтборОрганизация", "path": "ОтборОрганизация", "title": "Организация", "on": ["OnChange"] }
      ]},
      { "group": "horizontal", "name": "ГруппаОтборКонтрагент", "children": [
        { "check": "ОтборКонтрагентИспользование", "path": "ОтборКонтрагентИспользование", "titleLocation": "none", "on": ["OnChange"] },
        { "input": "ОтборКонтрагент", "path": "ОтборКонтрагент", "title": "Контрагент", "on": ["OnChange"] }
      ]}
    ]},
    { "table": "Список", "path": "Список", "on": ["Selection", "OnActivateRow"], "columns": [
      { "labelField": "СписокДата", "path": "Список.Дата", "title": "Дата" },
      { "labelField": "СписокНомер", "path": "Список.Номер", "title": "Номер" },
      { "labelField": "СписокКонтрагент", "path": "Список.Контрагент" },
      { "labelField": "СписокСумма", "path": "Список.Сумма" }
    ]}
  ],
  "attributes": [
    { "name": "Список", "type": "DynamicList", "mainTable": "Document.РеализацияТоваров" },
    { "name": "ОтборОрганизация", "type": "CatalogRef.Организации" },
    { "name": "ОтборОрганизацияИспользование", "type": "boolean" },
    { "name": "ОтборКонтрагент", "type": "CatalogRef.Контрагенты" },
    { "name": "ОтборКонтрагентИспользование", "type": "boolean" }
  ]
}
```


---

## Часть 3: Спецификация Схемы компоновки данных (СКД)

# Спецификация XML-формата схемы компоновки данных 1С (DCS)

Спецификация формата `DataCompositionSchema` — макетов типа «Схема компоновки данных» в конфигурации 1С:Предприятие 8.3.
Составлена на основе анализа 930 схем конфигурации «Бухгалтерия предприятия 3.0.180» (платформа 8.3.24).

---

## 0. Файловая структура

### Два файла на каждую схему

```
<Объект>/Templates/
  ИмяМакета.xml                  ← метаданные (UUID, имя, TemplateType)
  ИмяМакета/
    Ext/
      Template.xml               ← тело схемы (DataCompositionSchema)
```

Типичные имена макетов: `ОсновнаяСхемаКомпоновкиДанных`, `СхемаКомпоновкиДанных`, произвольные.

### Метаданные макета — шаблон

```xml
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses"
    xmlns:v8="http://v8.1c.ru/8.1/data/core"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    version="2.17">
  <Template uuid="XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX">
    <Properties>
      <Name>ОсновнаяСхемаКомпоновкиДанных</Name>
      <Synonym>
        <v8:item>
          <v8:lang>ru</v8:lang>
          <v8:content>Основная схема компоновки данных</v8:content>
        </v8:item>
      </Synonym>
      <Comment/>
      <TemplateType>DataCompositionSchema</TemplateType>
    </Properties>
  </Template>
</MetaDataObject>
```

Значение `TemplateType` для DCS всегда: **`DataCompositionSchema`**.

### Где встречаются DCS-макеты

| Тип объекта метаданных | Частота | Примечание |
|---|---|---|
| Reports (Отчёты) | ~420 | Основное место — каждый отчёт СКД |
| DataProcessors (Обработки) | ~11 | Обработки с отчётными функциями |
| Enums (Перечисления) | ~20 | Дополнительные ссылки |
| Catalogs (Справочники) | ~5 | Запросы к справочным данным |
| DocumentJournals | ~4 | Журналы документов |
| CommonTemplates | ~3 | Общие макеты |
| InformationRegisters | ~2 | Регистры сведений |
| Documents (Документы) | ~1 | Редко |

---

## 1. Пространства имён

Корневой элемент — `<DataCompositionSchema>`.

```xml
<DataCompositionSchema
    xmlns="http://v8.1c.ru/8.1/data-composition-system/schema"
    xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common"
    xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core"
    xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings"
    xmlns:v8="http://v8.1c.ru/8.1/data/core"
    xmlns:v8ui="http://v8.1c.ru/8.1/data/ui"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
```

| Префикс | URI | Назначение |
|---|---|---|
| *(default)* | `.../data-composition-system/schema` | Элементы схемы (dataSource, dataSet, field, parameter, ...) |
| `dcscom` | `.../data-composition-system/common` | Общие типы СКД (dimension, account, role, ...) |
| `dcscor` | `.../data-composition-system/core` | Ядро СКД (Field, SettingsParameterValue, ChoiceParameterLinks, ...) |
| `dcsset` | `.../data-composition-system/settings` | Настройки варианта (selection, filter, order, group, ...) |
| `v8` | `.../data/core` | Типы данных ядра (LocalStringType, Type, StandardPeriod, ...) |
| `v8ui` | `.../data/ui` | UI-типы (HorizontalAlign, ...) |
| `xs` | `.../XMLSchema` | Стандартные XSD-типы (string, dateTime, boolean, decimal, ...) |
| `xsi` | `.../XMLSchema-instance` | Атрибуты экземпляра (xsi:type, xsi:nil) |

Дополнительные пространства имён (появляются в `settingsVariant`):

| Префикс | URI | Где |
|---|---|---|
| `style` | `http://v8.1c.ru/8.1/data/ui/style` | В settings — стили оформления |
| `sys` | `http://v8.1c.ru/8.1/data/ui/fonts/system` | В settings — системные шрифты |
| `web` | `http://v8.1c.ru/8.1/data/ui/colors/web` | В settings — веб-цвета |
| `win` | `http://v8.1c.ru/8.1/data/ui/colors/windows` | В settings — цвета Windows |

---

## 2. Общая структура DataCompositionSchema

Элементы верхнего уровня (порядок фиксирован):

```
DataCompositionSchema
├── dataSource*              — источники данных (раздел 3)
├── dataSet*                 — наборы данных (раздел 4)
├── dataSetLink*             — связи между наборами (раздел 5)
├── calculatedField*         — вычисляемые поля (раздел 6)
├── totalField*              — итоговые поля (раздел 7)
├── parameter*               — параметры схемы (раздел 8)
├── template*                — макеты областей (раздел 9)
├── groupTemplate*           — привязки макетов группировок (раздел 10)
├── settingsVariant*         — варианты настроек (раздел 11)
```

`*` — 0..N элементов.

Минимальная DCS содержит: 1 dataSource + 1 dataSet + 1 settingsVariant.

---

## 3. Источники данных (dataSource)

```xml
<dataSource>
  <name>ИсточникДанных1</name>
  <dataSourceType>Local</dataSourceType>
</dataSource>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `name` | да | Уникальное имя, на которое ссылаются наборы данных |
| `dataSourceType` | да | Тип: `Local` (текущая информационная база) или `External` (внешний) |

В подавляющем большинстве случаев — один источник `Local`. Имя произвольное: `ИсточникДанных1`, `ИнформационнаяБаза` и т.п.

---

## 4. Наборы данных (dataSet)

Тип набора определяется атрибутом `xsi:type`. Три типа:

### 4.1. DataSetQuery — запрос

Самый распространённый тип. Содержит SQL-подобный запрос на языке 1С.

```xml
<dataSet xsi:type="DataSetQuery">
  <name>НаборДанных1</name>
  <field xsi:type="DataSetFieldField">...</field>   <!-- 0..N полей -->
  <dataSource>ИсточникДанных1</dataSource>
  <query>ВЫБРАТЬ ... ИЗ ...</query>
  <autoFillFields>false</autoFillFields>             <!-- опционально -->
</dataSet>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `name` | да | Уникальное имя набора |
| `field` | нет | Описания полей (раздел 4.4) |
| `dataSource` | да | Ссылка на имя dataSource |
| `query` | да | Текст запроса на языке 1С (XML-экранирование: `&amp;` для `&`, `&gt;` для `>`) |
| `autoFillFields` | нет | `false` — отключить автозаполнение полей из запроса (по умолчанию `true`) |

#### Особенности запросов в DCS

- Параметры: `&ИмяПараметра` (в XML: `&amp;ИмяПараметра`)
- Авторазметка полей в фигурных скобках: `{ВЫБРАТЬ ...}`, `{ГДЕ ...}`, `{ЛЕВОЕ СОЕДИНЕНИЕ ...}` — позволяют СКД автоматически модифицировать запрос
- Пакетные запросы: несколько запросов через `; ////////////////`
- Временные таблицы: `ПОМЕСТИТЬ ИмяВТ`, `ИНДЕКСИРОВАТЬ ПО`

### 4.2. DataSetObject — объект

Данные берутся из программно заполненной таблицы значений.

```xml
<dataSet xsi:type="DataSetObject">
  <name>НаборДанных1</name>
  <field xsi:type="DataSetFieldField">...</field>
  <dataSource>ИсточникДанных1</dataSource>
  <objectName>ТаблицаПроверки</objectName>
</dataSet>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `objectName` | да | Имя объекта (таблицы значений), передаваемого программно |

### 4.3. DataSetUnion — объединение

Объединяет поля из нескольких наборов. Сам не содержит запросов — объединяет подчинённые наборы.

```xml
<dataSet xsi:type="DataSetUnion">
  <name>РасчетНалога</name>
  <field xsi:type="DataSetFieldField">...</field>   <!-- агрегированные поля -->
  <item xsi:type="DataSetQuery">                     <!-- вложенные наборы -->
    <name>ДанныеПоСтоимости</name>
    ...
  </item>
  <item xsi:type="DataSetQuery">
    <name>ДанныеПоКадастру</name>
    ...
  </item>
</dataSet>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `field` | нет | Поля объединения (описывают результирующие колонки) |
| `item` | да | Вложенные наборы (DataSetQuery или другие) |

### 4.4. Поля набора данных (field)

Каждое поле — элемент `<field xsi:type="DataSetFieldField">`:

```xml
<field xsi:type="DataSetFieldField">
  <dataPath>ОстаточнаяСтоимость</dataPath>
  <field>ОстаточнаяСтоимость</field>
  <title xsi:type="v8:LocalStringType">
    <v8:item>
      <v8:lang>ru</v8:lang>
      <v8:content>Остаточная стоимость</v8:content>
    </v8:item>
  </title>
  <useRestriction>
    <condition>true</condition>
  </useRestriction>
  <role>
    <dcscom:dimension>true</dcscom:dimension>
  </role>
  <valueType>
    <v8:Type>xs:string</v8:Type>
    <v8:StringQualifiers>
      <v8:Length>11</v8:Length>
      <v8:AllowedLength>Variable</v8:AllowedLength>
    </v8:StringQualifiers>
  </valueType>
  <appearance>
    <dcscor:item xsi:type="dcsset:SettingsParameterValue">
      <dcscor:parameter>Формат</dcscor:parameter>
      <dcscor:value xsi:type="xs:string">ЧДЦ=2</dcscor:value>
    </dcscor:item>
  </appearance>
  <inputParameters>...</inputParameters>
  <presentationExpression>...</presentationExpression>
</field>
```

#### Элементы поля

| Элемент | Обязат. | Описание |
|---|---|---|
| `dataPath` | да | Путь к данным (имя поля в результате СКД). Через точку — реквизиты: `Номенклатура.Артикул` |
| `field` | да | Имя поля в запросе (может отличаться от dataPath) |
| `title` | нет | Локализованный заголовок (`v8:LocalStringType`) |
| `useRestriction` | нет | Ограничения использования поля (раздел 4.5) |
| `attributeUseRestriction` | нет | Ограничения использования реквизитов поля (раздел 4.5) |
| `role` | нет | Роль поля в СКД (раздел 4.6) |
| `valueType` | нет | Тип значения поля (раздел 4.7) |
| `appearance` | нет | Оформление — список параметров `dcscor:item` (раздел 4.8) |
| `inputParameters` | нет | Параметры ввода / связи параметров выбора (раздел 4.9) |
| `presentationExpression` | нет | Выражение для формирования представления (на языке 1С) |

### 4.5. Ограничения использования поля (useRestriction / attributeUseRestriction)

```xml
<useRestriction>
  <field>true</field>         <!-- запрет использования как поле в выборке -->
  <condition>true</condition>  <!-- запрет в условиях отбора -->
  <group>true</group>          <!-- запрет в группировках -->
  <order>true</order>          <!-- запрет в сортировке -->
</useRestriction>
```

Каждый подэлемент — `true`/`false` (по умолчанию `false` = разрешено). Можно указывать подмножество.

`attributeUseRestriction` — аналогичная структура, применяется к реквизитам (дочерним полям) поля.

### 4.6. Роли полей (role)

```xml
<role>
  <dcscom:dimension>true</dcscom:dimension>          <!-- поле — измерение -->
  <dcscom:account>true</dcscom:account>               <!-- поле — счёт -->
  <dcscom:accountTypeExpression>Счет.Вид</dcscom:accountTypeExpression>  <!-- выражение типа счёта -->
</role>
```

| Подэлемент | Описание |
|---|---|
| `dcscom:dimension` | Поле является измерением (`true`/`false`) |
| `dcscom:account` | Поле является счётом |
| `dcscom:accountTypeExpression` | Выражение для определения типа счёта |
| `dcscom:balance` | Поле является остатком |
| `dcscom:balanceGroup` | Группа остатка |
| `dcscom:periodNumber` | Номер периода (обычно `1`) |
| `dcscom:periodType` | Тип периода (`Main`, `Additional`) |

### 4.7. Тип значения (valueType)

```xml
<valueType>
  <v8:Type>xs:string</v8:Type>
  <v8:StringQualifiers>
    <v8:Length>11</v8:Length>
    <v8:AllowedLength>Variable</v8:AllowedLength>
  </v8:StringQualifiers>
</valueType>
```

Типы: `xs:string`, `xs:dateTime`, `xs:decimal`, `xs:boolean`, ссылочные типы конфигурации.

Ссылочные типы объявляются с inline namespace на элементе `<v8:Type>`:

```xml
<v8:Type xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config">d5p1:CatalogRef.Номенклатура</v8:Type>
```

Префикс (`d5p1`, `d4p1` и т.д.) — автогенерируемый, суть в URI `http://v8.1c.ru/8.1/data/enterprise/current-config`. Поддерживаются: `CatalogRef`, `DocumentRef`, `EnumRef`, `ChartOfAccountsRef`, `ChartOfCharacteristicTypesRef` и др.

Квалификаторы:
- `v8:StringQualifiers` → `v8:Length`, `v8:AllowedLength` (Fixed/Variable)
- `v8:DateQualifiers` → `v8:DateFractions` (Date/Time/DateTime)
- `v8:NumberQualifiers` → `v8:Digits`, `v8:FractionDigits`, `v8:AllowedSign` (Any/Nonnegative)

### 4.8. Оформление полей (appearance)

Список параметров оформления:

```xml
<appearance>
  <dcscor:item xsi:type="dcsset:SettingsParameterValue">
    <dcscor:parameter>Формат</dcscor:parameter>
    <dcscor:value xsi:type="xs:string">ЧДЦ=2</dcscor:value>
  </dcscor:item>
  <dcscor:item xsi:type="dcsset:SettingsParameterValue">
    <dcscor:parameter>ГоризонтальноеПоложение</dcscor:parameter>
    <dcscor:value xsi:type="v8ui:HorizontalAlign">Center</dcscor:value>
  </dcscor:item>
</appearance>
```

Типичные параметры оформления:

| Параметр | Тип значения | Пример |
|---|---|---|
| `Формат` | `xs:string` | `ЧДЦ=2`, `ЧГ=0`, `ЧН=0`, `ДФ=dd.MM.yyyy`, `Л=ru; ДФ=ММММ` |
| `ГоризонтальноеПоложение` | `v8ui:HorizontalAlign` | `Left`, `Center`, `Right` |

### 4.9. Параметры ввода (inputParameters)

Связи параметров выбора для интерактивных полей:

```xml
<inputParameters>
  <dcscor:item>
    <dcscor:parameter>СвязиПараметровВыбора</dcscor:parameter>
    <dcscor:value xsi:type="dcscor:ChoiceParameterLinks">
      <dcscor:item>
        <dcscor:choiceParameter>Отбор.Владелец</dcscor:choiceParameter>
        <dcscor:value>Организация</dcscor:value>
        <dcscor:mode xmlns:d8p1="http://v8.1c.ru/8.1/data/enterprise"
                     xsi:type="d8p1:LinkedValueChangeMode">Clear</dcscor:mode>
      </dcscor:item>
    </dcscor:value>
  </dcscor:item>
</inputParameters>
```

Используется для каскадных зависимостей в пользовательских настройках (например, подразделение зависит от организации).

---

## 5. Связи между наборами данных (dataSetLink)

Позволяют передавать параметры из одного набора в другой:

```xml
<dataSetLink>
  <sourceDataSet>Периоды</sourceDataSet>
  <destinationDataSet>ДанныеТ13</destinationDataSet>
  <sourceExpression>НачалоМесяца</sourceExpression>
  <destinationExpression>Месяц</destinationExpression>
  <parameter>НачалоМесяца</parameter>
  <parameterListAllowed>false</parameterListAllowed>
</dataSetLink>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `sourceDataSet` | да | Имя набора-источника |
| `destinationDataSet` | да | Имя целевого набора |
| `sourceExpression` | да | Выражение из источника (поле или формула) |
| `destinationExpression` | да | Выражение для сопоставления в целевом наборе |
| `parameter` | нет | Имя параметра для передачи значения |
| `parameterListAllowed` | нет | Допустим ли список значений (`true`/`false`) |

---

## 6. Вычисляемые поля (calculatedField)

Поля, вычисляемые выражением на языке 1С (не из запроса):

```xml
<calculatedField>
  <dataPath>УИД</dataPath>
  <expression>БухгалтерскиеОтчеты.ПолучитьУИДСсылкиСтрокой(Номенклатура)</expression>
  <title xsi:type="v8:LocalStringType">
    <v8:item>
      <v8:lang>ru</v8:lang>
      <v8:content>Уникальный идентификатор</v8:content>
    </v8:item>
  </title>
  <useRestriction>
    <condition>true</condition>
    <group>true</group>
    <order>true</order>
  </useRestriction>
</calculatedField>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `dataPath` | да | Путь к полю в результате |
| `expression` | да | Выражение на языке 1С (может вызывать методы общих модулей) |
| `title` | нет | Локализованный заголовок |
| `useRestriction` | нет | Ограничения использования (аналогично полям) |
| `valueType` | нет | Тип значения |
| `appearance` | нет | Оформление |

---

## 7. Итоговые поля (totalField)

Агрегатные функции для подведения итогов:

```xml
<totalField>
  <dataPath>Количество</dataPath>
  <expression>Сумма(Количество)</expression>
</totalField>
<totalField>
  <dataPath>Цена</dataPath>
  <expression>Максимум(Цена)</expression>
</totalField>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `dataPath` | да | Путь к полю |
| `expression` | да | Агрегатная функция: `Сумма(...)`, `Количество(...)`, `Максимум(...)`, `Минимум(...)`, `Среднее(...)` |
| `group` | нет | Имя группировки, для которой считать итоги. Без `group` — для всех группировок |

### Разные формулы для разных группировок

Одно поле может иметь несколько записей `totalField` с разными формулами для разных группировок:

```xml
<!-- Для группировки "ОбъектМетаданных" — агрегация самого поля -->
<totalField>
  <dataPath>ПравоИнтерактивное</dataPath>
  <expression>Максимум(ПравоИнтерактивное)</expression>
  <group>ОбъектМетаданных</group>
</totalField>
<!-- Для группировки "Отчет" — агрегация другого поля -->
<totalField>
  <dataPath>ПравоИнтерактивное</dataPath>
  <expression>Максимум(ПравоОтчета)</expression>
  <group>Отчет</group>
</totalField>
```

Это позволяет вычислять ресурс по-разному в зависимости от контекста группировки.

---

## 8. Параметры схемы (parameter)

Параметры, доступные для задания пользователем или программно:

```xml
<parameter>
  <name>Период</name>
  <title xsi:type="v8:LocalStringType">
    <v8:item>
      <v8:lang>ru</v8:lang>
      <v8:content>Период</v8:content>
    </v8:item>
  </title>
  <valueType>
    <v8:Type>v8:StandardPeriod</v8:Type>
  </valueType>
  <value xsi:type="v8:StandardPeriod">
    <v8:variant xsi:type="v8:StandardPeriodVariant">LastMonth</v8:variant>
  </value>
  <useRestriction>false</useRestriction>
  <expression>&amp;Период.ДатаНачала</expression>
  <availableAsField>false</availableAsField>
  <use>Always</use>
</parameter>
```

| Элемент | Обязат. | Описание |
|---|---|---|
| `name` | да | Имя параметра (используется в запросах как `&ИмяПараметра`) |
| `title` | нет | Локализованный заголовок |
| `valueType` | нет | Тип значения (раздел 4.7) |
| `value` | нет | Значение по умолчанию |
| `useRestriction` | нет | `true` — параметр скрыт от пользователя, `false` — доступен |
| `expression` | нет | Выражение для автоматического вычисления (например, `&Период.ДатаНачала`) |
| `availableAsField` | нет | `false` — параметр недоступен как поле в отчёте |
| `valueListAllowed` | нет | `true` — разрешает передавать список значений в параметр |
| `use` | нет | Режим: `Always` (всегда), `Auto` (автоматически) |

### Типы значений параметров

| Тип | XML-тип | Пример value |
|---|---|---|
| Дата | `xs:dateTime` | `0001-01-01T00:00:00` |
| Строка | `xs:string` | `Т13` |
| Стандартный период | `v8:StandardPeriod` | `<v8:variant>LastMonth</v8:variant>` |
| Ссылка | `d5p1:CatalogRef.ИмяСправочника` (с `xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config"`) | `xsi:nil="true"` |
| null | — | `xsi:nil="true"` |

Стандартные варианты периодов (`v8:StandardPeriodVariant`): `Custom`, `Today`, `ThisWeek`, `ThisMonth`, `ThisQuarter`, `ThisYear`, `LastMonth`, `LastQuarter`, `LastYear` и др.

---

## 9. Макеты областей (template)

Пользовательские шаблоны вывода (макеты ячеек):

```xml
<template>
  <name>Макет1</name>
  <template xmlns:dcsat="http://v8.1c.ru/8.1/data-composition-system/area-template"
            xsi:type="dcsat:AreaTemplate">
    <dcsat:item xsi:type="dcsat:TableRow">
      <dcsat:tableCell>
        <dcsat:item xsi:type="dcsat:Field">
          <dcsat:value xsi:type="dcscor:Parameter">ТипЦены</dcsat:value>
        </dcsat:item>
      </dcsat:tableCell>
    </dcsat:item>
  </template>
  <parameter xmlns:dcsat="http://v8.1c.ru/8.1/data-composition-system/area-template"
             xsi:type="dcsat:ExpressionAreaTemplateParameter">
    <dcsat:name>ТипЦены</dcsat:name>
    <dcsat:expression>Представление(ТипЦен)</dcsat:expression>
  </parameter>
</template>
```

Пространство имён `dcsat`: `http://v8.1c.ru/8.1/data-composition-system/area-template`.

| Элемент | Описание |
|---|---|
| `name` | Имя макета (ссылаются groupTemplate) |
| `template` (вложенный) | Описание строк/ячеек (`dcsat:AreaTemplate`) |
| `parameter` (Expression) | Параметры макета (`dcsat:ExpressionAreaTemplateParameter`) — выражения для подстановки |
| `parameter` (Details) | Параметры расшифровки (`dcsat:DetailsAreaTemplateParameter`) — для drilldown |

#### DetailsAreaTemplateParameter

Параметр расшифровки — активирует drilldown при клике на ячейку:

```xml
<parameter xmlns:dcsat="http://v8.1c.ru/8.1/data-composition-system/area-template"
           xsi:type="dcsat:DetailsAreaTemplateParameter">
  <dcsat:name>Расшифровка_ПоступлениеСырья</dcsat:name>
  <dcsat:fieldExpression>
    <dcsat:field>ИмяРесурса</dcsat:field>
    <dcsat:expression>"ПоступлениеСырья"</dcsat:expression>
  </dcsat:fieldExpression>
  <dcsat:mainAction>DrillDown</dcsat:mainAction>
</parameter>
```

Привязка к ячейке — через appearance `Расшифровка`:

```xml
<dcscor:item>
  <dcscor:parameter>Расшифровка</dcscor:parameter>
  <dcscor:value xsi:type="dcscor:Parameter">Расшифровка_ПоступлениеСырья</dcscor:value>
</dcscor:item>
```

---

## 10. Привязки макетов группировок (groupTemplate, groupHeaderTemplate)

Связывают группировку с пользовательским макетом. Два XML-элемента:

- `<groupTemplate>` — шаблон строки данных (`Header`) и итогов (`OverallHeader`)
- `<groupHeaderTemplate>` — шаблон заголовка группировки (шапка таблицы)

```xml
<groupHeaderTemplate>
  <groupName>ДанныеОтчета</groupName>
  <templateType>Header</templateType>
  <template>Макет1</template>
</groupHeaderTemplate>
<groupTemplate>
  <groupField>Счет</groupField>
  <templateType>Header</templateType>
  <template>Макет2</template>
</groupTemplate>
```

| Элемент | Описание |
|---|---|
| `groupField` | Привязка к полю группировки |
| `groupName` | Привязка к именованной группировке в структуре варианта |
| `templateType` | `Header` (строки данных), `OverallHeader` (итоги) |
| `template` | Ссылка на имя template из раздела 9 |

---

## 11. Варианты настроек (settingsVariant)

Каждый вариант — именованная конфигурация отчёта. Отчёт может иметь несколько вариантов.

```xml
<settingsVariant>
  <dcsset:name>Основной</dcsset:name>
  <dcsset:presentation xsi:type="v8:LocalStringType">
    <v8:item>
      <v8:lang>ru</v8:lang>
      <v8:content>Основной вариант отчёта</v8:content>
    </v8:item>
  </dcsset:presentation>
  <dcsset:settings xmlns:style="..." xmlns:sys="..." xmlns:web="..." xmlns:win="...">
    <!-- содержимое настроек -->
  </dcsset:settings>
</settingsVariant>
```

### 11.1. Структура settings

```
dcsset:settings
├── dcsset:selection              — выбранные поля (раздел 11.2)
├── dcsset:filter                 — отборы (раздел 11.3)
├── dcsset:order                  — сортировка (раздел 11.4)
├── dcsset:conditionalAppearance  — условное оформление (раздел 11.5)
├── dcsset:outputParameters       — параметры вывода (раздел 11.6)
├── dcsset:dataParameters         — значения параметров данных (раздел 11.7)
├── dcsset:item*                  — элементы структуры (раздел 11.8)
```

### 11.2. Выборка полей (selection)

```xml
<dcsset:selection>
  <dcsset:item xsi:type="dcsset:SelectedItemField">
    <dcsset:field>ТипОбъекта</dcsset:field>
    <dcsset:lwsTitle>                          <!-- опциональный заголовок -->
      <v8:item>
        <v8:lang>ru</v8:lang>
        <v8:content>Наименование</v8:content>
      </v8:item>
    </dcsset:lwsTitle>
  </dcsset:item>
  <dcsset:item xsi:type="dcsset:SelectedItemAuto"/>   <!-- авто-выбор -->
</dcsset:selection>
```

Типы элементов выборки:
- `dcsset:SelectedItemField` — конкретное поле (элемент `dcsset:field`)
- `dcsset:SelectedItemAuto` — автоматический подбор полей

### 11.3. Отборы (filter)

```xml
<dcsset:filter>
  <dcsset:item xsi:type="dcsset:FilterItemComparison">
    <dcsset:use>false</dcsset:use>                    <!-- включён/выключен -->
    <dcsset:left xsi:type="dcscor:Field">Организация</dcsset:left>
    <dcsset:comparisonType>Equal</dcsset:comparisonType>
    <dcsset:right xsi:type="xs:boolean">false</dcsset:right>
    <dcsset:presentation xsi:type="v8:LocalStringType">
      <v8:item>
        <v8:lang>ru</v8:lang>
        <v8:content>Описание фильтра</v8:content>
      </v8:item>
    </dcsset:presentation>
    <dcsset:viewMode>Normal</dcsset:viewMode>
    <dcsset:userSettingID>GUID</dcsset:userSettingID>
  </dcsset:item>
</dcsset:filter>
```

Типы элементов фильтра:
- `dcsset:FilterItemComparison` — сравнение поля с значением
- `dcsset:FilterItemGroup` — группа условий (И/ИЛИ)

Типы сравнения (`comparisonType`):

| Значение | Описание |
|---|---|
| `Equal` | Равно |
| `NotEqual` | Не равно |
| `Greater` | Больше |
| `GreaterOrEqual` | Больше или равно |
| `Less` | Меньше |
| `LessOrEqual` | Меньше или равно |
| `InList` | В списке |
| `NotInList` | Не в списке |
| `InHierarchy` | В иерархии |
| `InListByHierarchy` | В списке по иерархии |
| `Contains` | Содержит |
| `NotContains` | Не содержит |
| `BeginsWith` | Начинается с |
| `NotBeginsWith` | Не начинается с |
| `Filled` | Заполнено |
| `NotFilled` | Не заполнено |

Значение правой части (`right`) — может содержать списки:
```xml
<dcsset:right xsi:type="v8:ValueListType">
  <v8:valueType/>
  <v8:lastId xsi:type="xs:decimal">-1</v8:lastId>
</dcsset:right>
```

### 11.4. Сортировка (order)

```xml
<dcsset:order>
  <dcsset:item xsi:type="dcsset:OrderItemField">
    <dcsset:field>РазмерДанных</dcsset:field>
    <dcsset:orderType>Desc</dcsset:orderType>
  </dcsset:item>
  <dcsset:item xsi:type="dcsset:OrderItemAuto"/>
</dcsset:order>
```

Типы элементов сортировки:
- `dcsset:OrderItemField` — по полю (`dcsset:field` + `dcsset:orderType`: `Asc`/`Desc`)
- `dcsset:OrderItemAuto` — автоматическая сортировка

### 11.5. Условное оформление (conditionalAppearance)

```xml
<dcsset:conditionalAppearance>
  <dcsset:item>
    <dcsset:selection>
      <dcsset:item>
        <dcsset:field>ИмяПоля</dcsset:field>
      </dcsset:item>
    </dcsset:selection>
    <dcsset:filter>
      <dcsset:item xsi:type="dcsset:FilterItemComparison">
        <dcsset:left xsi:type="dcscor:Field">ИмяПоля</dcsset:left>
        <dcsset:comparisonType>Equal</dcsset:comparisonType>
        <dcsset:right xsi:type="xs:decimal">0</dcsset:right>
      </dcsset:item>
    </dcsset:filter>
    <dcsset:appearance>
      <dcscor:item xsi:type="dcsset:SettingsParameterValue">
        <dcscor:parameter>Текст</dcscor:parameter>
        <dcscor:value xsi:type="xs:string"/>
      </dcscor:item>
    </dcsset:appearance>
  </dcsset:item>
</dcsset:conditionalAppearance>
```

### 11.6. Параметры вывода (outputParameters)

```xml
<dcsset:outputParameters>
  <dcscor:item xsi:type="dcsset:SettingsParameterValue">
    <dcscor:use>false</dcscor:use>                     <!-- опционально -->
    <dcscor:parameter>Заголовок</dcscor:parameter>
    <dcscor:value xsi:type="v8:LocalStringType">
      <v8:item>
        <v8:lang>ru</v8:lang>
        <v8:content>Текст заголовка</v8:content>
      </v8:item>
    </dcscor:value>
  </dcscor:item>
</dcsset:outputParameters>
```

Типичные параметры вывода:

| Параметр | Тип значения | Описание |
|---|---|---|
| `Заголовок` | `v8:LocalStringType` | Заголовок отчёта |
| `МакетОформления` | `xs:string` | Имя макета оформления: `ОформлениеОтчетовЧерноБелый`, `Зеленый` и др. |
| `РасположениеПолейГруппировки` | `dcsset:DataCompositionGroupFieldsPlacement` | `Together`, `Separately`, `SeparatelyAndInGroups` |
| `РасположениеРеквизитов` | `dcsset:DataCompositionAttributesPlacement` | `Together`, `Separately`, `SeparatelyAndInGroups` |
| `ГоризонтальноеРасположениеОбщихИтогов` | `dcscor:DataCompositionTotalPlacement` | `None`, `Begin`, `End`, `Auto` |
| `ВертикальноеРасположениеОбщихИтогов` | `dcscor:DataCompositionTotalPlacement` | `None`, `Begin`, `End`, `Auto` |
| `ВыводитьЗаголовок` | `dcsset:DataCompositionTextOutputType` | `Auto`, `DontOutput`, `Output` |
| `ВыводитьПараметрыДанных` | `dcsset:DataCompositionTextOutputType` | То же |
| `ВыводитьОтбор` | `dcsset:DataCompositionTextOutputType` | То же |

### 11.7. Параметры данных (dataParameters)

Значения параметров схемы в конкретном варианте:

```xml
<dcsset:dataParameters>
  <dcscor:item xsi:type="dcsset:SettingsParameterValue">
    <dcscor:use>false</dcscor:use>
    <dcscor:parameter>Период</dcscor:parameter>
    <dcscor:value xsi:type="v8:StandardPeriod">
      <v8:variant xsi:type="v8:StandardPeriodVariant">LastMonth</v8:variant>
    </dcscor:value>
    <dcsset:viewMode>Normal</dcsset:viewMode>
    <dcsset:userSettingID>GUID</dcsset:userSettingID>
  </dcscor:item>
</dcsset:dataParameters>
```

| Элемент | Описание |
|---|---|
| `dcscor:use` | `true`/`false` — использовать значение или нет |
| `dcscor:parameter` | Имя параметра из раздела 8 |
| `dcscor:value` | Значение параметра |
| `dcsset:viewMode` | Режим отображения: `Normal`, `QuickAccess`, `Inaccessible` |
| `dcsset:userSettingID` | GUID пользовательской настройки |

### 11.8. Элементы структуры (structure items)

Структура отчёта — иерархия группировок, таблиц, диаграмм.

#### StructureItemGroup — группировка

```xml
<dcsset:item xsi:type="dcsset:StructureItemGroup">
  <dcsset:name>Группировка</dcsset:name>
  <dcsset:groupItems>
    <dcsset:item xsi:type="dcsset:GroupItemField">
      <dcsset:field>Организация</dcsset:field>
      <dcsset:groupType>Items</dcsset:groupType>
      <dcsset:periodAdditionType>None</dcsset:periodAdditionType>
      <dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>
      <dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>
    </dcsset:item>
  </dcsset:groupItems>
  <dcsset:order>
    <dcsset:item xsi:type="dcsset:OrderItemAuto"/>
  </dcsset:order>
  <dcsset:selection>
    <dcsset:item xsi:type="dcsset:SelectedItemAuto"/>
  </dcsset:selection>
  <dcsset:outputParameters>...</dcsset:outputParameters>
  <dcsset:item xsi:type="dcsset:StructureItemGroup">  <!-- вложенная группировка -->
    ...
  </dcsset:item>
</dcsset:item>
```

Типы группировки (`groupType`): `Items`, `Hierarchy`, `HierarchyOnly`.

Типы дополнения периодом (`periodAdditionType`): `None`, `Year`, `HalfYear`, `Quarter`, `Month`, `TenDays`, `Week`, `Day`.

Пустая группировка (без `groupItems`) = детальные записи.

#### StructureItemTable — таблица (кросс-таблица)

```xml
<dcsset:item xsi:type="dcsset:StructureItemTable">
  <dcsset:name>Таблица</dcsset:name>
  <dcsset:column>                               <!-- группировки колонок -->
    <dcsset:groupItems>...</dcsset:groupItems>
    <dcsset:order>...</dcsset:order>
    <dcsset:selection>...</dcsset:selection>
  </dcsset:column>
  <dcsset:row>                                  <!-- группировки строк -->
    <dcsset:name>Группировка</dcsset:name>
    <dcsset:groupItems>...</dcsset:groupItems>
    <dcsset:order>...</dcsset:order>
    <dcsset:selection>...</dcsset:selection>
  </dcsset:row>
</dcsset:item>
```

#### StructureItemChart — диаграмма

```xml
<dcsset:item xsi:type="dcsset:StructureItemChart">
  <dcsset:point>                                <!-- точки (ось X) -->
    <dcsset:groupItems>...</dcsset:groupItems>
    <dcsset:order>...</dcsset:order>
    <dcsset:selection>...</dcsset:selection>
  </dcsset:point>
  <dcsset:series>                               <!-- серии (необязательно) -->
    <dcsset:groupItems>...</dcsset:groupItems>
    ...
  </dcsset:series>
  <dcsset:selection>                            <!-- значения для отображения -->
    <dcsset:item xsi:type="dcsset:SelectedItemField">
      <dcsset:field>РазмерДанных</dcsset:field>
    </dcsset:item>
  </dcsset:selection>
  <dcsset:outputParameters>...</dcsset:outputParameters>
</dcsset:item>
```

---

## 12. Типы данных — сводка

### v8:LocalStringType — локализованная строка

```xml
<title xsi:type="v8:LocalStringType">
  <v8:item>
    <v8:lang>ru</v8:lang>
    <v8:content>Текст на русском</v8:content>
  </v8:item>
</title>
```

Также можно задать как простую строку: `xsi:type="xs:string"`.

### dcscor:SettingsParameterValue — параметр настройки

```xml
<dcscor:item xsi:type="dcsset:SettingsParameterValue">
  <dcscor:use>true</dcscor:use>                <!-- опционально -->
  <dcscor:parameter>ИмяПараметра</dcscor:parameter>
  <dcscor:value xsi:type="ТипЗначения">Значение</dcscor:value>
</dcscor:item>
```

### dcscor:Field — ссылка на поле

```xml
<dcsset:left xsi:type="dcscor:Field">ИмяПоля</dcsset:left>
```

---

## 13. Полный минимальный пример

Простая DCS: один запрос, два поля, один итог, один вариант:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<DataCompositionSchema xmlns="http://v8.1c.ru/8.1/data-composition-system/schema"
    xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common"
    xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core"
    xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings"
    xmlns:v8="http://v8.1c.ru/8.1/data/core"
    xmlns:v8ui="http://v8.1c.ru/8.1/data/ui"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dataSource>
    <name>ИсточникДанных1</name>
    <dataSourceType>Local</dataSourceType>
  </dataSource>
  <dataSet xsi:type="DataSetQuery">
    <name>НаборДанных1</name>
    <field xsi:type="DataSetFieldField">
      <dataPath>Наименование</dataPath>
      <field>Наименование</field>
      <title xsi:type="v8:LocalStringType">
        <v8:item>
          <v8:lang>ru</v8:lang>
          <v8:content>Наименование</v8:content>
        </v8:item>
      </title>
    </field>
    <field xsi:type="DataSetFieldField">
      <dataPath>Количество</dataPath>
      <field>Количество</field>
    </field>
    <dataSource>ИсточникДанных1</dataSource>
    <query>ВЫБРАТЬ
	Номенклатура.Наименование КАК Наименование,
	КОЛИЧЕСТВО(1) КАК Количество
ИЗ
	Справочник.Номенклатура КАК Номенклатура
СГРУППИРОВАТЬ ПО
	Номенклатура.Наименование</query>
  </dataSet>
  <totalField>
    <dataPath>Количество</dataPath>
    <expression>Сумма(Количество)</expression>
  </totalField>
  <settingsVariant>
    <dcsset:name>Основной</dcsset:name>
    <dcsset:presentation xsi:type="v8:LocalStringType">
      <v8:item>
        <v8:lang>ru</v8:lang>
        <v8:content>Основной</v8:content>
      </v8:item>
    </dcsset:presentation>
    <dcsset:settings xmlns:style="http://v8.1c.ru/8.1/data/ui/style"
        xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system"
        xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web"
        xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows">
      <dcsset:selection>
        <dcsset:item xsi:type="dcsset:SelectedItemField">
          <dcsset:field>Наименование</dcsset:field>
        </dcsset:item>
        <dcsset:item xsi:type="dcsset:SelectedItemField">
          <dcsset:field>Количество</dcsset:field>
        </dcsset:item>
      </dcsset:selection>
      <dcsset:item xsi:type="dcsset:StructureItemGroup">
        <dcsset:order>
          <dcsset:item xsi:type="dcsset:OrderItemAuto"/>
        </dcsset:order>
        <dcsset:selection>
          <dcsset:item xsi:type="dcsset:SelectedItemAuto"/>
        </dcsset:selection>
      </dcsset:item>
    </dcsset:settings>
  </settingsVariant>
</DataCompositionSchema>
```
