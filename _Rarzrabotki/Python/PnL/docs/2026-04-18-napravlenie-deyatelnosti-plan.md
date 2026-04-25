# НаправлениеДеятельности в PnL-документах — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Добавить `НаправлениеДеятельности` как ось сверки Excel↔ERP в документы `А_РасшифровкаЛистов`, `А_ОтчетPL`, в отчёт `А_ОтчетPL` (новый СКД-вариант) и в Python-пайплайн PnL, чтобы сводные листы Excel (PL_Логистика, PL_ЦО_Строительство и т.п.) сопоставлялись с агрегатом по направлению, а не с одним подразделением.

**Architecture:** Два реквизита (`НаправлениеДеятельности`, `ВключатьДочерние`) добавляются в ТЧ `Расшифровка` документа `А_РасшифровкаЛистов` и в шапку `А_ОтчетPL`. Новый СКД-вариант `А_ОтчетPL_РасшифровкаПоНаправлению` группирует по `Подразделение` (в колонках) и фильтрует по `Подразделение.А_НаправлениеДеятельности`. Форма документа выбирает вариант отчёта по флагу `ВключатьДочерние`. Python-пайплайн (06→07→12/13, новый 14) обогащает JSON-mapping и заполняет/мигрирует данные через COM.

**Tech Stack:**
- 1C:Enterprise 8.3.20+, BAS ERP 2.5 (конфигурация INDUSTRIALBUD)
- Python 3 + win32com (V83.COMConnector), openpyxl, rapidfuzz
- MCP server `1c-workerp` для верификации
- Base: `BaseERP` (SQLSERVER)

**Source spec:** `_Rarzrabotki/Python/PnL/docs/2026-04-18-napravlenie-deyatelnosti-design.md` — прочитан целиком, §1–11 покрыты ниже.

**Workdir:** `C:\Configuration_downloads\BASERP25` (main). Python venv: `..\venv\Scripts\python.exe` (относительно `_Rarzrabotki\Python\PnL\`). Все пути — абсолютные.

---

## Справочные данные (из спека §2.2 — UUID направлений)

| Имя направления | UUID | Используется как свод для листов |
|---|---|---|
| Девелопмент | `9d2c84a1-1ae7-11f0-80dc-00155d235309` | — |
| Закрытые обьекты | `9d021b71-1ae7-11f0-80dc-00155d235309` | — |
| И.П.С. | `9d021b70-1ae7-11f0-80dc-00155d235309` | — |
| **Логистика** | `9d2c848d-1ae7-11f0-80dc-00155d235309` | Техника, ЦО_логистика, (PL_Логистика если появится) |
| **Производство** | `9d2c848e-1ae7-11f0-80dc-00155d235309` | PL_Производство_СВОД, PL_ЦО_Производство |
| Спецтехника | `9d021b6e-1ae7-11f0-80dc-00155d235309` | — |
| **Строительство** | `9d021b84-1ae7-11f0-80dc-00155d235309` | PL Строительство Свод, PL_ЦО_Строительство |
| **ЦО** | `9d021b6f-1ae7-11f0-80dc-00155d235309` | PL_ЦО |

Справочник.НаправленияДеятельности проверен через MCP 2026-04-18 (8 элементов, все без пометки удаления).

---

## File structure

### Конфигурация 1С (изменяет пользователь через Designer/EDT)
- `Documents/А_РасшифровкаЛистов.xml` (метаданные объекта)
- `Documents/А_ОтчетPL.xml` (метаданные объекта)
- `Documents/А_ОтчетPL/Ext/ObjectModule.bsl` (валидация в ПередЗаписью)
- `Documents/А_ОтчетPL/Forms/ФормаДокумента/Ext/Form/Module.bsl` (ПриИзмененииПодразделения + переписать ОткрытьСверкуСЕРП)
- `Documents/А_ОтчетPL/Forms/ФормаДокумента/Ext/Form.xml` (разместить новые реквизиты на форме)
- `Reports/А_ОтчетPL/Templates/ОсновнаяСхемаКомпоновкиДанных/Ext/Template.xml` (поле НаправлениеДеятельности + новый settingsVariant)
- (опционально) `Documents/А_РасшифровкаЛистов/Forms/ФормаДокумента/...` — добавить колонки ТЧ в представление

### Python (изменяет агент)
- `_Rarzrabotki/Python/PnL/config.py` (MANUAL_SHEET_TO_STRUCT_OVERRIDES +direction_uuid/name/include_children)
- `_Rarzrabotki/Python/PnL/scripts/06_extract_struct_from_erp.py` (читать А_НаправлениеДеятельности + А_ЭтоПодразделениеНаправление)
- `_Rarzrabotki/Python/PnL/scripts/07_match_sheets_to_struct.py` (записывать direction_* + include_children)
- `_Rarzrabotki/Python/PnL/scripts/08_prepare_documents.py` (пробросить direction_uuid/include_children в 08_documents_to_import.json)
- `_Rarzrabotki/Python/PnL/scripts/12_upload_documents.py` (заполнять шапку НаправлениеДеятельности/ВключатьДочерние)
- `_Rarzrabotki/Python/PnL/scripts/13_fill_rasshifrovka_listov.py` (заполнять ТЧ Расшифровка новыми полями)
- `_Rarzrabotki/Python/PnL/scripts/14_migrate_otchetpl_napravlenie.py` (**НОВЫЙ** — разовая миграция 37 существующих документов)

### Документация
- `_Rarzrabotki/Python/PnL/docs/prompt_fill_rasshifrovka.md` (обновить: новые поля ТЧ, новые acceptance criteria)

---

## Phase overview + checkpoints

Работа разбита на 6 фаз. **После каждой фазы — chek-point:** агент показывает результат, пользователь даёт "далее" или корректировки. Переход к следующей фазе только после подтверждения.

| Phase | Что делаем | Кто | Обратимость |
|---|---|---|---|
| 0 | Baseline: MCP snapshot, CF backup | Агент | read-only |
| 1 | Метаданные 1С + Форма + СКД + ObjectModule | **Пользователь (Designer/EDT)** — агент даёт инструкцию | ручной откат из CF-бэкапа |
| 2 | Python extract/match — 06/07 + config.py | Агент (dry-verify после) | перезапуск 06/07 |
| 3 | Python upload — 08/12/13 dry+live | Агент (dry-run ОБЯЗАТЕЛЕН) | backup JSON в data/json/ |
| 4 | Миграция существующих А_ОтчетPL (скрипт 14) | Агент (dry-run ОБЯЗАТЕЛЕН) | backup JSON |
| 5 | Функциональные тесты + обновление доки | Агент + пользователь | — |

---

# Phase 0 — Baseline snapshot (read-only)

**Цель:** зафиксировать текущее состояние боевой базы и конфигурации, чтобы была возможность сравнить "до/после" и откатиться в крайнем случае.

**Files:**
- Create: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\data\json\00_baseline_pre_direction_migration.json`

### Task 0.1: MCP snapshot текущих документов

- [ ] **Step 1: Зафиксировать количество и список документов А_РасшифровкаЛистов**

Через MCP `mcp__1c-workerp__execute_query`:

```sql
ВЫБРАТЬ Д.Номер, Д.Дата, Д.ИмяФайла,
    КОЛИЧЕСТВО(ТЧ.НомерСтроки) КАК Строк,
    СУММА(ВЫБОР КОГДА ТЧ.Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд
ИЗ Документ.А_РасшифровкаЛистов КАК Д
ЛЕВОЕ СОЕДИНЕНИЕ Документ.А_РасшифровкаЛистов.Расшифровка КАК ТЧ
    ПО ТЧ.Ссылка = Д.Ссылка
ГДЕ НЕ Д.ПометкаУдаления
СГРУППИРОВАТЬ ПО Д.Номер, Д.Дата, Д.ИмяФайла
УПОРЯДОЧИТЬ ПО Д.Дата
```

Ожидается (подтверждено 2026-04-18): 3 строки с `Строк = 31/29/31`, `СПодразд = 25/25/26`, `ИмяФайла` = полный путь.

- [ ] **Step 2: Зафиксировать количество А_ОтчетPL**

```sql
ВЫБРАТЬ
    КОЛИЧЕСТВО(Д.Ссылка) КАК Всего,
    СУММА(ВЫБОР КОГДА Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК ПометкаУд,
    СУММА(ВЫБОР КОГДА Д.Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд
ИЗ Документ.А_ОтчетPL КАК Д
```

Ожидается: `Всего = 37`, `ПометкаУд = 0` (по состоянию 2026-04-18).

- [ ] **Step 3: Записать baseline в JSON**

Агент складывает результаты в `data/json/00_baseline_pre_direction_migration.json` с полями: `timestamp`, `rasshifrovka_listov` (3 doc summaries), `otchet_pl_count`, `napravleniya` (8 UUID).

### Task 0.2: Pre-flight — покрытие направления по подразделениям из 07_mapping

**Критично для корректности ожиданий Phase 2/3.** Проверяем, сколько подразделений из текущего 07_mapping имеют заполненное `А_НаправлениеДеятельности`.

- [ ] **Step 1: Запрос**

```sql
ВЫБРАТЬ
    СП.Ссылка,
    СП.Наименование,
    ВЫБОР КОГДА СП.А_НаправлениеДеятельности = ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка)
        ТОГДА "—" ИНАЧЕ СП.А_НаправлениеДеятельности.Наименование КОНЕЦ КАК Напр
ИЗ Справочник.СтруктураПредприятия КАК СП
ГДЕ СП.Ссылка В (&СписокСсылок)
УПОРЯДОЧИТЬ ПО СП.Наименование
```

Список ссылок — все 29 `struct_uuid` из `07_mapping_sheet_to_struct.json`.

- [ ] **Step 2: Фиксация ожидания**

Если N из 29 подразделений имеют `А_НаправлениеДеятельности` — это будет верхняя граница `with_direction` после Phase 3. Записать N в baseline JSON (`expected_with_direction_per_file`).

Если есть подразделения без направления — агент сообщает пользователю список и ждёт решения: либо (a) финансист заполнит направление в клиенте, либо (b) для этих листов будет пустое направление (документ тогда игнорирует новую ось).

### Task 0.3: CF backup конфигурации (ВЫПОЛНЯЕТ ПОЛЬЗОВАТЕЛЬ)

- [ ] **Step 1: Выгрузить текущий `.cf`-файл конфигурации**

**Агент сообщает пользователю:** «Перед изменением метаданных выгрузи `.cf` текущей конфигурации через Designer:
```
Configurator → Конфигурация → Сохранить конфигурацию в файл → ...
путь: C:\Configuration_downloads\BASERP25\_backup\BaseERP_pre_direction_2026-04-18.cf
```
Или через пакетный режим (skill `db-dump-cf`): `1cv8 DESIGNER /DumpCfg <путь>`. После сохранения дай команду 'baseline готов' — продолжим.»

**Checkpoint 0 — показать пользователю:**
- результаты baseline-снимка (2 query выше);
- факт создания `.cf` backup;
- подтверждение состава справочника НаправленияДеятельности (8 элементов + UUID в таблице выше);
- путь к созданному `data/json/00_baseline_pre_direction_migration.json`.

**Acceptance criteria Phase 0:**
- [ ] MCP-queries вернули ожидаемые значения (3/37/8).
- [ ] Task 0.2 pre-flight — зафиксировано число подразделений из 07_mapping с заполненным `А_НаправлениеДеятельности` (N из 29).
- [ ] `.cf` backup сохранён и его путь подтверждён пользователем.
- [ ] `00_baseline_pre_direction_migration.json` создан.

---

# Phase 1 — Метаданные 1С (ВЫПОЛНЯЕТ ПОЛЬЗОВАТЕЛЬ)

**Цель:** добавить два реквизита в `А_РасшифровкаЛистов.ТЧ.Расшифровка`, два реквизита в шапку `А_ОтчетPL`, новое поле `НаправлениеДеятельности` в СКД отчёта `А_ОтчетPL`, новый settingsVariant `А_ОтчетPL_РасшифровкаПоНаправлению`, обновить ObjectModule + Form.Module документа.

**Агент НЕ правит XML/CFE/BSL автоматически.** Агент готовит точную инструкцию и snippet'ы кода. Пользователь правит через Designer/EDT, обновляет БД, сообщает "метаданные обновлены".

**Files (пользователь правит):**
- `C:\Configuration_downloads\BASERP25\Documents\А_РасшифровкаЛистов.xml`
- `C:\Configuration_downloads\BASERP25\Documents\А_ОтчетPL.xml`
- `C:\Configuration_downloads\BASERP25\Documents\А_ОтчетPL\Ext\ObjectModule.bsl`
- `C:\Configuration_downloads\BASERP25\Documents\А_ОтчетPL\Forms\ФормаДокумента\Ext\Form\Module.bsl`
- `C:\Configuration_downloads\BASERP25\Documents\А_ОтчетPL\Forms\ФормаДокумента\Ext\Form.xml`
- `C:\Configuration_downloads\BASERP25\Reports\А_ОтчетPL\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml`

### Task 1.1: Расширить Документ.А_РасшифровкаЛистов — ТЧ Расшифровка

- [ ] **Step 1: Агент выдаёт пользователю инструкцию**

**Пользовательская инструкция:**

> Открыть Конфигуратор → Документ `А_РасшифровкаЛистов` → Табличная часть `Расшифровка` → Добавить два реквизита:
>
> | Реквизит | Тип | Синоним | Индексирование |
> |---|---|---|---|
> | `НаправлениеДеятельности` | `СправочникСсылка.НаправленияДеятельности` | Направление деятельности | Нет |
> | `ВключатьДочерние` | `Булево` | Включать дочерние | Нет |
>
> Обе колонки — без обязательности (заполнятся скриптом и вручную). После добавления сохранить конфигурацию.

### Task 1.2: Расширить Документ.А_ОтчетPL — шапка

- [ ] **Step 1: Агент выдаёт пользователю инструкцию**

**Пользовательская инструкция:**

> Открыть Конфигуратор → Документ `А_ОтчетPL` → Реквизиты (шапка) → Добавить:
>
> | Реквизит | Тип | Синоним |
> |---|---|---|
> | `НаправлениеДеятельности` | `СправочникСсылка.НаправленияДеятельности` | Направление деятельности |
> | `ВключатьДочерние` | `Булево` | Включать дочерние |

### Task 1.3: Обновить ObjectModule Документа.А_ОтчетPL — валидация

- [ ] **Step 1: Агент выдаёт snippet для вставки**

**Пользовательская инструкция:** добавить в начало процедуры `ПередЗаписью` (до `Если ОбменДанными.Загрузка`), в `C:\Configuration_downloads\BASERP25\Documents\А_ОтчетPL\Ext\ObjectModule.bsl`:

```bsl
// Валидация: свод по направлению требует заполненного НаправлениеДеятельности.
Если ВключатьДочерние И НЕ ЗначениеЗаполнено(НаправлениеДеятельности) Тогда
    Отказ = Истина;
    Сообщение = Новый СообщениеПользователю;
    Сообщение.Текст = "Установлен флаг «Включать дочерние», но «НаправлениеДеятельности» не заполнено. "
        + "Либо снимите флаг, либо укажите направление.";
    Сообщение.Поле = "НаправлениеДеятельности";
    Сообщение.УстановитьДанные(ЭтотОбъект);
    Сообщение.Сообщить();
    Возврат;
КонецЕсли;
```

*ВАЖНО:* валидация должна сработать даже при обычной записи (не только при Проведении), поэтому ставится выше `Если ОбменДанными.Загрузка`. При импорте через Python (`ОбменДанными.Загрузка = Истина`) валидация пропускается — это штатно.

### Task 1.4: Обновить Модуль формы ФормаДокумента А_ОтчетPL

- [ ] **Step 1: Добавить обработчик ПриИзмененииПодразделения**

Вставить в `C:\Configuration_downloads\BASERP25\Documents\А_ОтчетPL\Forms\ФормаДокумента\Ext\Form\Module.bsl` (после `ПослеЗаписиНаСервере`, до `ОткрытьСверкуСЕРП`):

```bsl
&НаКлиенте
Процедура ПодразделениеПриИзменении(Элемент)
    ПодтянутьНаправлениеПоПодразделению();
КонецПроцедуры

&НаСервере
Процедура ПодтянутьНаправлениеПоПодразделению()
    Если ЗначениеЗаполнено(Объект.Подразделение)
        И НЕ ЗначениеЗаполнено(Объект.НаправлениеДеятельности) Тогда
        Объект.НаправлениеДеятельности = Объект.Подразделение.А_НаправлениеДеятельности;
    КонецЕсли;
КонецПроцедуры
```

Автоподстановка — **только если направление ещё пустое**. Финансист сможет перекрыть вручную.

- [ ] **Step 2: Переписать процедуру ОткрытьСверкуСЕРП**

Заменить `ОткрытьСверкуСЕРП` и `СформироватьПользовательскиеНастройкиОтчета` в `Form/Module.bsl` на версию, поддерживающую два варианта:

```bsl
&НаКлиенте
Процедура ОткрытьСверкуСЕРП(Команда)

    ИспользуемПоНаправлению = Объект.ВключатьДочерние
        И ЗначениеЗаполнено(Объект.НаправлениеДеятельности);

    Если ИспользуемПоНаправлению Тогда
        ВариантКлюч = "А_ОтчетPL_РасшифровкаПоНаправлению";
    Иначе
        ВариантКлюч = "А_ОтчетPL_РасшифровкаДокументаОтчетPL";
    КонецЕсли;

    НастройкиСПериодом = СформироватьПользовательскиеНастройкиОтчета(Объект.Дата, ВариантКлюч);

    Отбор = Новый Структура;
    Если ИспользуемПоНаправлению Тогда
        Отбор.Вставить("НаправлениеДеятельности", Объект.НаправлениеДеятельности);
    ИначеЕсли ЗначениеЗаполнено(Объект.Подразделение) Тогда
        Отбор.Вставить("Подразделение", Объект.Подразделение);
    КонецЕсли;

    ПараметрыФормы = Новый Структура;
    ПараметрыФормы.Вставить("ВариантКлюч", ВариантКлюч);
    Если НастройкиСПериодом <> Неопределено Тогда
        ПараметрыФормы.Вставить("ПользовательскиеНастройки", НастройкиСПериодом);
    КонецЕсли;
    Если Отбор.Количество() > 0 Тогда
        ПараметрыФормы.Вставить("Отбор", Отбор);
    КонецЕсли;
    ПараметрыФормы.Вставить("СформироватьПриОткрытии", Истина);

    ОткрытьФорму("Отчет.А_ОтчетPL.Форма", ПараметрыФормы, ЭтаФорма);

КонецПроцедуры

&НаСервереБезКонтекста
Функция СформироватьПользовательскиеНастройкиОтчета(Дата, ВариантКлюч)

    Отчет = Отчеты.А_ОтчетPL.Создать();
    Вариант = Отчет.СхемаКомпоновкиДанных.ВариантыНастроек.Найти(ВариантКлюч);
    Если Вариант = Неопределено Тогда
        Возврат Неопределено;
    КонецЕсли;

    Компоновщик = Новый КомпоновщикНастроекКомпоновкиДанных;
    Источник = Новый ИсточникДоступныхНастроекКомпоновкиДанных(Отчет.СхемаКомпоновкиДанных);
    Компоновщик.Инициализировать(Источник);
    Компоновщик.ЗагрузитьНастройки(Вариант.Настройки);

    Период = Новый СтандартныйПериод;
    Период.ДатаНачала = НачалоМесяца(Дата);
    Период.ДатаОкончания = КонецМесяца(Дата);

    Для Каждого Элемент Из Компоновщик.ПользовательскиеНастройки.Элементы Цикл
        Если ТипЗнч(Элемент) = Тип("ЗначениеПараметраНастроекКомпоновкиДанных")
            И Элемент.ИдентификаторПользовательскойНастройки = "Период" Тогда
            Элемент.Значение = Период;
            Элемент.Использование = Истина;
            Прервать;
        КонецЕсли;
    КонецЦикла;

    Возврат Компоновщик.ПользовательскиеНастройки;

КонецФункции
```

Ключевое отличие от текущей версии:
1. Вариант выбирается по флагу `ВключатьДочерние`.
2. Отбор — по `НаправлениеДеятельности` или по `Подразделение`, в зависимости от варианта.
3. Функция `СформироватьПользовательскиеНастройкиОтчета` принимает `ВариантКлюч` параметром.

### Task 1.5: Разместить новые реквизиты на форме А_ОтчетPL

- [ ] **Step 1: Инструкция пользователю**

> Открыть ФормаДокумента Документа А_ОтчетPL в Конфигураторе. В диалоге «Реквизиты формы» нажать «Добавить поля основного реквизита» — Конфигуратор предложит добавить новые реквизиты `НаправлениеДеятельности`, `ВключатьДочерние`.
>
> Разместить на форме (группа «Шапка» рядом с `Подразделение`): `НаправлениеДеятельности` (ввод по строке из справочника), `ВключатьДочерние` (флажок).
>
> Для поля `Подразделение` в «Обработчики событий» указать уже созданный в модуле `ПодразделениеПриИзменении`.

### Task 1.6: Расширить СКД отчёта А_ОтчетPL

Файл: `C:\Configuration_downloads\BASERP25\Reports\А_ОтчетPL\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml` (текущая длина 437 строк).

- [ ] **Step 1: Добавить в `dataSet` поле `НаправлениеДеятельности`**

**Инструкция пользователю:**
> Открыть Конфигуратор → Отчет `А_ОтчетPL` → Основная схема компоновки данных → вкладка «Наборы данных». В запросе набора `Данные`:
>
> 1. Добавить в секцию ВЫБРАТЬ поле `Данные.Подразделение.А_НаправлениеДеятельности КАК НаправлениеДеятельности`.
> 2. На вкладке «Поля» появится автоматически поле `НаправлениеДеятельности` (тип `СправочникСсылка.НаправленияДеятельности`). Проверить синоним «Направление деятельности».

- [ ] **Step 2: Создать новый settingsVariant «А_ОтчетPL_РасшифровкаПоНаправлению»**

**Инструкция пользователю:**
> В конструкторе СКД → вкладка «Настройки» → Доступные варианты → Добавить. Параметры варианта:
>
> - **Имя:** `А_ОтчетPL_РасшифровкаПоНаправлению`
> - **Представление:** `Расшифровка по направлению`
>
> В варианте:
> - **Отбор:** `НаправлениеДеятельности = &НаправлениеДеятельности` (с Пользовательской настройкой "НаправлениеДеятельности", быстрый доступ).
> - **Параметры:** `Период` (стандартный, как в существующем варианте).
> - **Выбранные поля:** `ДДС, СуммаФ1, СуммаФ2, СуммаPL, СуммаЕРП, Разница` (как в варианте «Расшифровка документа PL»).
> - **Структура:**
>   - Группировка **в строках**: `Группа → СтатьяPL`
>   - Группировка **в колонках**: `Подразделение` (это ключевое отличие — разворачиваем агрегат по дочерним).
>
> После сохранения — проверить что в XML появился `<settingsVariant><dcsset:name>А_ОтчетPL_РасшифровкаПоНаправлению</dcsset:name>...` (в том же файле Template.xml).

### Task 1.7: Обновить БД конфигурации

- [ ] **Step 1: Инструкция пользователю**

> В Конфигураторе выполнить «Обновить конфигурацию базы данных» (F7 → Применить изменения). Если запросит реструктуризацию — подтвердить (добавление реквизитов к документам — онлайн, без блокировки сеансов). Если есть активные сеансы — завершить их в Администрирование → Активные пользователи.
>
> После обновления БД — перезапустить сеансы клиентов, чтобы они увидели новые реквизиты.

### Task 1.8: Верификация метаданных через MCP

- [ ] **Step 1: Проверить реквизиты документов**

Агент выполняет (после команды пользователя «метаданные обновлены»):

```
mcp__1c-workerp__get_metadata_structure(metaType="Documents", name="А_РасшифровкаЛистов")
mcp__1c-workerp__get_metadata_structure(metaType="Documents", name="А_ОтчетPL")
```

Ожидается:
- В ТЧ `А_РасшифровкаЛистов.Расшифровка` появились `НаправлениеДеятельности` (Справочник.НаправленияДеятельности) и `ВключатьДочерние` (Булево).
- В шапке `А_ОтчетPL` появились такие же два реквизита.

- [ ] **Step 2: Проверить новый вариант СКД через тест-запрос**

Попытаться построить компоновку через `execute_query` не получится, но можно проверить через прямое чтение XML:

```
Read C:\Configuration_downloads\BASERP25\Reports\А_ОтчетPL\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml
Grep "А_ОтчетPL_РасшифровкаПоНаправлению" Template.xml
Grep "НаправлениеДеятельности" Template.xml
```

Ожидается: оба паттерна найдены.

- [ ] **Step 3: Проверить валидацию в ПередЗаписью (smoke-test через MCP)**

Через `mcp__1c-workerp__execute_query` прочитать один существующий документ А_ОтчетPL, затем через MCP-find взять его Ссылку — записать (`update_document` без изменений) и проверить что он пишется без ошибки. Если выставить `ВключатьДочерние=True` без направления — ожидать `Отказ = Истина` и сообщение.

*Примечание:* MCP-tool `update_document` не всегда может трогать шапку конкретного документа с кастомным ПередЗаписью. Альтернатива — пользователь сам в клиенте 1С ставит флаг на одном документе, убеждается что сохранение отказало с понятным сообщением, снимает флаг.

**Checkpoint 1 — показать пользователю:**
- MCP подтверждение реквизитов (скриншот вывода `get_metadata_structure`);
- grep-результаты по Template.xml (новый variant + поле);
- smoke-test валидации.

**Acceptance criteria Phase 1:**
- [ ] `get_metadata_structure("Documents","А_РасшифровкаЛистов")` показывает в ТЧ Расшифровка оба новых реквизита.
- [ ] `get_metadata_structure("Documents","А_ОтчетPL")` показывает в шапке оба новых реквизита.
- [ ] В Template.xml СКД есть строка `<dcsset:name>А_ОтчетPL_РасшифровкаПоНаправлению</dcsset:name>`.
- [ ] В Template.xml есть `<dataPath>НаправлениеДеятельности</dataPath>` (или `field>НаправлениеДеятельности<`).
- [ ] Валидация `ВключатьДочерние без направления` сработала (ручной smoke-test пользователем).

---

# Phase 2 — Python extract/match (06 + config + 07)

**Цель:** обогатить JSON структуры подразделений направлениями, расширить overrides, перегенерировать 07_mapping_sheet_to_struct.json с полями `direction_uuid`, `direction_name`, `include_children`.

**Files:**
- Modify: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\06_extract_struct_from_erp.py`
- Modify: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\config.py`
- Modify: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\07_match_sheets_to_struct.py`

### Task 2.1: Обновить 06_extract_struct_from_erp.py — читать направление

- [ ] **Step 1: Заменить QUERY на обогащённый вариант**

Новое содержимое константы `QUERY` в `06_extract_struct_from_erp.py`:

```python
QUERY = """
ВЫБРАТЬ
    СП.Ссылка КАК Ссылка,
    СП.Наименование КАК Наименование,
    СП.Код КАК Код,
    СП.Родитель КАК Родитель,
    СП.Родитель.Наименование КАК РодительНаименование,
    СП.А_НаправлениеДеятельности КАК НаправлениеДеятельности,
    СП.А_НаправлениеДеятельности.Наименование КАК НаправлениеИмя,
    СП.А_ЭтоПодразделениеНаправление КАК ЭтоНаправление,
    СП.ПометкаУдаления КАК ПометкаУдаления
ИЗ
    Справочник.СтруктураПредприятия КАК СП
ГДЕ
    НЕ СП.ПометкаУдаления
УПОРЯДОЧИТЬ ПО
    СП.Наименование
"""
```

- [ ] **Step 2: Заменить сборку `out` на обогащённую**

В функции `main()` заменить цикл наполнения `out`:

```python
    for i in range(n):
        r = tz.Получить(i)
        napr_ref = r.НаправлениеДеятельности
        napr_empty = napr_ref.Пустая() if napr_ref else True
        parent_ref = r.Родитель
        parent_empty = parent_ref.Пустая() if parent_ref else True
        out.append({
            "uuid": uuid_str(conn, r.Ссылка),
            "name": str(r.Наименование),
            "code": str(r.Код),
            "parent_uuid": None if parent_empty else uuid_str(conn, parent_ref),
            "parent_name": str(r.РодительНаименование) if r.РодительНаименование else "",
            "direction_uuid": None if napr_empty else uuid_str(conn, napr_ref),
            "direction_name": str(r.НаправлениеИмя) if r.НаправлениеИмя else "",
            "is_direction": bool(r.ЭтоНаправление),
        })
```

Поля `parent` (было просто имя родителя — строка) заменены на `parent_uuid` + `parent_name` — это нужно для построения иерархии в варианте "по направлению" (для Phase 5 ручной сверки). Поле `parent` оставлять **не нужно** — потребители из старого кода (08_prepare_documents, 07_match) используют только `uuid` и `name`.

- [ ] **Step 3: Проверить обратную совместимость**

```
Grep '06_struct_predpr' C:/Configuration_downloads/BASERP25/_Rarzrabotki/Python/PnL/scripts/
```

Ожидаемые потребители: 07_match_sheets_to_struct.py, возможно 08/диагностические. Убедиться что никто не читает поле `parent` (старое имя) — если читает, заменить на `parent_name`.

- [ ] **Step 4: Запустить 06 и проверить JSON**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
..\venv\Scripts\python.exe scripts\06_extract_struct_from_erp.py
```

Ожидание: `Wrote .../06_struct_predpr.json  (items: N)` где N ≥ 200.

Проверить вручную:
```
Read C:/Configuration_downloads/BASERP25/_Rarzrabotki/Python/PnL/data/json/06_struct_predpr.json (limit 30)
```
Должна быть структура `{total, items: [{uuid, name, code, parent_uuid, parent_name, direction_uuid, direction_name, is_direction}]}`.

У подразделений-направлений `is_direction = true` и `direction_uuid = uuid` (само себя). У рядовых подразделений (Крушинка, MAN №3) — `is_direction = false` и `direction_uuid` указывает на UUID направления.

### Task 2.2: Расширить config.MANUAL_SHEET_TO_STRUCT_OVERRIDES

- [ ] **Step 1: Добавить поля `direction_uuid`, `direction_name`, `include_children` для сводных листов**

В `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\config.py` обновить следующие записи:

```python
    "Техника": {
        "struct_uuid": "950c32e4-ca27-11f0-a2e2-c54425f51b91",
        "struct_name": "Логистика",
        "direction_uuid": "9d2c848d-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Логистика",
        "include_children": True,
        "reason": "Решение финансиста: лист Техника сворачивает всю Логистику в один PnL-лист",
    },
    "PL_ЦО": {
        "struct_uuid": "19c7dbef-ca2d-11f0-a2e2-c54425f51b91",
        "struct_name": "ЦО",
        "direction_uuid": "9d021b6f-1ae7-11f0-80dc-00155d235309",
        "direction_name": "ЦО",
        "include_children": True,
        "reason": "Сводный лист ЦО = корневое подразделение ЦО",
    },
    "PL Строительство Свод": {
        "struct_uuid": "98e5f138-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Строительство",
        "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Строительство",
        "include_children": True,
        "reason": "Сводный лист Строительство = корневое подразделение Строительство",
    },
    "PL_ЦО_Строительство": {
        "struct_uuid": "98e5f138-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Строительство",
        "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Строительство",
        "include_children": True,
        "reason": "ЦО по Строительству = корневое подразделение Строительство",
    },
    "PL_Производство_СВОД": {
        "struct_uuid": "65c4e837-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Производство",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": True,
        "reason": "Сводный лист Производство = корневое подразделение Производство",
    },
    "PL_ЦО_Производство": {
        "struct_uuid": "65c4e837-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Производство",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": True,
        "reason": "ЦО по Производству = корневое подразделение Производство",
    },
    "ЦО_логистика": {
        "struct_uuid": "950c32e4-ca27-11f0-a2e2-c54425f51b91",
        "struct_name": "Логистика",
        "direction_uuid": "9d2c848d-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Логистика",
        "include_children": True,
        "reason": "ЦО по Логистике = корневое подразделение Логистика",
    },
```

Записи, которые **НЕ трогаем** (элементы, не своды) — `ПРОООН Черкаси ДСНС`, `МК Глобино`, `МД ООН 2025`, `МД ВООЗ 2025`, `Экспедирование`, `MAN №3/4/5`. Для них `direction_*` и `include_children` не задаются — скрипт 07 сам подтянет `direction_uuid` из 06_struct_predpr и выставит `include_children=False`.

- [ ] **Step 2: Smoke-test: файл остался валидным Python**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
..\venv\Scripts\python.exe -c "import config; print(len(config.MANUAL_SHEET_TO_STRUCT_OVERRIDES))"
```

Ожидается: `15` (как и было). Все записи сохранились, 7 обогащены направлением.

### Task 2.3: Доработать 07_match_sheets_to_struct.py — запись direction_*

- [ ] **Step 1: Добавить helper для подстановки направления из 06**

В `07_match_sheets_to_struct.py`, сразу после загрузки `struct` в `main()`:

```python
    struct_by_uuid = {s["uuid"]: s for s in struct}
```

- [ ] **Step 2: В месте формирования `result` добавить direction_uuid/name/include_children**

Заменить блок формирования записи (строки ~73-80):

```python
        # Подтянуть direction_uuid из 06_struct_predpr по struct_uuid (если найден)
        direction_uuid = ""
        direction_name = ""
        if struct_uuid:
            s = struct_by_uuid.get(struct_uuid)
            if s:
                direction_uuid = s.get("direction_uuid") or ""
                direction_name = s.get("direction_name") or ""

        result.append({
            "sheet_name": sn,
            "struct_uuid": struct_uuid,
            "struct_name": struct_name,
            "match_type": match_type,
            "confidence": confidence,
            "direction_uuid": direction_uuid,
            "direction_name": direction_name,
            "include_children": False,
            "candidates": candidates,
        })
```

- [ ] **Step 3: В секции применения overrides подхватить direction_* из override**

Заменить блок `for sheet_name, ov in overrides.items():` (строки ~87-110):

```python
    for sheet_name, ov in overrides.items():
        ov_dir_uuid = ov.get("direction_uuid") or ""
        ov_dir_name = ov.get("direction_name") or ""
        ov_include = bool(ov.get("include_children", False))

        # Если в override не указан direction_uuid — подтянуть из 06 по struct_uuid
        if not ov_dir_uuid:
            s = struct_by_uuid.get(ov["struct_uuid"])
            if s:
                ov_dir_uuid = s.get("direction_uuid") or ""
                ov_dir_name = s.get("direction_name") or ""

        if sheet_name in existing_names:
            r = existing_names[sheet_name]
            r["struct_uuid"] = ov["struct_uuid"]
            r["struct_name"] = ov["struct_name"]
            r["match_type"] = "manual_override"
            r["confidence"] = 100
            r["direction_uuid"] = ov_dir_uuid
            r["direction_name"] = ov_dir_name
            r["include_children"] = ov_include
            r["candidates"] = []
            r["notes"] = ov.get("reason", "")
            override_applied += 1
            if sheet_name in unmapped:
                unmapped.remove(sheet_name)
        else:
            result.append({
                "sheet_name": sheet_name,
                "struct_uuid": ov["struct_uuid"],
                "struct_name": ov["struct_name"],
                "match_type": "manual_override",
                "confidence": 100,
                "direction_uuid": ov_dir_uuid,
                "direction_name": ov_dir_name,
                "include_children": ov_include,
                "candidates": [],
                "notes": ov.get("reason", ""),
            })
            override_added += 1
```

- [ ] **Step 4: Запустить 07**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
..\venv\Scripts\python.exe scripts\07_match_sheets_to_struct.py
```

Ожидается счётчики в stdout, идентичные текущим (9 exact, 5 fuzzy, 15 manual_override — 29 mappings).

- [ ] **Step 5: Верификация результата**

```bash
..\venv\Scripts\python.exe -c "import json,pathlib; p=pathlib.Path('data/json/07_mapping_sheet_to_struct.json'); m=json.loads(p.read_text(encoding='utf-8'))['mappings']; withdir=sum(1 for x in m if x.get('direction_uuid')); incch=sum(1 for x in m if x.get('include_children')); print(f'total={len(m)} with_direction={withdir} include_children={incch}')"
```

Ожидается: `total=29 include_children=7` (7 сводных листов: Техника, PL_ЦО, PL Строительство Свод, PL_ЦО_Строительство, PL_Производство_СВОД, PL_ЦО_Производство, ЦО_логистика). `with_direction` ≈ N (см. Phase 0 Task 0.3 — это число подразделений из mapping, у которых заполнено `А_НаправлениеДеятельности`). В идеале равно 29; если меньше — это ожидаемый дефект данных (не бага скрипта).

Точечная проверка:
```bash
..\venv\Scripts\python.exe -c "import json; m=json.load(open('data/json/07_mapping_sheet_to_struct.json','r',encoding='utf-8'))['mappings']; d={x['sheet_name']:x for x in m}; print('PL_ЦО:', d['PL_ЦО']); print('Крушинка:', d.get('Крушинка'))"
```

Ожидается:
- `PL_ЦО`: `direction_uuid='9d021b6f-...'`, `direction_name='ЦО'`, `include_children=True`.
- `Крушинка`: `direction_uuid` указывает на одно из направлений (определяется реквизитом подразделения), `include_children=False`.

**Checkpoint 2 — показать пользователю:**
- stdout 06 (items count);
- stdout 07 (exact/fuzzy/manual);
- результаты двух smoke-queries по JSON (total mappings, with_direction, include_children — и pretty-print для PL_ЦО/Крушинки).

**Acceptance criteria Phase 2:**
- [ ] `06_struct_predpr.json` содержит `direction_uuid`/`direction_name`/`is_direction` у каждой записи.
- [ ] `07_mapping_sheet_to_struct.json` содержит 29 записей (как до правок), ≥20 с `direction_uuid`, 7 с `include_children=True`.
- [ ] Для сводных листов `include_children=True` И `direction_uuid` заполнен одновременно.
- [ ] Для всех элементных листов из 07_mapping (exact/fuzzy) `direction_uuid` не пустой (подтянут из 06).

---

# Phase 3 — Python upload (08 + 12 + 13)

**Цель:** заполнить 3 существующих документа А_РасшифровкаЛистов обогащёнными данными (direction + include_children) и обеспечить чтобы 12_upload для будущих импортов тоже заполнял шапку А_ОтчетPL направлением.

**Files:**
- Modify: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\08_prepare_documents.py`
- Modify: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\12_upload_documents.py`
- Modify: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\13_fill_rasshifrovka_listov.py`

### Task 3.1: Пробросить direction_uuid/include_children в 08_documents_to_import.json

- [ ] **Step 1: Обновить 08_prepare_documents.py**

В `main()` добавить сбор direction_* из 07_mapping (параллельно sheet_to_uuid):

```python
    sheet_to_uuid = {m["sheet_name"]: m["struct_uuid"] for m in struct_map}
    sheet_to_sname = {m["sheet_name"]: m["struct_name"] for m in struct_map}
    sheet_to_diruuid = {m["sheet_name"]: m.get("direction_uuid", "") for m in struct_map}
    sheet_to_dirname = {m["sheet_name"]: m.get("direction_name", "") for m in struct_map}
    sheet_to_incch = {m["sheet_name"]: bool(m.get("include_children", False)) for m in struct_map}
```

В блок формирования `docs.append(...)` добавить поля:

```python
            docs.append({
                "date": date_iso,
                "period_label": period["label"],
                "organization": config.ORGANIZATION_NAME,
                "sheet_name": sname,
                "podrazdelenie_uuid": sheet_to_uuid.get(sname, ""),
                "podrazdelenie_name": sheet_to_sname.get(sname, ""),
                "podrazdelenie_stroka": sname.strip(),
                "direction_uuid": sheet_to_diruuid.get(sname, ""),
                "direction_name": sheet_to_dirname.get(sname, ""),
                "include_children": sheet_to_incch.get(sname, False),
                "rows": rows_out,
                "group_totals": group_totals_out,
                "general_totals": general_totals_out,
            })
```

- [ ] **Step 2: Запустить 08**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
..\venv\Scripts\python.exe scripts\08_prepare_documents.py
```

Ожидание: `Wrote .../08_documents_to_import.json (docs: N, rows: M)`.

- [ ] **Step 3: Верификация**

```bash
..\venv\Scripts\python.exe -c "import json; d=json.load(open('data/json/08_documents_to_import.json','r',encoding='utf-8'))['documents']; incch=sum(1 for x in d if x.get('include_children')); wdir=sum(1 for x in d if x.get('direction_uuid')); print(f'docs={len(d)} with_direction={wdir} include_children={incch}')"
```

Ожидание: docs ≈ несколько десятков (по числу period × sheet), with_direction > 0, include_children > 0.

### Task 3.2: Обновить 12_upload_documents.py — шапка А_ОтчетPL

- [ ] **Step 1: Добавить заполнение направления**

В `12_upload_documents.py` в секции `stage = "SetHeader"` (строки ~114-119) после строки `doc.ПодразделениеСтрока = podr_stroka` добавить:

```python
            stage = "SetHeader"
            doc.Организация = org
            doc.ПодразделениеСтрока = podr_stroka
            if d.get("podrazdelenie_uuid"):
                doc.Подразделение = ref_by_uuid(
                    conn, "СтруктураПредприятия", d["podrazdelenie_uuid"]
                )
            # Новые реквизиты: направление + флаг свода
            if d.get("direction_uuid"):
                doc.НаправлениеДеятельности = ref_by_uuid(
                    conn, "НаправленияДеятельности", d["direction_uuid"]
                )
            doc.ВключатьДочерние = bool(d.get("include_children", False))
```

*Важно:* вызов `ref_by_uuid(conn, "НаправленияДеятельности", ...)` работает по той же функции, что уже используется в скрипте.

- [ ] **Step 2: Smoke-test dry-run (без записи)**

```bash
..\venv\Scripts\python.exe scripts\12_upload_documents.py --dry-run --limit 3
```

Должно вывести 3 строки DRY без Exception. (В текущем коде `--dry-run` не вызывает `find_existing_doc`, только печатает sheet_name и row count — этого достаточно.)

*Примечание:* **Боевой перезапуск 12_upload_documents в этой фазе НЕ выполняем.** Скрипт 12 используется для массового импорта новых месяцев; миграцию существующих 37 документов выполняет скрипт 14 (Phase 4). Обновление 12 здесь нужно только для того, чтобы **будущие** импорты (март 2026 и далее) заполняли направление.

### Task 3.3: Обновить 13_fill_rasshifrovka_listov.py — ТЧ Расшифровка

- [ ] **Step 1: Расширить функцию `fill_rasshifrovka` для новых полей**

Заменить блок обработки `if m and m.get("match_type") in AUTO_MATCH_TYPES and m.get("struct_uuid"):` в `13_fill_rasshifrovka_listov.py` (строки 115-122):

```python
        dir_uuid = m.get("direction_uuid") if m else ""
        inc_children = bool(m.get("include_children", False)) if m else False

        if m and m.get("match_type") in AUTO_MATCH_TYPES and m.get("struct_uuid"):
            detail["struct_uuid"] = m["struct_uuid"]
            detail["struct_name"] = m.get("struct_name")
            detail["action"] = "filled"
            if not dry_run:
                uid = conn.NewObject("УникальныйИдентификатор", m["struct_uuid"])
                row.Подразделение = conn.Справочники.СтруктураПредприятия.ПолучитьСсылку(uid)

        # Направление и флаг — ставим всегда, когда есть в mapping
        # (даже для match_type='none', если override явно указал направление)
        detail["direction_uuid"] = dir_uuid or None
        detail["direction_name"] = (m.get("direction_name") if m else None) or None
        detail["include_children"] = inc_children

        if not dry_run and dir_uuid:
            dir_ref_id = conn.NewObject("УникальныйИдентификатор", dir_uuid)
            row.НаправлениеДеятельности = conn.Справочники.НаправленияДеятельности.ПолучитьСсылку(dir_ref_id)
        if not dry_run:
            row.ВключатьДочерние = inc_children

        rows_detail.append(detail)
```

- [ ] **Step 2: Расширить агрегаты в возвращаемом словаре**

В конце `fill_rasshifrovka`, изменить блок подсчёта:

```python
    total = len(rows_detail)
    with_dep = sum(1 for r in rows_detail if r["action"] == "filled")
    with_dir = sum(1 for r in rows_detail if r.get("direction_uuid"))
    incch_count = sum(1 for r in rows_detail if r.get("include_children"))
    return {
        "total": total,
        "with_dep": with_dep,
        "empty_dep": total - with_dep,
        "with_direction": with_dir,
        "include_children_rows": incch_count,
        "rows": rows_detail,
    }
```

- [ ] **Step 3: В `main()` отразить новые счётчики в логе**

В блоке `entries.append({...})` внутри `main()` добавить новые агрегаты:

```python
        entries.append({
            "file": label,
            "excel_path": path,
            "doc_uuid": uuid_str(conn, doc_ref),
            "total": res["total"],
            "with_dep": res["with_dep"],
            "empty_dep": res["empty_dep"],
            "with_direction": res["with_direction"],
            "include_children_rows": res["include_children_rows"],
            "rows": res["rows"],
            "skipped_hidden": hidden,
            "backup_file": backup_file.name if backup_file else None,
        })
```

И в print после fill_rasshifrovka:

```python
        print(f"  Всего: {res['total']}, с Подразделением: {res['with_dep']}, "
              f"пустых: {res['empty_dep']}, с Направлением: {res['with_direction']}, "
              f"свод-строк: {res['include_children_rows']}")
```

### Task 3.4: Dry-run 13 + верификация

- [ ] **Step 1: Dry-run**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
..\venv\Scripts\python.exe scripts\13_fill_rasshifrovka_listov.py --dry-run
```

Ожидание (по 3 документам):
- Декабрь 2025: total=31, with_dep≈25, with_direction ≈ with_dep (предполагая что у всех подразделений из 07_mapping заполнено `А_НаправлениеДеятельности` — см. Phase 0 Task 0.2).
- Январь 2026: total=29, with_dep≈25, with_direction ≈ with_dep.
- Лютий 2026: total=31, with_dep≈26, with_direction ≈ with_dep.
- `include_children_rows` ≈ 4-6 на файл (число сводных листов, присутствующих в конкретном Excel).

Точные значения зависят от того, какие листы присутствуют в каждом Excel (напр. `PL Строительство Свод` есть только в двух из трёх файлов).

- [ ] **Step 2: Проверить dry-run лог**

```
Read C:/Configuration_downloads/BASERP25/_Rarzrabotki/Python/PnL/data/json/13_fill_rasshifrovka_log_dryrun.json (limit 60)
```

Для строки `PL_ЦО`:
```json
{
  "sheet_name": "PL_ЦО",
  "match_type": "manual_override",
  "struct_uuid": "19c7dbef-ca2d-11f0-a2e2-c54425f51b91",
  "struct_name": "ЦО",
  "action": "filled",
  "direction_uuid": "9d021b6f-1ae7-11f0-80dc-00155d235309",
  "direction_name": "ЦО",
  "include_children": true
}
```

Для `Крушинка`:
```json
{
  ...
  "direction_uuid": "<UUID какого-то из 8 направлений>",
  "direction_name": "<имя>",
  "include_children": false
}
```

Для служебных (Метрики, Индекс) — все три поля `null/false`.

- [ ] **Step 3: База НЕ должна измениться — проверить MCP**

```sql
ВЫБРАТЬ ПЕРВЫЕ 5 Ссылка.Номер, ИмяЛиста, ВключатьДочерние, НаправлениеДеятельности
ИЗ Документ.А_РасшифровкаЛистов.Расшифровка
УПОРЯДОЧИТЬ ПО Ссылка.Дата, НомерСтроки
```

Ожидание: у всех строк `ВключатьДочерние = Ложь`, `НаправлениеДеятельности = пустая ссылка` (т.к. 13 ещё не запускался в live).

### Task 3.5: Боевой запуск 13

- [ ] **Step 1: Показать пользователю итог dry-run и попросить подтверждение**

Агент: «Dry-run лог 13 подтверждает корректное заполнение — X строк с направлением из Y всего. Запустить боевой?»

После "да":

- [ ] **Step 2: Live-запуск**

```bash
..\venv\Scripts\python.exe scripts\13_fill_rasshifrovka_listov.py
```

Ожидание: 3 документа обновлены, backup в `data/json/13_backup_*.json` сделан для каждого (перед очисткой).

- [ ] **Step 3: Верификация через MCP**

```sql
ВЫБРАТЬ
    Ссылка.Номер КАК Номер,
    Ссылка.Дата КАК Дата,
    КОЛИЧЕСТВО(*) КАК Строк,
    СУММА(ВЫБОР КОГДА Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд,
    СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
    СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
ИЗ Документ.А_РасшифровкаЛистов.Расшифровка
СГРУППИРОВАТЬ ПО Ссылка.Номер, Ссылка.Дата
УПОРЯДОЧИТЬ ПО Ссылка.Дата
```

Ожидание:
- 3 строки, `Строк = 31/29/31`.
- `СПодразд = 25/25/26` (как было).
- `СНапр ≥ СПодразд` — направление заполнено минимум для всех строк с подразделением.
- `Свод` = число сводных листов в каждом файле (типичный диапазон 4-6: PL_ЦО, PL_ЦО_Строительство, PL_Производство_СВОД и т.п.).

- [ ] **Step 4: Точечная проверка строки-свода**

```sql
ВЫБРАТЬ ПЕРВЫЕ 5
    Ссылка.Номер, ИмяЛиста, Подразделение, НаправлениеДеятельности, ВключатьДочерние
ИЗ Документ.А_РасшифровкаЛистов.Расшифровка
ГДЕ ВключатьДочерние
УПОРЯДОЧИТЬ ПО Ссылка.Дата
```

Для каждой строки `ВключатьДочерние = Истина` должно быть заполнено и `НаправлениеДеятельности` (без исключений).

**Checkpoint 3 — показать пользователю:**
- stdout 13 (счётчики по каждому из 3 файлов);
- результаты двух MCP-queries (агрегат + точечная выборка сводных);
- путь к backup-файлам.

**Acceptance criteria Phase 3:**
- [ ] 13_fill_rasshifrovka_log.json содержит 3 entries, для каждой `with_direction ≥ with_dep` и `include_children_rows` > 0.
- [ ] MCP-query подтверждает заполнение направления ≥ подразделения по каждому документу.
- [ ] Все строки с `ВключатьДочерние = Истина` имеют заполненное `НаправлениеДеятельности`.
- [ ] Backup-файлы (13_backup_00000000N_*.json) созданы для каждого из 3 документов.
- [ ] 12_upload_documents.py смоук-тестирован (dry-run без Exception) — боевой запуск не делаем в этой фазе.

---

# Phase 4 — Миграция существующих А_ОтчетPL (скрипт 14)

**Цель:** обогатить 37 существующих документов `А_ОтчетPL` значениями `НаправлениеДеятельности` и `ВключатьДочерние` — без перепроведения (реквизиты шапки не влияют на движения, документ движений не имеет).

**Files:**
- Create: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\14_migrate_otchetpl_napravlenie.py`

### Task 4.1: Написать 14_migrate_otchetpl_napravlenie.py

- [ ] **Step 1: Создать файл**

Полное содержимое `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL\scripts\14_migrate_otchetpl_napravlenie.py`:

```python
"""Step 14 (one-off): Enrich existing Документ.А_ОтчетPL with НаправлениеДеятельности + ВключатьДочерние.

Стратегия:
1. Взять все непомеченные документы А_ОтчетPL.
2. Для каждого:
   - Посмотреть 07_mapping по ПодразделениеСтрока: если найдено, использовать direction_uuid + include_children оттуда.
   - Иначе: direction_uuid = Подразделение.А_НаправлениеДеятельности (если ссылка заполнена), include_children = False.
   - Если уже заполнено и значения совпадают — skip (idempotent).
3. Записать документ БЕЗ проведения (doc.Записать() без аргументов — реквизиты шапки, не движения).
4. Лог в data/json/14_migrate_otchetpl_log.json.

Флаги:
- --dry-run: не писать, только печатать план.
- --limit N: обработать только первые N.
"""
import argparse
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


def load_mapping_by_podr_stroka():
    raw = json.loads(
        (config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8")
    )
    out = {}
    for m in raw["mappings"]:
        out[m["sheet_name"]] = {
            "direction_uuid": m.get("direction_uuid") or "",
            "direction_name": m.get("direction_name") or "",
            "include_children": bool(m.get("include_children", False)),
        }
    return out


def fetch_all_otchetpl_refs(conn):
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ Ссылка
    ИЗ Документ.А_ОтчетPL
    ГДЕ НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Дата
    """
    tz = q.Выполнить().Выгрузить()
    refs = []
    for i in range(tz.Количество()):
        refs.append(tz.Получить(i).Ссылка)
    return refs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int, default=None)
    args = ap.parse_args()

    mode = "DRY-RUN" if args.dry_run else "LIVE"
    print(f"=== 14 Migrate А_ОтчетPL направление ({mode}) ===\n")

    conn = connect_erp()
    mapping = load_mapping_by_podr_stroka()
    print(f"Mapping: {len(mapping)} листов")

    refs = fetch_all_otchetpl_refs(conn)
    if args.limit:
        refs = refs[: args.limit]
    print(f"Документов к обработке: {len(refs)}\n")

    empty_napr_ref = conn.Справочники.НаправленияДеятельности.ПустаяСсылка()

    log = []
    updated = skipped = failed = 0

    for idx, ref in enumerate(refs, 1):
        try:
            obj = ref.ПолучитьОбъект()
            num = str(obj.Номер).strip()
            date_str = str(obj.Дата)[:10]
            podr_str = str(obj.ПодразделениеСтрока) if obj.ПодразделениеСтрока else ""

            # 1) Решить какое направление и флаг должны быть
            m = mapping.get(podr_str)
            target_dir_uuid = ""
            target_dir_name = ""
            target_incch = False
            source = ""

            if m and m.get("direction_uuid"):
                target_dir_uuid = m["direction_uuid"]
                target_dir_name = m["direction_name"]
                target_incch = m["include_children"]
                source = "mapping"
            elif obj.Подразделение and not obj.Подразделение.Пустая():
                napr = obj.Подразделение.А_НаправлениеДеятельности
                if napr and not napr.Пустая():
                    target_dir_uuid = uuid_str(conn, napr)
                    target_dir_name = str(napr.Наименование)
                    source = "podrazdelenie"

            # 2) Текущее состояние
            cur_napr = obj.НаправлениеДеятельности
            cur_napr_uuid = "" if (cur_napr is None or cur_napr.Пустая()) else uuid_str(conn, cur_napr)
            cur_incch = bool(obj.ВключатьДочерние)

            # 3) Нужно ли обновлять?
            need_dir = target_dir_uuid and target_dir_uuid.lower() != cur_napr_uuid.lower()
            need_incch = target_incch != cur_incch
            action = "skip"

            if not need_dir and not need_incch:
                skipped += 1
                log.append({
                    "num": num, "date": date_str, "podr_stroka": podr_str,
                    "source": source or "current_ok",
                    "action": "skip",
                    "current_direction": cur_napr_uuid,
                    "current_include_children": cur_incch,
                    "target_direction": target_dir_uuid,
                    "target_include_children": target_incch,
                })
                print(f"  [{idx:3}/{len(refs)}] SKIP {num} / {podr_str}  (napr={cur_napr_uuid or '—'} incch={cur_incch})")
                continue

            action = "update"

            if not args.dry_run:
                if target_dir_uuid:
                    uid = conn.NewObject("УникальныйИдентификатор", target_dir_uuid)
                    obj.НаправлениеДеятельности = conn.Справочники.НаправленияДеятельности.ПолучитьСсылку(uid)
                else:
                    obj.НаправлениеДеятельности = empty_napr_ref
                obj.ВключатьДочерние = target_incch
                obj.ОбменДанными.Загрузка = True  # пропустить валидацию ПередЗаписью при записи
                obj.Записать()  # без проведения — реквизиты шапки

            updated += 1
            log.append({
                "num": num, "date": date_str, "podr_stroka": podr_str,
                "source": source,
                "action": "update" if not args.dry_run else "update-dry",
                "current_direction": cur_napr_uuid,
                "current_include_children": cur_incch,
                "target_direction": target_dir_uuid,
                "target_direction_name": target_dir_name,
                "target_include_children": target_incch,
            })
            tag = "DRY" if args.dry_run else "UPD"
            print(f"  [{idx:3}/{len(refs)}] {tag} {num} / {podr_str}  "
                  f"napr: {cur_napr_uuid or '—'} → {target_dir_uuid or '—'}  incch: {cur_incch} → {target_incch}  ({source})")

        except Exception as ex:
            failed += 1
            err = traceback.format_exc(limit=2)
            log.append({
                "num": num if 'num' in locals() else "?",
                "action": "error",
                "error": str(ex),
                "trace": err[:500],
            })
            print(f"  [{idx:3}/{len(refs)}] ERR {ex}")

    suffix = "_dryrun" if args.dry_run else ""
    out = config.JSON_DIR / f"14_migrate_otchetpl_log{suffix}.json"
    out.write_text(
        json.dumps({
            "run_at": datetime.now().isoformat(timespec="seconds"),
            "mode": mode,
            "total": len(refs),
            "updated": updated,
            "skipped": skipped,
            "failed": failed,
            "entries": log,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nLog: {out}  (updated={updated}, skipped={skipped}, failed={failed})")


if __name__ == "__main__":
    main()
```

Ключевые моменты:
1. `ОбменДанными.Загрузка = True` — обходит кастомную валидацию в `ПередЗаписью` (на этапе миграции мы гарантируем корректность данных через python-логику).
2. `Записать()` без аргументов — реквизиты шапки, не требует проведения (документ не имеет движений).
3. Идемпотентность — если значения уже совпадают, skip.
4. Источник (`source`) в логе: `mapping` (из 07) или `podrazdelenie` (из А_НаправлениеДеятельности подразделения).

### Task 4.2: Dry-run 14

- [ ] **Step 1: Запуск**

```bash
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\PnL
..\venv\Scripts\python.exe scripts\14_migrate_otchetpl_napravlenie.py --dry-run
```

Ожидание: `total=37 updated=37 skipped=0 failed=0` (при первом прогоне все 37 должны быть обновлены).

- [ ] **Step 2: Анализ лога**

```
Read C:/Configuration_downloads/BASERP25/_Rarzrabotki/Python/PnL/data/json/14_migrate_otchetpl_log_dryrun.json (limit 80)
```

Проверить распределение источников:
```bash
..\venv\Scripts\python.exe -c "import json; d=json.load(open('data/json/14_migrate_otchetpl_log_dryrun.json','r',encoding='utf-8'))['entries']; from collections import Counter; print('source:', Counter(e.get('source') for e in d))"
```

Ожидание: `{'mapping': X, 'podrazdelenie': Y}` — X + Y = 37 (ни одного None/''). Если есть None/'' — это документы у которых ни маппинг не нашёлся, ни Подразделение не заполнено → требует ручного анализа до live-запуска.

- [ ] **Step 3: База НЕ должна измениться**

```sql
ВЫБРАТЬ
    СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
    СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
ИЗ Документ.А_ОтчетPL
ГДЕ НЕ ПометкаУдаления
```

Ожидание: `СНапр=0, Свод=0` (dry-run базу не трогал).

### Task 4.3: Live-запуск 14

- [ ] **Step 1: Показать пользователю dry-run итог + попросить подтверждение**

Агент: «Dry-run 14: 37 документов к обновлению, источник — {mapping: X, podrazdelenie: Y}, 0 ошибок. Запустить боевой?»

После «да»:

- [ ] **Step 2: Live**

```bash
..\venv\Scripts\python.exe scripts\14_migrate_otchetpl_napravlenie.py
```

Ожидание: те же 37 updated, 0 failed.

- [ ] **Step 3: Верификация**

```sql
ВЫБРАТЬ
    КОЛИЧЕСТВО(*) КАК Всего,
    СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
    СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
ИЗ Документ.А_ОтчетPL
ГДЕ НЕ ПометкаУдаления
```

Ожидание: `Всего=37`, `СНапр` близко к 37 (все, у которых был Подразделение или в 07_mapping), `Свод` > 0 (число документов по сводным листам — типично 6-12).

- [ ] **Step 4: Точечная выборка**

```sql
ВЫБРАТЬ ПЕРВЫЕ 10
    Номер, Дата, ПодразделениеСтрока, НаправлениеДеятельности.Наименование КАК Направление, ВключатьДочерние
ИЗ Документ.А_ОтчетPL
ГДЕ НЕ ПометкаУдаления
    И ВключатьДочерние
УПОРЯДОЧИТЬ ПО Дата
```

Ожидание: у каждого документа с `ВключатьДочерние=Истина` направление заполнено.

### Task 4.4: Idempotency test

- [ ] **Step 1: Повторный dry-run**

```bash
..\venv\Scripts\python.exe scripts\14_migrate_otchetpl_napravlenie.py --dry-run
```

Ожидание: `updated=0, skipped=37` — после live-запуска повторный проход ничего не меняет.

**Checkpoint 4 — показать пользователю:**
- dry-run первого прохода: 37 updated;
- live-прогон: 37 updated, 0 failed;
- MCP-query после live: Всего=37, СНапр≈37, Свод>0;
- dry-run второй раз: 0 updated (идемпотентно).

**Acceptance criteria Phase 4:**
- [ ] 14_migrate_otchetpl_log.json показывает 37 updated, 0 failed.
- [ ] MCP-query: `СНапр` ≥ числа документов, у которых было заполнено Подразделение (т.е. ≈ 37).
- [ ] Все документы с `ВключатьДочерние = Истина` имеют заполненное направление.
- [ ] Повторный dry-run отчитывается `updated=0, skipped=37` (идемпотентность).

---

# Phase 5 — Функциональное тестирование + документация

**Цель:** убедиться что UI-поведение соответствует спеку §6.3 (автоподстановка, кнопка, сверка с АнализДоходовРасходов) и обновить промт 13.

### Task 5.1: Ручной тест — автоподстановка направления

- [ ] **Step 1: Инструкция пользователю**

> Открыть в клиенте 1С существующий документ А_ОтчетPL (любой, напр. из декабря 2025 с `ПодразделениеСтрока = "Крушинка"`). Очистить поле `Подразделение`, очистить `НаправлениеДеятельности`. Выбрать в поле `Подразделение` = Крушинка. Ожидание: `НаправлениеДеятельности` заполнилось автоматически (значением `Крушинка.А_НаправлениеДеятельности`). Отменить изменения (Esc/закрыть без записи).

Записать результат (ОК / ошибка) в checkpoint.

### Task 5.2: Ручной тест — кнопка «Открыть сверку с ЕРП» в обычном режиме

- [ ] **Step 1: Инструкция пользователю**

> Открыть документ А_ОтчетPL (Крушинка, любой месяц). Убедиться что `ВключатьДочерние = Ложь`. Нажать «Открыть сверку с ЕРП».
>
> Ожидание: отчёт открывается с вариантом `А_ОтчетPL_РасшифровкаДокументаОтчетPL` (заголовок «Расшифровка документа PL»), отбор по `Подразделение = Крушинка`, период = месяц документа.

### Task 5.3: Ручной тест — кнопка в режиме «по направлению»

- [ ] **Step 1: Инструкция пользователю**

> Открыть документ А_ОтчетPL с `ПодразделениеСтрока = "Техника"` (или `PL_ЦО_Строительство`) — после скрипта 14 у него должно быть `ВключатьДочерние = Истина`, `НаправлениеДеятельности = Логистика` (или Строительство). Нажать «Открыть сверку с ЕРП».
>
> Ожидание: отчёт открывается с вариантом «Расшифровка по направлению», отбор по `НаправлениеДеятельности = Логистика`, в колонках — все дочерние подразделения (MAN №3, №4, №5, Экспедирование, Логистика).

### Task 5.4: Сверка с АнализДоходовРасходов (ключевой business-test)

- [ ] **Step 1: Инструкция пользователю**

> Открыть отчёт `АнализДоходовРасходов` → вариант «Направление-Подразделение-рабочая». Настроить:
> - Период: Декабрь 2025 (совпадает с документом А_РасшифровкаЛистов №000000001).
> - Отбор: НаправлениеДеятельности = Логистика.
>
> Запомнить итоговую сумму по направлению (колонка «Итого»).
>
> Открыть документ `А_ОтчетPL` с `ПодразделениеСтрока = "Техника"`, период декабрь 2025 → нажать «Открыть сверку с ЕРП». В новом отчёте «Расшифровка по направлению» — найти колонку «СуммаЕРП» → сравнить с «Итого» из АнализДоходовРасходов.
>
> Ожидание: значения **совпадают** (до копейки) — это подтверждает что отбор по `НаправлениеДеятельности` агрегирует те же движения что и типовой отчёт.

### Task 5.5: Обновить промт 13 — новые поля ТЧ

- [ ] **Step 1: Обновить prompt_fill_rasshifrovka.md**

В `_Rarzrabotki/Python/PnL/docs/prompt_fill_rasshifrovka.md` внести точечные правки:

1. В разделе «Точные метаданные Документ.А_РасшифровкаЛистов → Табличная часть Расшифровка» добавить в таблицу реквизиты:

```markdown
| `НаправлениеДеятельности` | `СправочникСсылка.НаправленияДеятельности` | UUID направления из 07_mapping (direction_uuid) |
| `ВключатьДочерние` | `Булево` | True для сводных листов (PL_ЦО, Техника и т.п.) |
```

2. В разделе «Шаг 4 — Заполнить ТЧ для каждого документа» — в примере кода после установки `row.Подразделение` добавить установку `row.НаправлениеДеятельности` и `row.ВключатьДочерние` (код уже в скрипте, промт просто отражает реальность).

3. В «Acceptance criteria» добавить пункты:
   - все строки с `ВключатьДочерние=Истина` имеют заполненное `НаправлениеДеятельности`;
   - в логе 13_fill_rasshifrovka_log.json появились поля `with_direction` и `include_children_rows`.

### Task 5.6: Финальный аудит метаданных + данных

- [ ] **Step 1: Агент запускает финальный MCP-скрипт**

```sql
-- 1) Проверить реквизиты добавлены
-- (автоматизировать через get_metadata_structure, см. Task 1.8)

-- 2) Покрытие А_РасшифровкаЛистов
ВЫБРАТЬ Ссылка.Номер КАК Номер,
    КОЛИЧЕСТВО(*) КАК Всего,
    СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
    СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
ИЗ Документ.А_РасшифровкаЛистов.Расшифровка
СГРУППИРОВАТЬ ПО Ссылка.Номер

-- 3) Покрытие А_ОтчетPL
ВЫБРАТЬ
    КОЛИЧЕСТВО(*) КАК Всего,
    СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
    СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
ИЗ Документ.А_ОтчетPL
ГДЕ НЕ ПометкаУдаления
```

- [ ] **Step 2: Записать итог в финальный отчёт**

Создать `data/json/99_direction_migration_summary.json` со структурой: before/after (из 00_baseline + текущий прогон), timestamp, список изменений.

**Checkpoint 5 (финальный) — показать пользователю:**
- отчёт по всем ручным тестам (5.1-5.4) с подтверждением «ОК» по каждому;
- финальный MCP-аудит (3 queries);
- сравнение сумм «АнализДоходовРасходов vs А_ОтчетPL по направлению» = совпадают;
- diff prompt_fill_rasshifrovka.md;
- `99_direction_migration_summary.json` создан.

**Acceptance criteria Phase 5 (§6.3 спека):**
- [ ] Автоподстановка НаправлениеДеятельности при выборе Подразделения работает.
- [ ] Кнопка «Открыть сверку с ЕРП» в обычном режиме открывает вариант `А_ОтчетPL_РасшифровкаДокументаОтчетPL` с отбором по Подразделению.
- [ ] Кнопка в режиме ВключатьДочерние открывает `А_ОтчетPL_РасшифровкаПоНаправлению` с отбором по Направлению и разворотом по подразделениям в колонках.
- [ ] Сводная цифра А_ОтчетPL по направлению совпадает с «Итого» АнализДоходовРасходов за тот же период/направление.
- [ ] `prompt_fill_rasshifrovka.md` отражает новые поля.

---

## Финальные acceptance criteria (§6 спека)

### 6.1 Структурные
- [ ] `Документ.А_РасшифровкаЛистов.Расшифровка` содержит `НаправлениеДеятельности`, `ВключатьДочерние` (Phase 1).
- [ ] `Документ.А_ОтчетPL` (шапка) содержит те же реквизиты (Phase 1).
- [ ] `Отчет.А_ОтчетPL` содержит variant `А_ОтчетPL_РасшифровкаПоНаправлению` (Phase 1).

### 6.2 Данные
- [ ] В 3 документах А_РасшифровкаЛистов ≥95% непустых строк имеют `НаправлениеДеятельности` (Phase 3).
- [ ] У всех строк с `ВключатьДочерние=Истина` заполнено `НаправлениеДеятельности` (Phase 3).
- [ ] Все 37 А_ОтчетPL имеют `НаправлениеДеятельности` (Phase 4).

### 6.3 Функциональные
- Phase 5 Tasks 5.1-5.4.

### 6.4 Идемпотентность
- [ ] Повторный 13_fill_rasshifrovka_listov.py не меняет данные (Phase 3).
- [ ] Повторный 14_migrate_otchetpl_napravlenie.py отчитывается 0 updated (Phase 4 Task 4.4).

---

## Риски и откат (из §9 спека)

| Риск | Митигация в плане |
|---|---|
| Переиндексация при добавлении реквизита к А_ОтчетPL (37 строк — не много, но пропагандируем осторожность) | CF backup в Phase 0 Task 0.3; обновление БД в нерабочее время — Task 1.7. |
| ВключатьДочерние=Истина без НаправлениеДеятельности — человеческая ошибка | ПередЗаписью валидация — Task 1.3. |
| direction_uuid в 07 расходится с Подразделение.А_НаправлениеДеятельности | 07 использует override в приоритете; в логе 13 поле `direction_uuid` per-row — можно проверить. Спека §9 допускает этот случай. |
| Отчёт «ПоНаправлению» показывает иначе, чем АнализДоходовРасходов | Ручная сверка — Task 5.4 (golden case: декабрь 2025, Логистика). |
| Служебные/общие листы (PL_Свод, ЦО_Підсумовування) без направления | Остаются с пустым направлением и `ВключатьДочерние=Ложь` → кнопка использует старый вариант отчёта по Подразделению (§9 последний риск). |

**Откат:**
- Phase 1 → восстановить `.cf` backup из Phase 0 Task 0.3.
- Phase 3 → восстановить строки из `13_backup_00000000N_*.json` вручную в клиенте.
- Phase 4 → повторный прогон 14 с вручную очищенными полями (или точечное восстановление по полям `current_direction` / `current_include_children` из 14_migrate_otchetpl_log.json).

---

## Execution handoff

**Plan complete. Two execution options:**

**1. Subagent-Driven (recommended)** — агент диспатчит subagent на каждую задачу, ревью между; быстрая итерация. Однако Phase 1 требует пользователя (Designer/EDT), поэтому первые 7 задач Phase 1 — handoff пользователю, а верификация (Task 1.8) агентом.

**2. Inline execution** — через `superpowers:executing-plans`. Фазы идут последовательно, чек-поинт после каждой; пользователь подтверждает переход.

Рекомендация: **Inline execution** (план рассчитан на живую коллаборацию с пользователем: Phase 0 → пользователь делает CF-backup → Phase 1 → пользователь правит метаданные → Phase 2-4 → пользователь даёт «далее» → Phase 5 → ручные тесты).

После утверждения плана — дождаться команды "выполняй" и пошагово пройти Phase 0.
