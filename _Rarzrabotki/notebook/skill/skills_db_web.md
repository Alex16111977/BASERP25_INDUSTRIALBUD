# Управление базами данных 1С, Веб-публикация и Веб-тестирование

> Управление информационными базами, веб-публикация через Apache, тестирование через веб-клиент.
> Источник: cc-1c-skills (Nikolay-Shirokov/cc-1c-skills)

---

## Часть 1: Управление проектом (.v8-project.json)

# Конфигурация проекта (.v8-project.json)

Файл `.v8-project.json` — единый конфиг проекта для всех навыков Claude Code. Хранит пути к платформе 1С, список баз данных и настройки инструментов (Apache, ffmpeg, TTS).

Размещается в корне проекта (рядом с `.git/`). Создаётся навыком `/db-list add` или вручную.

> **Безопасность**: файл содержит секреты (пароли баз данных, API-ключи TTS) и добавлен в `.gitignore` — он не попадает в репозиторий. Каждый разработчик заводит свой `.v8-project.json` локально.

## Полная схема

```jsonc
{
  // === Платформа ===
  "v8path": "C:\\Program Files\\1cv8\\8.3.25.1257\\bin",

  // === Базы данных ===
  "databases": [
    {
      "id": "dev",                          // уникальный идентификатор
      "name": "Разработка",                 // отображаемое имя
      "type": "file",                       // "file" или "server"
      "path": "C:\\Bases\\MyApp_Dev",       // каталог (для file)
      "user": "Admin",                      // пользователь 1С
      "password": "",                       // пароль
      "aliases": ["dev", "разработка"],     // альтернативные имена
      "branches": ["dev", "feature/*"],     // привязка к Git-веткам
      "configSrc": "C:\\WS\\myapp\\cfsrc",  // каталог XML-выгрузки конфигурации
      "webUrl": "http://localhost:8081/dev"  // URL веб-клиента (для /web-test)
    },
    {
      "id": "test",
      "name": "Тестовая",
      "type": "server",                     // серверная база
      "server": "srv01",                    // адрес сервера 1С
      "ref": "MyApp_Test",                  // имя базы на сервере
      "user": "Admin",
      "password": "123",
      "aliases": ["test", "тест"]
    }
  ],
  "default": "dev",

  // === Инструменты ===
  "webPath": "C:\\tools\\apache24",                  // каталог Apache
  "ffmpegPath": "C:\\tools\\ffmpeg\\bin\\ffmpeg.exe", // путь к ffmpeg
  "tts": {                                            // настройки озвучки
    "provider": "edge",
    "voice": "ru-RU-DmitryNeural"
  }
}
```

## Корневые поля

| Поле | Тип | Обяз. | По умолчанию | Описание | Кто заполняет |
|------|-----|:-----:|-------------|----------|---------------|
| `v8path` | string | да | — | Путь к каталогу `bin` платформы 1С | `/db-list add` или руками |
| `databases` | array | да | — | Список баз данных | `/db-list add` |
| `default` | string | нет | — | `id` базы по умолчанию | `/db-list` |
| `webPath` | string | нет | `tools/apache24` | Каталог Apache HTTP Server | Руками |
| `ffmpegPath` | string | нет | `tools/ffmpeg/bin/ffmpeg.exe` | Путь к ffmpeg | Руками |
| `tts` | object | нет | Edge TTS, DmitryNeural | Настройки озвучки видео | Руками |

## Базы данных (`databases[]`)

| Поле | Тип | Обяз. | Описание | Кто заполняет |
|------|-----|:-----:|----------|---------------|
| `id` | string | да | Уникальный идентификатор | `/db-list add` |
| `name` | string | да | Отображаемое имя | `/db-list add` |
| `type` | `"file"` / `"server"` | да | Тип подключения | `/db-list add` |
| `path` | string | для file | Каталог файловой базы | `/db-list add` |
| `server` | string | для server | Адрес сервера 1С | `/db-list add` |
| `ref` | string | для server | Имя базы на сервере | `/db-list add` |
| `user` | string | нет | Пользователь 1С | `/db-list add` или руками |
| `password` | string | нет | Пароль | `/db-list add` или руками |
| `aliases` | string[] | нет | Альтернативные имена для обращения к базе | `/db-list add` или руками |
| `branches` | string[] | нет | Git-ветки или glob-паттерны (`release/*`, `feature/*`) | Руками |
| `configSrc` | string | нет | Каталог XML-выгрузки конфигурации | Руками |
| `webUrl` | string | нет | URL веб-клиента для `/web-test` | Руками |

### Разрешение базы

Все навыки `/db-*`, `/epf-build`, `/epf-dump`, `/erf-build`, `/erf-dump`, `/web-publish` используют единый алгоритм:

1. Если пользователь указал **параметры подключения** (путь, сервер) — используются напрямую
2. Если указал **базу по имени** — поиск: `id` → `aliases` (с учётом морфологии) → `name` (нечёткое)
3. Если **не указал** — сопоставление текущей ветки Git с `branches` (точно или по glob-паттерну)
4. Fallback на `default`
5. Если не найдено — Claude спросит пользователя
6. Если база не зарегистрирована — Claude предложит `/db-list add`

## Настройки инструментов

### `webPath` — Apache HTTP Server

Путь к каталогу Apache. Используется навыками `/web-publish`, `/web-info`, `/web-stop`, `/web-unpublish`.

Если не задан — ищется в `tools/apache24` от корня проекта. При первом вызове `/web-publish` Apache скачивается автоматически.

Подробнее — в [гайде по веб-публикации](web-guide.md).

### `ffmpegPath` — ffmpeg

Путь к исполняемому файлу ffmpeg. Используется навыком `/web-test` для записи видео.

Если не задан — ищется по порядку:
1. `tools/ffmpeg/bin/ffmpeg.exe` (от корня проекта)
2. `ffmpeg` в системном PATH

Подробнее — в [гайде по записи видео](web-test-recording-guide.md).

### `tts` — озвучка видеоинструкций

| Поле | Тип | По умолчанию | Описание |
|------|-----|-------------|----------|
| `provider` | string | `"edge"` | Провайдер: `"edge"`, `"elevenlabs"`, `"openai"` |
| `voice` | string | `"ru-RU-DmitryNeural"` | Голос (имя или ID в зависимости от провайдера) |
| `apiKey` | string | — | API-ключ (для elevenlabs, openai) |
| `apiUrl` | string | — | URL сервиса (для openai-совместимых) |
| `model` | string | — | Модель (для openai) |

Подробнее о выборе провайдера и голосов — в [гайде по записи видео](web-test-recording-guide.md#доступные-голоса-и-провайдеры).

### `webUrl` — URL веб-клиента (per-database)

URL для открытия базы в браузере через `/web-test`. Задаётся в записи конкретной базы.

Если не задан — `/web-test` берёт URL из активной веб-публикации (`/web-publish`).

Полезно, если веб-клиент доступен по нестандартному адресу (другой порт, внешний сервер, reverse proxy).

## Минимальный пример

```json
{
  "v8path": "C:\\Program Files\\1cv8\\8.3.25.1257\\bin",
  "databases": [
    {
      "id": "dev",
      "name": "Разработка",
      "type": "file",
      "path": "C:\\Bases\\MyApp"
    }
  ]
}
```

## Полный пример

```json
{
  "v8path": "C:\\Program Files\\1cv8\\8.3.25.1257\\bin",
  "databases": [
    {
      "id": "dev",
      "name": "Разработка",
      "type": "file",
      "path": "C:\\Bases\\MyApp_Dev",
      "user": "Admin",
      "password": "",
      "aliases": ["dev", "разработка"],
      "branches": ["dev", "develop", "feature/*"],
      "configSrc": "C:\\WS\\myapp\\cfsrc",
      "webUrl": "http://localhost:8081/dev"
    },
    {
      "id": "prod",
      "name": "Рабочая",
      "type": "server",
      "server": "srv01",
      "ref": "MyApp_Prod",
      "user": "Admin",
      "password": "secret",
      "aliases": ["prod", "рабочая", "боевая"],
      "branches": ["main", "release/*"]
    }
  ],
  "default": "dev",
  "webPath": "C:\\tools\\apache24",
  "ffmpegPath": "C:\\tools\\ffmpeg\\bin\\ffmpeg.exe",
  "tts": {
    "provider": "edge",
    "voice": "ru-RU-DmitryNeural"
  }
}
```

## Связанные навыки

- [Базы данных](db-guide.md) — `/db-list`, `/db-create`, `/db-load-xml`, `/db-dump-xml` и другие
- [Веб-публикация](web-guide.md) — `/web-publish`, `/web-info`, `/web-stop`
- [Тестирование в браузере](web-test-guide.md) — `/web-test`
- [Запись видеоинструкций](web-test-recording-guide.md) — запись видео, субтитры, озвучка


---

## Часть 2: Веб-тестирование через Playwright

# Тестирование через веб-клиент 1С

Навык `/web-test` автоматизирует действия в веб-клиенте 1С через Playwright — навигация по разделам, заполнение форм, чтение таблиц и отчётов, фильтрация списков. Замыкает цикл: правка исходников → загрузка → обновление → публикация → **автоматическое тестирование**.

## Навык

| Навык | Скрипт | Описание |
|-------|:------:|----------|
| `/web-test` | `.mjs` (Node.js) | Автоматизация 1С через браузер — навигация, формы, таблицы, отчёты |

## Предусловия

- База опубликована через Apache (`/web-publish`)
- Node.js 18+ установлен
- Зависимости установлены: `cd .claude/skills/web-test/scripts && npm install`

## Рабочий цикл

```
/web-publish → /web-test → результат
     ↑                          |
     └── правки → /db-load-xml → /db-update ──┘
```

## Сценарии использования

### Навигация и чтение данных

```
> Открой базу erp в браузере, перейди в раздел Склад и покажи какие команды там есть
```

Claude откроет браузер, перейдёт в раздел и покажет список команд.

```
> Открой список поступлений товаров и покажи первые 10 строк
```

Claude откроет список и прочитает таблицу.

### Поиск и открытие элементов

```
> Найди в списке номенклатуры товар "Вентилятор" и открой его карточку
```

Claude отфильтрует список, откроет найденный элемент двойным кликом и прочитает реквизиты формы.

```
> Открой справочник Контрагенты и найди "Торговый дом"
```

Claude может работать с иерархическими справочниками — поиск автоматически сглаживает иерархию.

### Создание документа

```
> Создай заказ клиента: организация "Андромеда Плюс", контрагент "Торговый дом Комплексный",
> добавь строку: номенклатура "Вентилятор", количество 5
```

Claude откроет форму создания, заполнит шапку и добавит строку в табличную часть.

### Работа с отчётами

```
> Открой отчёт "Остатки и доступность товаров",
> установи отбор Склад = "Склад бытовой техники", сформируй и прочитай результат
```

Claude заполнит фильтры отчёта по человекочитаемым именам (не надо знать технические имена DCS), нажмёт "Сформировать" и прочитает структурированные данные: заголовки, строки, итоги.

### Сравнение данных

```
> Сформируй отчёт по остаткам для "Склад бытовой техники" и "Западный склад",
> сравни итоги по доступным товарам
```

Claude напишет сценарий, который сформирует отчёт дважды с разными фильтрами и сравнит результаты.

### Проверка после загрузки расширения

```
> Загрузи расширение ТестОшибки и проверь через браузер, что при создании заказа клиента
> появляется ошибка "Тестовая ошибка из расширения"
```

Claude загрузит расширение через `/db-load-xml`, затем через `/web-test` откроет форму и проверит ожидаемое поведение.

### Открытие внешней обработки

```
> Открой обработку build/РедакторДвижений.epf в веб-клиенте и покажи что на форме
```

Claude откроет EPF через Ctrl+O, автоматически обработает диалог безопасности (если есть) и прочитает форму.

### Пошаговая отладка

```
> Запусти браузер на базе erp
> Перейди в раздел Продажи
> Посмотри что на форме
> Открой первый заказ
> Какие реквизиты заполнены?
```

Claude будет выполнять команды по одной, показывая состояние формы между шагами.

## Режимы работы

### Автономный режим (run)

Одна команда: открывает браузер → логинится → выполняет сценарий → закрывает браузер → завершает процесс. Не оставляет висящих процессов.

```bash
RUN=".claude/skills/web-test/scripts/run.mjs"
node $RUN run http://localhost:8081/erp scenario.js
```

Claude пишет `.js` файл со сценарием и запускает его. Ответ — JSON:
```json
{ "ok": true, "output": "...console.log output...", "elapsed": 12.3 }
```

При ошибке — автоматический скриншот (пока модалка ещё видна) и стек вызова:
```json
{ "ok": false, "error": "Тестовая проверка: запись запрещена", "screenshot": "error-shot.png",
  "stack": { "raw": "...", "entries": [{"location": "Модуль(4)", "code": "ВызватьИсключение..."}] } }
```
Стек извлекается автоматически — через OpenReport (платформенные исключения) или "О программе" → "Информация для техподдержки" (ВызватьИсключение).

### Интерактивный режим (start/exec/stop)

Браузер остаётся открытым между командами. Состояние (открытые вкладки, формы) сохраняется.

```bash
node $RUN start http://localhost:8081/erp    # запустить сессию (фоновый процесс)
cat <<'SCRIPT' | node $RUN exec -            # выполнить скрипт
await navigateSection('Продажи');
SCRIPT
node $RUN shot current-state.png             # скриншот
node $RUN stop                               # завершить сессию
```

### Когда какой

| Режим | Когда использовать |
|-------|-------------------|
| Автономный (`run`) | Готовый сценарий целиком, субагенты, CI |
| Интерактивный (`start/exec`) | Пошаговое исследование, отладка, разговор с пользователем |

## Пример: автономный сценарий

Сравнение остатков по двум складам — один файл, один запуск:

```js
// scenario-compare-stocks.js

// 1. Открыть отчёт
await navigateSection('Склад и доставка');
await openCommand('Отчеты по складу');
await clickElement('Остатки и доступность товаров', { dblclick: true });

// 2. Первый склад
await fillFields({ 'Склад': 'Склад бытовой техники' });
await clickElement('Сформировать');
await wait(5);
const report1 = await readSpreadsheet();
console.log('=== Склад бытовой техники ===');
console.log('Строк:', report1.data?.length, '| Доступно:', report1.totals?.['Доступно']);

// 3. Второй склад
await fillFields({ 'Склад': 'Западный склад' });
await clickElement('Сформировать');
await wait(5);
const report2 = await readSpreadsheet();
console.log('=== Западный склад ===');
console.log('Строк:', report2.data?.length, '| Доступно:', report2.totals?.['Доступно']);

// 4. Сравнение
const parse = s => parseFloat((s || '0').replace(/\s/g, '').replace(',', '.'));
const diff = parse(report1.totals?.['Доступно']) - parse(report2.totals?.['Доступно']);
console.log('Разница:', diff.toFixed(0));

await closeForm({ save: false });
```

Запуск: `node $RUN run http://localhost:8081/erp scenario-compare-stocks.js`

### Расшифровка отчёта

```js
// 1. Сформировать отчёт
await clickElement('Сформировать');
await wait(5);
const report = await readSpreadsheet();

// 2. Двойной клик по ячейке → диалог "Выбор поля"
await clickElement({ row: 0, column: 'К6' }, { dblclick: true });

// 3. Выбрать поле расшифровки
await clickElement('Регистратор');
await clickElement('Выбрать');
await wait(10);

// 4. Прочитать результат
const drilldown = await readSpreadsheet();
console.log('Расшифровка:', JSON.stringify(drilldown.rows));
```

## API

Все функции доступны как глобальные переменные в скриптах. `console.log()` выводит данные в ответ.

### Навигация

| Функция | Описание | Возвращает |
|---------|----------|------------|
| `navigateSection(name)` | Перейти в раздел (fuzzy match) | `{ sections, commands }` |
| `openCommand(name)` | Открыть команду из панели функций | form state |
| `navigateLink(path)` | Открыть по пути метаданных (`Документ.ЗаказКлиента`) | form state |
| `openFile(path)` | Открыть внешнюю обработку/отчёт (EPF/ERF) через «Файл → Открыть» | form state |
| `switchTab(name)` | Переключить открытую вкладку | form state |

### Чтение

| Функция | Описание | Возвращает |
|---------|----------|------------|
| `getFormState()` | Структура формы: поля, кнопки, таблица, фильтры, состояние окон | `{ form, formCount, openForms, fields, buttons, tabs, table, filters, reportSettings? }` |
| `readTable({ maxRows?, offset? })` | Данные таблицы с пагинацией | `{ columns, rows: [{col: val}], total }` |
| `readSpreadsheet()` | Результат отчёта | `{ title?, headers?, data?, totals?, total }` |
| `getSections()` | Разделы и команды | `{ activeSection, sections, commands }` |
| `getPageState()` | Разделы + открытые вкладки | `{ activeSection, activeTab, sections, tabs }` |

#### getFormState — подробнее

Основной способ «увидеть» что на экране:

- **form** — номер активной формы, `null` когда ничего не открыто (десктоп)
- **formCount** — количество открытых форм. `0` = десктоп. Работает даже если панель открытых окон скрыта
- **openForms** — `[0, 1, 2]` — номера всех открытых форм в DOM
- **modal** — `true` когда активная форма — модальный диалог, блокирующий интерфейс
- **openTabs** — `[{ name, active? }]` из панели открытых окон (только когда панель включена в настройках 1С)
- **fields** — `[{ name, value, label?, actions?, required? }]`. `actions` = select/clear/open. `required: true` = незаполненное обязательное поле
- **table** — `{ name, columns, rowCount }` (метаданные; для данных — `readTable()`)
- **reportSettings** — DCS-фильтры в читаемом виде: `[{ name: "Склад", enabled: true, value: "..." }]`
- **errorModal** — 1С показала ошибку
- **confirmation** — диалог Да/Нет, вызовите `clickElement('Да')` или `clickElement('Нет')`
- **platformDialogs** — `[{ type, title }]` — платформенные диалоги (О программе, Информация для техподдержки). Невидимы для обычного определения форм, но блокируют интерфейс. `closeForm()` закрывает их. Автоочистка через `dismissPendingErrors` перед каждым action

#### readTable — подробнее

Каждая строка — объект `{ columnName: value }`. Специальные поля для иерархии и дерева:

- `_kind: 'group'` — группа в иерархическом списке
- `_tree: 'expanded'|'collapsed'` — состояние узла дерева
- `_level: N` — уровень вложенности
- `_selected: true` — строка выделена (подсвечена). Используйте с `clickElement({ modifier: 'ctrl'|'shift' })` для проверки мультиселекции
- На объекте результата: `hierarchical: true`, `viewMode: 'tree'`

#### clickElement — клик по ячейке SpreadsheetDocument

Для расшифровки отчётов первый аргумент `clickElement` принимает объект `{ row, column }` вместо текста. Координаты соответствуют выводу `readSpreadsheet()`:

```js
const report = await readSpreadsheet();
// report.data[0] = { 'К1': 'Материалы строительные', 'К6': '150 000' }

// По индексу строки данных + имя колонки
await clickElement({ row: 0, column: 'К6' }, { dblclick: true });

// По значению ячейки в строке (fuzzy match)
await clickElement({ row: { 'К1': 'Материалы' }, column: 'К6' }, { dblclick: true });

// Строка итогов
await clickElement({ row: 'totals', column: 'К6' }, { dblclick: true });
```

Текстовый поиск тоже работает — если элемент не найден в основном DOM, `clickElement` ищет в SpreadsheetDocument iframe'ах:

```js
await clickElement('150 000', { dblclick: true }); // найдёт ячейку в отчёте
```

### Действия

| Функция | Описание | Возвращает |
|---------|----------|------------|
| `clickElement(text, {dblclick?, modifier?})` | Клик по кнопке/ссылке/строке. `{dblclick: true}` для открытия, `{modifier: 'ctrl'\|'shift'}` для мультиселекции. Первый аргумент может быть `{row, column}` для клика по ячейке SpreadsheetDocument (см. выше) | form state или `{ submenu }` |
| `fillFields({name: value})` | Заполнить поля (текст, чекбокс, радио, ссылки, DCS-фильтры). Пустое значение (`''`/`null`) = очистка | `{ filled: [{field, ok, method}], form }` |
| `selectValue(field, search, opts?)` | Выбрать из справочника. search: текст, `{поле: значение}` или `''`/`null` для очистки. `{ type }` для составного типа | form state с `selected` |
| `fillTableRow(fields, {tab?, add?, row?})` | Заполнить строку. Значение: строка, `{ value, type }` для составного типа, `''`/`null` для очистки | form state |
| `deleteTableRow(row, {tab?})` | Удалить строку по индексу | form state |
| `closeForm({save?})` | Закрыть форму. `save: false` = "Нет", `save: true` = "Да". Возвращает `closed: true/false` | form state с `closed` |
| `filterList(text, {field?, exact?})` | Фильтр списка. Без field = все колонки, с field = расширенный поиск | form state |
| `unfilterList({field?})` | Снять фильтры (все или конкретный) | form state |

#### fillFields — типы полей

| Значение | Тип поля | Метод |
|----------|---------|--------|
| `'Андромеда Плюс'` | Ссылочное | clipboard paste + typeahead |
| `'5000'` | Текст | clipboard paste |
| `'true'` / `'да'` | Чекбокс | toggle |
| `'Оплата поставщику'` | Радио | fuzzy match по меткам |
| `'Склад бытовой техники'` (DCS) | Фильтр отчёта | авто-включение чекбокса + заполнение |
| `''` / `null` | Любое (кроме чекбокс/радио) | очистка через Shift+F4 |

### Утилиты

| Функция | Описание |
|---------|----------|
| `screenshot()` | Скриншот (PNG Buffer) |
| `wait(seconds)` | Пауза, возвращает form state |
| `getPage()` | Сырой Playwright Page для горячих клавиш и нестандартных операций |
| `startRecording(path, opts?)` | Начать запись видео (CDP screencast → ffmpeg → MP4) |
| `stopRecording()` | Остановить запись, вернуть `{ file, duration, size }` |
| `showCaption(text, opts?)` | Текстовая подпись поверх страницы (`speech` — текст озвучки) |
| `hideCaption()` | Убрать подпись |
| `showTitleSlide(text, opts?)` | Полноэкранный титульный слайд (`subtitle`, `background`, `speech`) |
| `hideTitleSlide()` | Убрать титульный слайд |
| `showImage(path, opts?)` | Полноэкранное изображение (`style`: blur/dark/light/full, `speech`) |
| `hideImage()` | Убрать изображение |
| `addNarration(videoPath, opts?)` | Озвучка видео по субтитрам (Edge TTS / ElevenLabs / OpenAI) |
| `getCaptions()` | Субтитры из текущей/последней записи |
| `isRecording()` | Идёт ли запись (boolean) |
| `setHighlight(on)` | Включить/выключить авто-выделение элементов при действиях |
| `isHighlightMode()` | Активен ли режим авто-выделения (boolean) |
| `highlight(text)` | Ручное выделение элемента (по имени, fuzzy match) |
| `unhighlight()` | Снять выделение |

## Клавиатурные сочетания

```js
const page = await getPage();
await page.keyboard.press('F8');  // пример: создать новый элемент в сфокусированном ссылочном поле
```

| Клавиша | Контекст | Действие |
|---------|----------|----------|
| `F8` | Ссылочное поле | Создать новый элемент |
| `Shift+F4` | Любое поле | Очистить значение (автоматизировано: `fillFields({ поле: '' })`) |
| `F4` | Ссылочное поле | Форма выбора |
| `Alt+F` | Список/таблица | Расширенный поиск |

## Типичные ошибки

Все функции бросают исключение при ошибке (не возвращают `{ error }`). Сценарий прерывается на проблемном шаге с информативным сообщением. В интерактиве — `try/catch` для обработки.

| Проблема | Решение |
|----------|---------|
| `no form found` — форма не открыта | Добавьте `await wait(2)` после навигации |
| `not found. Available: ...` — элемент не найден | Проверьте имя через `getFormState()`, используйте вариант из Available |
| `fillFields: N of M field(s) failed` | Текст ошибки содержит список проблемных полей и доступные варианты |
| Пустой `readSpreadsheet()` | Увеличьте `await wait(N)` перед чтением |

## Особенности

- **Headed mode** — 1С требует видимый браузер, headless не поддерживается
- **Время запуска** — первое подключение к 1С занимает 30-60 секунд (ожидание встроено)
- **Fuzzy matching** — все поиски: точное совпадение → начало строки → вхождение. Буквы ё и е считаются эквивалентными
- **Clipboard paste** — поля заполняются через Ctrl+V (корректно триггерит события 1С)
- **Неразрывные пробелы** — 1С использует `\u00a0`, внутри API нормализация автоматическая
- **Ошибки** — все функции бросают исключение при ошибке (сценарий прерывается), `try/catch` для обработки
- **Панель разделов** — `navigateSection()` работает при любом расположении панели (сбоку, сверху), но требует режим «Картинка и текст» или «Текст». Режим «Только картинки» не поддерживается — API не может прочитать имена разделов из иконок

## Связанные навыки

- [Запись видеоинструкций](web-test-recording-guide.md) — запись видео, субтитры, подсветка, TTS-озвучка
- [Веб-публикация](web-guide.md) — `/web-publish`, `/web-info`, `/web-stop`, `/web-unpublish`
- [Базы данных](db-guide.md) — `/db-load-xml`, `/db-update`, `/db-run`
- [Расширения](cfe-guide.md) — `/cfe-init`, `/cfe-borrow`, `/cfe-patch-method`


---

## Часть 3: Запись видеоинструкций

# Запись видеоинструкций

Навык `/web-test` умеет записывать видеоинструкции по работе в 1С: автоматические действия в браузере записываются в MP4 с субтитрами, подсветкой элементов и голосовой озвучкой. Результат — готовое обучающее видео.

```
сценарий → запись экрана → субтитры → подсветка → озвучка голосом → MP4
```

## Предусловия

Все пути и настройки хранятся в `.v8-project.json` — см. [справочник формата](v8-project-guide.md).

### ffmpeg (обязательно)

Выберите один из вариантов:

1. **В проект** (рекомендуется) — скачать essentials build с https://www.gyan.dev/ffmpeg/builds/, распаковать в `tools/ffmpeg/`. Код найдёт `tools/ffmpeg/bin/ffmpeg.exe` автоматически

2. **Глобально** — скачать, распаковать в любой каталог, добавить `bin/` в системный PATH

3. **Через конфиг** — указать путь в `.v8-project.json`:
   ```json
   { "ffmpegPath": "C:\\tools\\ffmpeg\\bin\\ffmpeg.exe" }
   ```

### node-edge-tts (для озвучки)

```bash
npm install --prefix tools/tts node-edge-tts
```

Бесплатный, без API-ключа. Если не установлен — запись видео работает, только озвучка недоступна.

### Конфигурация голоса в `.v8-project.json`

```json
{
  "ffmpegPath": "tools/ffmpeg/bin/ffmpeg.exe",
  "tts": {
    "provider": "edge",
    "voice": "ru-RU-DmitryNeural"
  }
}
```

## Быстрый старт

Минимальный сценарий — запись 3 шагов с озвучкой:

```js
// Начинаем запись
await startRecording('recordings/demo.mp4');

// Субтитры + действия
await showCaption('Переходим в раздел «Продажи»');
await wait(1.5);
await navigateSection('Продажи');

await showCaption('Открываем заказы клиентов');
await wait(1.5);
await openCommand('Заказы клиентов');

await showCaption('Создаём новый заказ');
await wait(1.5);
await clickElement('Создать');
await wait(2);

// Завершаем запись
await hideCaption();
const video = await stopRecording();
console.log(`Записано: ${video.duration.toFixed(1)}s`);

// Озвучка
const narrated = await addNarration(video.file, {
  ffmpegPath: 'tools/ffmpeg/bin/ffmpeg.exe',
  voice: 'ru-RU-DmitryNeural',
});
console.log(`Озвучено: ${narrated.file}`);
```

Результат: `recordings/demo-narrated.mp4` — видео с голосовым сопровождением.

## Сценарии использования

### Запись без озвучки

Простейший вариант — субтитры на экране, без голоса:

```
> Запиши видеоинструкцию: открой раздел Продажи, создай заказ клиента,
> заполни организацию и контрагента. Без озвучки
```

Claude запишет видео с субтитрами и подсветкой элементов.

### Запись с озвучкой

Полный pipeline — голос озвучивает каждый шаг:

```
> Запиши озвученную видеоинструкцию по созданию заказа клиента.
> Голос — Светлана
```

Claude запишет видео, затем наложит голосовую дорожку. Субтитры показываются на экране, параллельно звучит голос.

### Переозвучка другим голосом

Видео уже записано — хотите другой голос? Не нужно перезаписывать:

```
> Переозвучь recordings/demo.mp4 голосом Светланы
```

Claude вызовет `addNarration` с другим голосом. Тексты берутся из файла `.captions.json`, который сохраняется рядом с видео при записи.

### Редактирование субтитров

После записи рядом с видео появляется файл `video.captions.json`:

```json
{
  "videoTimestamps": true,
  "captions": [
    { "text": "Переходим в раздел «Продажи»", "speech": "Переходим в раздел Продажи", "time": 3160 },
    { "text": "Открываем заказы клиентов", "speech": "Открываем заказы клиентов", "time": 7040, "voice": "bqbHGIIO5oETYIqhWmfk" }
  ]
}
```

Можно отредактировать `speech` (текст озвучки) или добавить `voice` (голос для конкретной реплики) и переозвучить:

```
> Отредактируй субтитры в recordings/demo.captions.json — замени "Продажи" на
> "раздел Продажи", потом переозвучь
```

## Приёмы

### Титульный слайд

Полноэкранная заставка в начале видео. Поддерживает озвучку через `speech`:

```js
await startRecording('recordings/demo.mp4');
await showTitleSlide('Создание заказа клиента', {
  subtitle: '1С:Бухгалтерия в примерах',
  speech: 'Создание заказа клиента. Бухгалтерия в примерах.'
});
await wait(1);
await hideTitleSlide();
// ... далее контент
```

### Слайды из презентации

Показать изображение (скриншот слайда, схему и т.д.) как полноэкранный оверлей с озвучкой:

```js
await showImage('slides/overview.png', {
  speech: 'На этом слайде показана общая схема процесса'
});
await wait(1);
await hideImage();
```

Стили оформления (`style`):
- `'blur'` (по умолчанию) — размытый фон из картинки + тень. Лучший для презентаций
- `'dark'` — тёмный фон + тень
- `'light'` — белый фон + тень
- `'full'` — на весь экран без отступов

```js
await showImage('slides/diagram.png', { style: 'dark', speech: 'Диаграмма процесса' });
```

### Подсветка элементов

Полупрозрачная рамка на элементе, который сейчас используется. Два режима:

- **Авторежим** — `setHighlight(true)` перед началом действий. Каждая функция (`navigateSection`, `clickElement`, `fillFields` и т.д.) автоматически подсвечивает элемент перед действием
- **Ручная** — `highlight('Провести')` для произвольной подсветки конкретного элемента

```js
setHighlight(true);   // включить авто
// ... все действия подсвечиваются автоматически
setHighlight(false);  // выключить перед stopRecording
```

### Паузы и ритм

Ритм «субтитр → пауза → действие» даёт зрителю время прочитать, что произойдёт:

```js
await showCaption('Проводим документ');   // зритель читает
await wait(1.5);                           // пауза 1.5 сек
await clickElement('Провести');            // действие
```

Пауза после действия нужна только когда загружается следующая форма:

```js
await clickElement('Создать');
await wait(2);   // форма загружается
```

### Два голоса (подкаст / диалог)

Параметр `voice` в `showCaption` задаёт голос для конкретной реплики. `addNarration` автоматически использует его вместо глобального:

```js
const MALE   = 'bqbHGIIO5oETYIqhWmfk'; // Alexander
const FEMALE = '0ArNnoIAWKlT4WweaVMY'; // Elena Gromova

// speechRate: 85 — ElevenLabs медленнее Edge TTS, нужен запас
await startRecording('podcast.mp4', { speechRate: 85 });

await showImage('slides/slide-01.png', { style: 'full', speech: false });
await showCaption('', { speech: 'Привет! Сегодня поговорим...', voice: MALE });
await wait(0.8);
await showCaption('', { speech: 'А я буду задавать вопросы...', voice: FEMALE });
await wait(0.8);

const video = await stopRecording();
const result = await addNarration(video.file, {
  provider: 'elevenlabs',
  apiKey: 'sk_...',
  // глобальный voice не нужен — каждый caption несёт свой
});
```

Приёмы:
- `showCaption('', { speech, voice })` — пустой текст (без субтитра на экране), но speech записывается для озвучки
- `showImage` со `speech: false` — слайд без озвучки, реплики идут через `showCaption`
- `speechRate: 85` — для ElevenLabs увеличиваем множитель (по умолчанию 70мс/символ), чтобы фразы не наезжали друг на друга

### Разделение текста и озвучки

Параметр `speech` в `showCaption` позволяет показывать одно, а озвучивать другое:

```js
// Субтитр технический, озвучка человечная
await showCaption('Дт 60.02 — Кт 51', {
  speech: 'Дебет шестьдесят ноль два — кредит пятьдесят один'
});

// Показать субтитр, но НЕ озвучивать
await showCaption('Технические детали', { speech: false });
```

Это полезно для:
- **Бухгалтерских проводок** — на экране формула, голосом — словами
- **Технических данных** — показать, но не зачитывать
- **Информационных плашек** — немой субтитр на несколько секунд

## Доступные голоса и провайдеры

### Какой провайдер выбрать?

| Провайдер | Тембр | Произношение русского | Цена |
|-----------|-------|----------------------|------|
| **Edge TTS** | Синтетичнее | Корректные ударения, правильная артикуляция | Бесплатно |
| **ElevenLabs** | Живее, естественнее | Возможны ошибки в ударениях и артикуляции (напр. «докумЭнт», «крЕдит» вместо «кредИт») | Платно (starter+) |
| **OpenAI** | Зависит от голоса | Зависит от сервиса | Платно |

**Для русскоязычных видеоинструкций рекомендуется Edge TTS** — он бесплатный и даёт надёжное качество русской речи. Голоса DmitryNeural и SvetlanaNeural специально обучены для русского языка: правильно расставляют ударения, корректно артикулируют и делают паузы в нужных местах.

**ElevenLabs** даёт более живой, «человечный» тембр — голос звучит менее синтетически. Однако мультиязычная модель иногда ошибается в произношении русских слов (особенно профессиональная терминология). Если выбираете ElevenLabs для русского контента — берите **professional-голоса** с образовательным или деловым профилем (например, Olga Orlova, Artem), они дают лучший результат, чем англоязычные premade-голоса через мультиязычную модель. Управлять ударениями через API нельзя — фонемные теги (SSML) поддерживаются только для английских моделей.

### Edge TTS (бесплатный) — рекомендуется для русского

| Голос | Описание |
|-------|----------|
| `ru-RU-DmitryNeural` | Мужской, русский — спокойный, деловой |
| `ru-RU-SvetlanaNeural` | Женский, русский — чёткий, уверенный |

Полный список: `en-US-AriaNeural`, `en-US-GuyNeural`, `de-DE-ConradNeural` и другие. Edge TTS поддерживает десятки языков.

Конфигурация не нужна — Edge TTS используется по умолчанию. Для смены голоса:

```json
{
  "tts": {
    "voice": "ru-RU-SvetlanaNeural"
  }
}
```

### ElevenLabs (платный) — живой тембр

Модель `eleven_multilingual_v2` поддерживает русский. Тембр заметно живее, чем у Edge TTS, но возможны артикуляционные ошибки на русской терминологии.

Для русского контента выбирайте **professional-голоса** с образовательным/деловым профилем из Voice Library:

| Голос | ID | Профиль |
|-------|----|---------|
| Olga Orlova | `d60rsXo2p0OwikDR5bS7` | Clear and Engaging |
| Artem | `WTn2eCRCpoFAC50VD351` | Friendly & Professional |
| Denis | `0BcDz9UPwL3MpsnTeUlO` | Pleasant, Engaging and Friendly |
| Alexander | `bqbHGIIO5oETYIqhWmfk` | Pleasant, Warm and Natural |
| Elena Gromova | `0ArNnoIAWKlT4WweaVMY` | Podcasts & Conversation |
| Victor | `9fjVd0EYNNXHllJquVdT` | Moscow accent |

```json
{
  "tts": {
    "provider": "elevenlabs",
    "apiKey": "sk_...",
    "voice": "d60rsXo2p0OwikDR5bS7"
  }
}
```

`voice` — ID голоса (не имя). Professional-голоса добавляются в аккаунт через Voice Library в личном кабинете. Требуется платный тариф (starter и выше).

Особенности: лимит на параллельные запросы (2–3 одновременно), система автоматически ограничивает размер пакета.

### OpenAI-compatible (платный)

```json
{
  "tts": {
    "provider": "openai",
    "apiKey": "sk-...",
    "voice": "alloy"
  }
}
```

Голоса: `alloy`, `echo`, `fable`, `onyx`, `nova`, `shimmer`.

Поле `apiUrl` позволяет подключить любой OpenAI-совместимый сервис (например, локальный TTS-прокси).

## Полный пример

Типовая структура озвученного сценария:

```js
await startRecording('output.mp4');

// Титульный слайд с озвучкой
await showTitleSlide('Заголовок', {
  subtitle: 'Подзаголовок',
  speech: 'Заголовок. Подзаголовок.'
});
await wait(1);
await hideTitleSlide();

// Слайд из презентации (опционально)
await showImage('slides/overview.png', {
  speech: 'Описание слайда для озвучки'
});
await wait(1);
await hideImage();

setHighlight(true);

// ... шаги с showCaption + действия ...

await hideCaption();
setHighlight(false);
const video = await stopRecording();

const narrated = await addNarration(video.file, {
  ffmpegPath: 'tools/ffmpeg/bin/ffmpeg.exe',
  voice: 'ru-RU-SvetlanaNeural',
});
```

## Типичные проблемы

| Проблема | Решение |
|----------|---------|
| `ffmpeg not found` | Установите ffmpeg (см. Предусловия) |
| Файл записи 0 байт | Проверьте права на запись в выходной каталог |
| Видео дёргается | Добавьте `wait()` между шагами |
| `Already recording` | Вызовите `stopRecording()` перед новой записью |
| `No captions available` | Используйте `showCaption()` во время записи |
| TTS timeout | Проверьте интернет-соединение (Edge TTS требует сеть) |
| Озвучка обрезается | Увеличьте паузы `wait()` между субтитрами |
| Фразы наезжают друг на друга | Увеличьте `speechRate` в `startRecording` (85 для ElevenLabs) |

## Связанные навыки

- [Тестирование через веб-клиент](web-test-guide.md) — навигация, формы, таблицы, отчёты
- [Веб-публикация](web-guide.md) — `/web-publish`, `/web-info`, `/web-stop`
