# П1 «Специфікація і еталон» Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ввести в BuhBud (`bas_industrialbud`) справочник многоуровневой спецификации домов `А_СпецификацииДомов` со всеми ТЧ (включая закупочные — только метаданные), 6 перечислений, реквизит `СС.ЭталоннаяСпецификация` с серверной командой «Зробити еталоном для СС», метаданные снимка нормы в РК, проверки §4.6, разузлование и весь инструментарий бухгалтера для эталона (§5–6, Р22–Р26, Р31, Р32, Р35) — этапы 0, 1, 2 спецификации, каждый с загрузкой в базу и приёмкой.

**Architecture:** Источник истины для загрузки — основной каталог `C:/Configuration_downloads/BASERP25/_Rarzrabotki/BASEBuh` (он равен живой базе; дамп worktree отстал — заметки E1–E3). Правки ведутся в дампе worktree на синхронизированных копиях, затем `to_main.py` копирует их в основной каталог с проверкой хешей (Rule #4) и только оттуда идёт Partial-загрузка скриптом `load_stage.ps1` (гейты → `db-load-xml` → разбор лога → `/CheckModules` → `db-update -Dynamic "+"`). Логика — в общем модуле `А_СпецификацииДомовСервер` (Сервер + ВнешнееСоединение), модули объекта и форм тонкие; общий модуль собирается из частей `spec_domov/bsl/*.bsl`, а тексты запросов вставляются из `spec_domov/zaprosy/*.sql` — тех же файлов, которые до этого исполнил Python COM тест (Rule #-1). Приёмка — Python COM тесты в транзакции с откатом, данные (эталоны, закупочные) в базу не пишутся.

**Tech Stack:** 1С:Предприятие 8.3.20.1914 (BAS Бухгалтерія 2.1.34.2, режим совместимости 8.3.12 — без `Асинх`/`Ждать`), BSL, XML-выгрузка формата 2.13 (UTF-8 BOM + CRLF), навыки `meta-compile`, `meta-validate`, `form-validate`, `cf-validate`, `db-load-xml`, `db-update`, `db-dump-xml`; Python 3 + `pywin32` (`V83.COMConnector`), PowerShell 5.1.

---

## Контекст, который нужно прочитать до начала

| Документ | Зачем |
|---|---|
| `docs/superpowers/specs/2026-10-03-mnogourovnevaya-specifikaciya-domov-design.md` §2, §4, §5, §6, §9, §10.1 (этапы 0–2), §10.2 | требования; номера Р… и Т… ниже — оттуда |
| `_Rarzrabotki/notebook/knowledge_Бухгалтерия_РасчетКомплектаций/Разработка связей/research/2026-10-03_mnogourovnevaya_specifikaciya/M5_proekt_mnogourovnevoy_specifikacii.md` §2–§4 | схема, разузлование, ціни норми |
| `docs/superpowers/plans/_p1_notes/E1_obrazcy_metadannyh.md` | XML-образцы, грабли навыков |
| `docs/superpowers/plans/_p1_notes/E2_zagruzka_i_sverka.md` | рецепт загрузки, транзакции COM, этап 0 |
| `docs/superpowers/plans/_p1_notes/E3_tochki_vstraivaniya.md` | точки вставки в СС/РК, проверенные запросы, числа приёмки |
| Память: `config_xml_partial_load_deletes_objects`, `feedback_partial_load_only`, `partial_load_form_module`, `bsl_checkmodules_pered_zagruzkoy`, `bsl_write_lf_crlf_canon`, `uk_lang_sinonim_pustoy`, `feedback_vypolnit_reserved`, `bsl_module_object_locals_shadow_attributes`, `buhbud_zapis_bez_registracii_obmena`, `com_date_shift_3_hours`, `feedback_kod_srvr_localhost` | грабли, которые план уже учитывает |

## Решения и толкования этого плана (зафиксировать в отчёте этапа 0)

| № | Решение | Почему |
|---|---|---|
| Д1 | **Направление Rule #4 при отставшем worktree.** Затрагиваемые существующие файлы сначала копируются основной каталог → worktree (Task 2, с манифестом хешей), правятся в worktree, затем `to_main.py` копирует их в основной каталог, только если основной не менялся с момента синхронизации. Грузится всегда основной каталог. `git checkout --` в основном каталоге запрещён. | E1 §0, E2 §3, E3 §0: worktree-дамп удалил бы 2 объекта и 3 ТЧ РК |
| Д2 | **Права (решение пользователя 04.10.2026: «как сейчас у СС — все 24»):** вариант (б) — `ЗапускТонкогоКлиента` получает полный набор прав справочника, как у СС (правят все 24 пользователя без админ-прав). Дополнительно, как у всех `А_*`: `ДобавлениеИзменениеДанныхБухгалтерии` — полный набор, `ЧтениеДанныхБухгалтерии` — чтение. | пользователь 04.10.2026; E1 §6.1 |
| Д3 | **«паков»** (32 карточки) в синонимы Р23 не входит. | E3 §4.2 |
| Т9 | **«Деталь» без карточки при записи** допускается только в статусе «Чернетка» (рабочее состояние после «Заповнити з СС», 95 строк на 30 м без карточки); в «Діє»/«Закрита» — только с «Не постачається» и причиной. «Не постачається» без примечания не записывается ни в каком статусе. | буквальное §4.6 сделало бы невозможной запись черновика эталона (§5 шаг 1, 12–13 ч работы) |
| Т10 | **Зелёное правило перепривязки «единственная сирота»:** в этапе СС одна сирота и одна свободная позиция, количество совпало → зелёное предложение. | приёмка этапа 2 ждёт зелёное для «Копмлпект → Комплект», а [M5] §3.10 даёт жёлтое (E3 §8 п. 4) |
| Т11 | Копия узла из другой спецификации для узла «одна позиция на строке-узле» (Септик) ищет позицию сначала по ключу, затем в том же этапе по тексту с расстоянием правки ≤ 2. | 30 м поз. 170 «Копмлпект септика» ↔ 15 м поз. 168 «Комплект септика» — ключи разные, а Р32 требует переиспользовать Септик |

## Как исполнять этот план

- **Пути.** `WT` = `C:/Configuration_downloads/BASERP25/.claude/worktrees/budget-calc-oct-dec-2026-20c998`; `MAIN` = `C:/Configuration_downloads/BASERP25`. Все файлы плана — в `WT`; дамп BuhBud в нём — `WT/_Rarzrabotki/BASEBuh`, основной — `MAIN/_Rarzrabotki/BASEBuh`. Скрипты `spec_domov` запускать из `WT`: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/<скрипт>.py`. Временные файлы — `C:/Users/SUPPOR~1/AppData/Local/Temp/claude/plan_p1/`.
- **Запись файлов дампа и BSL** — только через `write_text` из `common.py` (UTF-8 с BOM, CRLF). Инструмент Write пишет LF без BOM: им создаются только файлы `spec_domov/*.py|*.json|*.sql|*.bsl|*.txt|*.ps1` (исходники), а в дамп они попадают через генераторы.
- **TDD для метаданных и BSL.** Тест падает до загрузки («объект/функция не найдены через COM»). Шаг «PASS» у задач, которые меняют общий модуль или формы, выполняется в задаче загрузки своей партии (1A, 1B, 2A, 2B, 2C, 2D): так грузится 6 раз, а не 20. Внутри задачи шаг «PASS» — статическая проверка `check_static.py`, она ловит баланс блоков, зарезервированные имена и псевдонимы до загрузки.
- **Транзакции.** Каждый негативный кейс приёмки — своя транзакция (`with tx(B)`); `ЗафиксироватьТранзакцию`/`CommitTransaction` и `.Удалить()` в тестах запрещены (гейт в `check_static.py`).
- **Коммиты** — только перечисленные файлы, никогда `git add -A` (в `WT` лежат чужие `Documents/А_ПредварительныйБюджет*`). `ConfigDumpInfo.xml` не коммитить. В конце каждого сообщения: `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- **Если `db-update` требует монопольный режим или его блокирует классификатор прав** — не выгонять пользователей: остановиться, сообщить пользователю, что загрузка выполнена и ждёт «Обновить конфигурацию БД» (F7), и продолжить после его подтверждения.

---

## Карта файлов

### Дамп BuhBud (`_Rarzrabotki/BASEBuh/…`, правка в `WT`, копия в `MAIN`)

| Файл | Создаётся / меняется | Ответственность | Задача |
|---|---|---|---|
| `_Rarzrabotki/BASEBuh/Enums/А_ВидыСпецификацийДомов.xml` | создаётся | Будинок / Вузол / Закупівельна | 12 |
| `_Rarzrabotki/BASEBuh/Enums/А_СтатусыСпецификацийДомов.xml` | создаётся | Чернетка / Діє / Закрита | 12 |
| `_Rarzrabotki/BASEBuh/Enums/А_ВидыСтрокСпецификацииДомов.xml` | создаётся | Деталь / Вузол / Група за загальною назвою | 12 |
| `_Rarzrabotki/BASEBuh/Enums/А_ИсточникиКарточкиСпецификации.xml` | создаётся | откуда карточка строки | 12 |
| `_Rarzrabotki/BASEBuh/Enums/А_ГруппыЗакупочнойСпецификации.xml` | создаётся | группы закупочной (логика — П2) | 12 |
| `_Rarzrabotki/BASEBuh/Enums/А_ИсточникиЦеныЗакупки.xml` | создаётся | источник цены закупочной (логика — П2) | 12 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml` | создаётся, в 26/41/42 дополняется регистрация форм | справочник: шапка §4.2, ТЧ `Состав`, `Дома`, `СоставОснования`, `ДомаФакта`, `ЗакупочныйСостав`, `ФактПоДомам` | 13 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl` | создаётся | тонкий: `ПередЗаписью`, `ПриКопировании` → общий модуль | 25 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl` | создаётся (этап 1), перегенерируется (этап 2) | форма элемента: «Склад», «Будинки», «Дерево», «Розгорнутий склад», «Порівняння», «Закупівля», «Факт по будинках» | 26, 41 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаСписка.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl` | создаётся (1), перегенерируется (2) | список: колонка «Еталон для СС», «Перевірити цикли», «Перепривʼязати позиції СС…» | 26, 42 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаГруппы.xml` + `/Ext/Form.xml` | создаётся | папка | 26 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаПодбораКарточки.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl` | создаётся | подбор карточки с контекстом (Р35) | 42 |
| `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаПерепривязки.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl` | создаётся | «Перепривʼязати позиції СС» | 42 |
| `_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер.xml` + `/Ext/Module.bsl` | создаётся (1B), дополняется (2A–2C) | вся логика П1 | 20–24, 30–39 |
| `_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости.xml` | меняется | реквизит `ЭталоннаяСпецификация` после `ЗапретитьРедактированиеСС` | 14 |
| `_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml` | меняется | поле `ЭталоннаяСпецификация` (только просмотр) в `Группа4` | 27 |
| `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml` | меняется | `СпецификацияДома`, `ДатаНормы`, ТЧ `НормаДома`, `ЭкземплярыУзлов`, колонки `ТабличнаяЧастьОстатков.ПозицияСС/НеИзНормы`, синоним `Спецификация` | 14 |
| `_Rarzrabotki/BASEBuh/Configuration.xml` | меняется | регистрация 1 справочника, 6 перечислений, 1 общего модуля | 15, 27 |
| `_Rarzrabotki/BASEBuh/Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml` | меняется | полный набор прав справочника | 15 |
| `_Rarzrabotki/BASEBuh/Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml` | меняется | чтение справочника | 15 |
| `_Rarzrabotki/BASEBuh/Roles/ЗапускТонкогоКлиента/Ext/Rights.xml` | меняется | полный набор прав справочника (Д2, вариант б) | 15 |
| `_Rarzrabotki/BASEBuh/Subsystems/Склад/Subsystems/А_СписанняБудинків.xml` + `/Ext/CommandInterface.xml` | меняется | справочник в навигации рядом с СС и РК | 15 |

### Python и исходники (`_Rarzrabotki/Python/test/spec_domov/…`)

| Файл | Ответственность | Задача |
|---|---|---|
| `common.py` | пути, COM-подключение, `tx()`, `q()`, `err()`, `call()`, `messages()`, `md5_nocr()`, `read_text()`/`write_text()` (BOM+CRLF), `zapros()` | 1 |
| `test_common.py` | самопроверка `common.py` | 1 |
| `sync_main_to_wt.py`, `sync_manifest.json` | синхронизация основной → worktree с манифестом хешей | 2 |
| `sessions_gate.py`, `preflight_config.py`, `preflight_roles.py`, `check_log.py`, `test_gates.py` | гейты загрузки | 3 |
| `to_main.py`, `dump_check.py`, `test_to_main.py` | Rule #4 worktree → основной, обратная выгрузка и канон | 4 |
| `load_stage.ps1` | рецепт загрузки партии (ASCII, без BOM) | 5 |
| `vendor/f31_buh_moves.py`, `vendor/f33_buh_receipts.py`, `baseline/z6_raw_20261003_234112.json.gz` | независимость этапа 0 от untracked-файлов и Temp | 6 |
| `e0_snimok.py`, `e0_kontrol.py`, `out/e0_*.json` | этап 0 | 7, 8 |
| `meta/e1_enums.json`, `meta/e1_catalog.json`, `meta/e1_module.json` | DSL `meta-compile` | 12, 13, 25 |
| `postpatch.py`, `test_postpatch.py` | нормализация после `meta-compile`: `cfg:`, `/>`, uk-синонимы, `CheckUnique`, `ChoiceParameters` | 11 |
| `gen_meta.py` | `meta-compile` в scratch + постпатч + копия в дамп worktree | 12 |
| `patch_existing.py` | хирургические вставки в СС, РК, форму СС | 14, 27 |
| `register.py` | `Configuration.xml`, роли, подсистема | 15, 27 |
| `check_static.py` | статический гейт XML и BSL | 16 |
| `build_bsl.py` | сборка `Module.bsl` и `ObjectModule.bsl` из частей + вставка запросов | 20 |
| `gen_forms.py` | генерация пяти форм | 26, 41, 42 |
| `fixtures.py` | создание тестовых элементов внутри транзакции | 19 |
| `bsl/srv_10_obschee.bsl` … `bsl/srv_B0_kopiya.bsl`, `bsl/obj_module.bsl`, `bsl/forma_*.bsl` | исходники BSL | 20–25, 30–42 |
| `zaprosy/*.sql` | тексты запросов: один источник для BSL и Python-теста | 18, 29 |
| `load_lists/e1a.txt` … `load_lists/e2d.txt` | списки Partial-загрузки | 17, 28, 33, 37, 40, 43 |
| `test_e1_files.py`, `test_e1_metadata.py`, `test_e1_zaprosy.py`, `test_e1_funkcii.py`, `test_e1_priemka.py` | этап 1 | 9, 10, 18, 19, 20 |
| `test_e2_zaprosy.py`, `test_e2a_edinicy.py`, `test_e2a_zapolnit.py`, `test_e2b_uzly.py`, `test_e2b_nulevye.py`, `test_e2b_ceny.py`, `test_e2c_perepriv.py`, `test_e2c_kopiya.py`, `test_e2d_formy.py`, `test_e2_priemka.py` | этап 2 | 29–44 |
| `out/*.json` | малые сводки приёмки (коммитятся); большие снимки — в Temp | 7, 8, 17, 28, 44 |

### Списки Partial-загрузки (пути от корня дампа)

```
load_lists/e1a.txt  (метаданные, Task 17)
Configuration.xml
Enums/А_ВидыСпецификацийДомов.xml
Enums/А_СтатусыСпецификацийДомов.xml
Enums/А_ВидыСтрокСпецификацииДомов.xml
Enums/А_ИсточникиКарточкиСпецификации.xml
Enums/А_ГруппыЗакупочнойСпецификации.xml
Enums/А_ИсточникиЦеныЗакупки.xml
Catalogs/А_СпецификацииДомов.xml
Catalogs/СтруктураСебестоимости.xml
Documents/РасчетКомплектаций.xml
Roles/ДобавлениеИзменениеДанныхБухгалтерии.xml
Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml
Roles/ЧтениеДанныхБухгалтерии.xml
Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml
Roles/ЗапускТонкогоКлиента.xml
Roles/ЗапускТонкогоКлиента/Ext/Rights.xml
Subsystems/Склад/Subsystems/А_СписанняБудинків.xml
Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml

load_lists/e1b.txt  (код и формы этапа 1, Task 28)
Configuration.xml
CommonModules/А_СпецификацииДомовСервер.xml
CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl
Catalogs/А_СпецификацииДомов.xml
Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl
Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаСписка.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаГруппы.xml
Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента.xml

load_lists/e2a.txt, e2b.txt, e2c.txt  (Tasks 33, 37, 40)
CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl

load_lists/e2d.txt  (формы этапа 2, Task 43)
Catalogs/А_СпецификацииДомов.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаСписка.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаПодбораКарточки.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаПерепривязки.xml
```

Правила списков (E2 §4.2): новый объект верхнего уровня — с `Configuration.xml`; форма — только корневым `Forms/<Форма>.xml`; модули — прямым путём; роль — `Roles/X.xml` + `Roles/X/Ext/Rights.xml`; дочерняя подсистема — без `Configuration.xml`.

---

## Оглавление задач

| № | Задача | Этап | Загрузка |
|---:|---|---|---|
| 1 | Каркас `spec_domov` и `common.py` | 0 | — |
| 2 | Синхронизация основного каталога → worktree (Д1) | 0 | — |
| 3 | Гейты загрузки | 0 | — |
| 4 | `to_main.py` (Rule #4) и `dump_check.py` | 0 | — |
| 5 | `load_stage.ps1` — рецепт загрузки партии | 0 | — |
| 6 | Vendor-копии запросов Z6 и архив эталонного снимка | 0 | — |
| 7 | Этап 0 — снимок контрольных сумм (`e0_snimok.py`) | 0 | — |
| 8 | Этап 0 — сверка конфигурации, обмена и ролей (`e0_kontrol.py`), фиксация решений | 0 | — |
| 9 | COM-тест метаданных этапа 1 (падает до загрузки) | 1 | 1A |
| 10 | Статический тест файлов этапа 1A | 1 | 1A |
| 11 | Постпатч после `meta-compile` (`postpatch.py`) | 1 | 1A |
| 12 | Шесть перечислений спецификации (`gen_meta.py`) | 1 | 1A |
| 13 | Справочник `А_СпецификацииДомов` (метаданные §4.1–4.5) | 1 | 1A |
| 14 | Новые реквизиты СС и РК (`patch_existing.py`, часть метаданных) | 1 | 1A |
| 15 | Регистрация — `Configuration.xml`, роли (Д2), подсистема | 1 | 1A |
| 16 | Статический гейт `check_static.py` | 1 | 1A |
| 17 | Загрузка 1A (метаданные) и приёмка метаданных | 1 | 1A |
| 18 | Rule #-1 — запросы к новому справочнику (Python COM до BSL) | 1 | 1B |
| 19 | Фикстуры и приёмочный тест этапа 1 (падает до 1B) | 1 | 1B |
| 20 | Сборщик модулей `build_bsl.py`, общее и единицы | 1 | 1B |
| 21 | Дерево и разузлование (§6, Т6) | 1 | 1B |
| 22 | Входимость, уровень в домах и поиск циклов (§4.6, Т6) | 1 | 1B |
| 23 | Проверки при записи и при «Діє», `ПередЗаписью`, `ПриКопировании` (§4.6, Т9) | 1 | 1B |
| 24 | «Зробити еталоном для СС» без ПередЗаписью СС и без регистрации обмена (Р3, Б1) | 1 | 1B |
| 25 | Общий модуль (метаданные) и модуль объекта справочника | 1 | 1B |
| 26 | Генератор форм `gen_forms.py` и формы этапа 1 | 1 | 1B |
| 27 | Поле эталона на форме СС и регистрация общего модуля | 1 | 1B |
| 28 | Загрузка 1B и приёмка этапа 1 | 1 | 1B |
| 29 | Rule #-1 — запросы этапа 2 (Python COM до BSL) | 2 | 2A |
| 30 | Единицы этапа 2 — кэш, единица по тексту, подсказка k, размер из названия (Р23, Р25) | 2 | 2A |
| 31 | Факт домов, цены последней закупки, склад проекта, кандидаты, документы комплектации (Р31) | 2 | 2A |
| 32 | «Заповнити з СС», «Позначити зелені перевіреними», «Взяти з оригіналу Excel», подбор карточки (Р25, Р35) | 2 | 2A |
| 33 | Загрузка 2A и прогон тестов партии | 2 | 2A |
| 34 | Узлы строителя — создать, объединить, расформировать, из документа комплектации, узел в узле (§9, Р31) | 2 | 2B |
| 35 | «Нульові позиції з фактом» и печать для экономиста (Р22) | 2 | 2B |
| 36 | «Розрахувати ціни норми» (И4) и «Порівняти з СС» (Р24, Р26, Р27) | 2 | 2B |
| 37 | Загрузка 2B и прогон тестов партии | 2 | 2B |
| 38 | «Перепривʼязати позиції СС» ([M5] §3.10, Т10) | 2 | 2C |
| 39 | «Скопіювати картки й дерево вузлів з іншої специфікації» (Р32, Т11) и приёмка 15 м | 2 | 2C |
| 40 | Загрузка 2C и прогон тестов партии | 2 | 2C |
| 41 | Форма элемента этапа 2 — команды бухгалтера | 2 | 2D |
| 42 | Формы подбора карточки, перепривязки и список этапа 2; генерация форм | 2 | 2D |
| 43 | Загрузка 2D (формы) и проверка форм в базе | 2 | 2D |
| 44 | Приёмка этапа 2 целиком (§10.1) | 2 | — |
| 45 | Завершение — сверка основного каталога, память, git | 0–2 | — |

---

## Фаза 0 — инструменты и этап 0 (только чтение)


### Task 1: Каркас `spec_domov` и `common.py`

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/common.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/test_common.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/out/.gitkeep`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_common.py`:

```python
# -*- coding: utf-8 -*-
"""Самопроверка common.py (только чтение базы): транзакция всегда откатывается, ошибка 1С читается, BOM+CRLF, md5 без CR."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import call, connect, md5_nocr, messages, q, read_text, tx, write_text, zapros_path  # noqa: E402


def test_tx_rollback(B):
    with tx(B):
        assert B.TransactionActive()
        B.BeginTransaction()                     # как будто BSL открыл вложенную и упал
    assert not B.TransactionActive()


def test_err_and_q(B):
    rows = q(B, "ВЫБРАТЬ 1 КАК Н")
    assert rows == [{"Н": 1}], rows
    z = B.NewObject("Запрос")
    z.Text = "ВЫБРАТЬ ОШИБКА ИЗ"
    res, text = call(z.Execute)
    assert res is None and text, text
    assert isinstance(messages(B), list)



def test_files():
    d = tempfile.mkdtemp()
    p1, p2 = os.path.join(d, "a.bsl"), os.path.join(d, "b.bsl")
    write_text(p1, "Рядок 1\nРядок 2\n")
    raw = open(p1, "rb").read()
    assert raw.startswith(b"\xef\xbb\xbf"), "нет BOM"
    assert raw.count(b"\r\n") == 2 and raw.replace(b"\r\n", b"").count(b"\n") == 0, "не CRLF"
    open(p2, "wb").write("Рядок 1\nРядок 2\n".encode("utf-8"))
    assert md5_nocr(p1) == md5_nocr(p2)
    text, bom, crlf = read_text(p1)
    assert text == "Рядок 1\nРядок 2\n" and bom and crlf
    assert zapros_path("x").endswith(os.path.join("zaprosy", "x.sql"))


if __name__ == "__main__":
    B = connect()
    test_tx_rollback(B)
    test_err_and_q(B)
    test_files()
    print("OK test_common")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `cd C:/Configuration_downloads/BASERP25/.claude/worktrees/budget-calc-oct-dec-2026-20c998 && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_common.py`
Expected: FAIL — `ModuleNotFoundError: No module named 'common'`.

- [ ] **Step 3: Реализовать `common.py`**

`_Rarzrabotki/Python/test/spec_domov/common.py`:

```python
# -*- coding: utf-8 -*-
"""Общие функции П1 «Специфікація і еталон»: пути, COM bas_industrialbud, транзакция с откатом, файлы дампа.

База — ТОЛЬКО Srvr="localhost" (память feedback_kod_srvr_localhost). Данные в базу не пишутся:
всё, что тест создаёт, живёт внутри tx(B) и откатывается.
"""
import contextlib
import hashlib
import io
import os
import sys

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

WT = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\budget-calc-oct-dec-2026-20c998"
MAIN = r"C:\Configuration_downloads\BASERP25"
BB_WT = os.path.join(WT, r"_Rarzrabotki\BASEBuh")
BB_MAIN = os.path.join(MAIN, r"_Rarzrabotki\BASEBuh")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
TMP = r"C:\Users\SUPPOR~1\AppData\Local\Temp\claude\plan_p1"
SKILLS = os.path.join(MAIN, r".claude\skills")
V8_BIN = r"C:\Program Files\1cv8\8.3.20.1914\bin"
CONN = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'


def connect():
    import win32com.client
    return win32com.client.Dispatch("V83.COMConnector").Connect(CONN)


def err(e):
    """Текст ошибки 1С из pywintypes.com_error; AttributeError — опечатка в имени модуля/метода."""
    info = getattr(e, "excepinfo", None)
    if info and info[2]:
        return str(info[2])
    return f"{type(e).__name__}: {e}"


def call(fn, *args):
    """(результат, None) или (None, текст ошибки) — для негативных тестов."""
    import pywintypes
    try:
        return fn(*args), None
    except (pywintypes.com_error, AttributeError) as e:
        return None, err(e)


@contextlib.contextmanager
def tx(B):
    """Всё внутри — в транзакции, которая ВСЕГДА откатывается (вложенные тоже)."""
    B.BeginTransaction()
    try:
        yield
    finally:
        while B.TransactionActive():
            B.RollbackTransaction()
    assert not B.TransactionActive()


def arr(B, items):
    a = B.NewObject("Массив")
    for x in items:
        a.Добавить(x)
    return a


def q(B, text, **params):
    """Результат запроса списком словарей; значения — как отдаёт COM (ссылки — COM-объекты)."""
    z = B.NewObject("Запрос")
    z.Text = text
    for k, v in params.items():
        z.SetParameter(k, v)
    vt = z.Execute().Выгрузить()
    cols = [vt.Колонки.Получить(i).Имя for i in range(vt.Колонки.Количество())]
    return [{c: vt.Получить(i).Получить(j) for j, c in enumerate(cols)} for i in range(vt.Количество())]


def vt_rows(vt, cols):
    """ТаблицаЗначений (COM) → список словарей по заданным колонкам."""
    return [{c: getattr(vt.Получить(i), c) for c in cols} for i in range(vt.Количество())]


def sv(B, v):
    """Строковое представление значения 1С (ссылки, перечисления — в языке сеанса, uk)."""
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return B.String(v)


def same(B, a, b):
    """Равенство ссылок/перечислений через представление и УИД (COM-обёртки сравнивать == нельзя)."""
    if a is None or b is None:
        return a is b
    try:
        return B.String(a.УникальныйИдентификатор()) == B.String(b.УникальныйИдентификатор())
    except Exception:
        return B.String(a) == B.String(b)


def messages(B):
    m = B.ПолучитьСообщенияПользователю(True)
    return [str(m.Получить(i).Текст) for i in range(m.Количество())]


def changes_rows(B, ss_ref):
    """Строки регистрации СС к обмену: (узел, номер сообщения). Число строк при повторной регистрации не меняется — сравнивать сами строки."""
    rows = q(B, """ВЫБРАТЬ
	ПРЕДСТАВЛЕНИЕ(Изм.Узел) КАК Узел,
	ЕСТЬNULL(Изм.НомерСообщения, -1) КАК НомерСообщения
ИЗ
	Справочник.СтруктураСебестоимости.Изменения КАК Изм
ГДЕ
	Изм.Ссылка = &СС""", СС=ss_ref)
    return sorted((str(r["Узел"]), int(r["НомерСообщения"])) for r in rows)


def count_rows(B, table, where=""):
    return int(q(B, f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Н ИЗ {table} КАК Т {where}")[0]["Н"])


def read_text(path):
    """(текст с \\n, был_BOM, был_CRLF)."""
    raw = open(path, "rb").read()
    return raw.decode("utf-8-sig").replace("\r\n", "\n"), raw.startswith(b"\xef\xbb\xbf"), b"\r\n" in raw


def write_text(path, text, bom=True, crlf=True):
    """Канон дампа BASEBuh: UTF-8 с BOM и CRLF (память bsl_write_lf_crlf_canon)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    text = text.replace("\r\n", "\n")
    if crlf:
        text = text.replace("\n", "\r\n")
    with open(path, "wb") as f:
        f.write((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))


def md5_nocr(path):
    raw = open(path, "rb").read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    return hashlib.md5(raw.replace(b"\r\n", b"\n")).hexdigest()


def zapros_path(name):
    return os.path.join(HERE, "zaprosy", name + ".sql")


def zapros(name):
    """Текст запроса — тот же файл, который build_bsl.py вставляет в BSL (Rule #-1: один источник)."""
    return io.open(zapros_path(name), encoding="utf-8-sig").read().replace("\r\n", "\n").strip("\n")


def vidy_subkonto(B):
    """[Номенклатура, Склады] — склад не всегда третье субконто (E3 §3.2)."""
    pvh = B.ПланыВидовХарактеристик.ВидыСубконтоХозрасчетные
    return arr(B, [pvh.Номенклатура, pvh.Склады])
```

`_Rarzrabotki/Python/test/spec_domov/out/.gitkeep` — пустой файл.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_common.py`
Expected: `OK test_common`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/common.py _Rarzrabotki/Python/test/spec_domov/test_common.py _Rarzrabotki/Python/test/spec_domov/out/.gitkeep \
  docs/superpowers/plans/_p1_notes/E1_obrazcy_metadannyh.md docs/superpowers/plans/_p1_notes/E2_zagruzka_i_sverka.md \
  docs/superpowers/plans/_p1_notes/E3_tochki_vstraivaniya.md
git commit -m "test(spec_domov): каркас П1 — common.py (COM, транзакция с откатом, BOM+CRLF); заметки E1–E3 к плану

Заметки закоммичены сразу: Task 29 копирует из E3 тексты проверенных запросов по номерам строк.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Синхронизация основного каталога → worktree (Д1)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/sync_main_to_wt.py`
- Create (генерируется): `_Rarzrabotki/Python/test/spec_domov/sync_manifest.json`
- Modify (копии из основного): 9 файлов `_Rarzrabotki/BASEBuh/…` из списка `SYNC` ниже

- [ ] **Step 1: Написать падающий тест** — режим `--check` того же скрипта сверяет worktree с основным каталогом.

`_Rarzrabotki/Python/test/spec_domov/sync_main_to_wt.py`:

```python
# -*- coding: utf-8 -*-
"""Д1: существующие файлы дампа, которые правит П1, копируются ОСНОВНОЙ → WORKTREE; хеши (без CR) пишутся в манифест.

Основной каталог = живая база (E2 §2), worktree-дамп отстал (E1 §0, E3 §0). to_main.py потом копирует правки
обратно, только если основной файл не менялся с момента синхронизации (хеш манифеста).
Запуск:
    python sync_main_to_wt.py --check             # 0 — worktree = основной по всем файлам SYNC
    python sync_main_to_wt.py                     # копировать все файлы SYNC и обновить манифест
    python sync_main_to_wt.py --file <rel>        # пересинхронизировать один файл (после отказа to_main.py)
"""
import argparse
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE, md5_nocr  # noqa: E402

MANIFEST = os.path.join(HERE, "sync_manifest.json")
SYNC = [
    "Configuration.xml",
    "Catalogs/СтруктураСебестоимости.xml",
    "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента.xml",
    "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml",
    "Documents/РасчетКомплектаций.xml",
    "Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml",
    "Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml",
    "Roles/ЗапускТонкогоКлиента/Ext/Rights.xml",
    "Subsystems/Склад/Subsystems/А_СписанняБудинків.xml",
    "Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml",
]


def load_manifest():
    if os.path.exists(MANIFEST):
        return json.load(open(MANIFEST, encoding="utf-8"))
    return {}


def save_manifest(m):
    json.dump(m, open(MANIFEST, "w", encoding="utf-8"), ensure_ascii=False, indent=1, sort_keys=True)


def copy_one(rel, manifest):
    src, dst = os.path.join(BB_MAIN, rel), os.path.join(BB_WT, rel)
    assert os.path.exists(src), f"нет в основном каталоге: {rel}"
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)                          # байт в байт: BOM и CRLF сохраняются
    h = md5_nocr(src)
    assert md5_nocr(dst) == h, rel
    manifest[rel] = h
    print(f"  синхронизирован {rel} md5={h}")


def check():
    bad = []
    for rel in SYNC:
        dst = os.path.join(BB_WT, rel)
        if not os.path.exists(dst) or md5_nocr(dst) != md5_nocr(os.path.join(BB_MAIN, rel)):
            bad.append(rel)
    for rel in bad:
        print("  РАСХОДИТСЯ:", rel)
    print(f"файлов {len(SYNC)}, расхождений {len(bad)}")
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--file")
    a = ap.parse_args()
    if a.check:
        sys.exit(check())
    m = load_manifest()
    for rel in ([a.file] if a.file else SYNC):
        copy_one(rel, m)
    save_manifest(m)
    sys.exit(check() if not a.file else 0)
```

- [ ] **Step 2: Запустить проверку — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/sync_main_to_wt.py --check`
Expected: FAIL (код 1) — `РАСХОДИТСЯ: Configuration.xml`, `Documents/РасчетКомплектаций.xml`, обе роли, подсистема `А_СписанняБудинків` (в worktree её нет).

- [ ] **Step 3: Выполнить синхронизацию**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/sync_main_to_wt.py`
Expected: 9 строк `синхронизирован …`, затем `файлов 9, расхождений 0`, код 0. Создан `sync_manifest.json`.

- [ ] **Step 4: Запустить проверку — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/sync_main_to_wt.py --check`
Expected: `файлов 9, расхождений 0`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/sync_main_to_wt.py _Rarzrabotki/Python/test/spec_domov/sync_manifest.json \
  "_Rarzrabotki/BASEBuh/Configuration.xml" "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости.xml" \
  "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента.xml" \
  "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml" \
  "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml" \
  "_Rarzrabotki/BASEBuh/Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml" \
  "_Rarzrabotki/BASEBuh/Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml" "_Rarzrabotki/BASEBuh/Roles/ЗапускТонкогоКлиента/Ext/Rights.xml" \
  "_Rarzrabotki/BASEBuh/Roles/ЗапускТонкогоКлиента/Ext/Rights.xml" \
  "_Rarzrabotki/BASEBuh/Subsystems/Склад/Subsystems/А_СписанняБудинків.xml" \
  "_Rarzrabotki/BASEBuh/Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml"
git commit -m "sync(BASEBuh): объекты П1 и Configuration.xml = живая bas_industrialbud на 04.10.2026

Копия из основного каталога (он равен базе, 3856 объектов). Configuration.xml регистрирует объекты,
файлы которых в основном каталоге не закоммичены (AIRDoc_*, А_ФинансовыйОтчетПроизводства и др.):
worktree-дамп не грузится, загрузка идёт только из основного каталога.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Гейты загрузки

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/sessions_gate.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/preflight_config.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/preflight_roles.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/check_log.py`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_gates.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_gates.py`:

```python
# -*- coding: utf-8 -*-
"""Гейты должны ЛОВИТЬ аварию: Configuration.xml без объекта базы → код 1; лог с «Объект удален» → код 1."""
import os
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, HERE, read_text, write_text  # noqa: E402

PY = [sys.executable, "-X", "utf8"]


def run(*args):
    return subprocess.run(PY + list(args), capture_output=True, text=True, encoding="utf-8").returncode


def test_config_gate_catches_lost_object():
    d = tempfile.mkdtemp()
    cfg = os.path.join(d, "Configuration.xml")
    text, _, _ = read_text(os.path.join(BB_MAIN, "Configuration.xml"))
    line = "\t\t\t<Catalog>СтруктураСебестоимости</Catalog>\n"
    assert text.count(line) == 1
    write_text(cfg, text.replace(line, ""))
    assert run(os.path.join(HERE, "preflight_config.py"), cfg) == 1, "гейт не поймал потерю СС"
    assert run(os.path.join(HERE, "preflight_config.py")) == 0, "основной каталог должен проходить"


def test_log_gate():
    d = tempfile.mkdtemp()
    bad, good = os.path.join(d, "bad.log"), os.path.join(d, "good.log")
    open(bad, "w", encoding="utf-8-sig").write("Загрузка...\nОбъект удален: Справочник.Х\n")
    open(good, "w", encoding="utf-8-sig").write("Загрузка конфигурации успешно завершена\n")
    assert run(os.path.join(HERE, "check_log.py"), "load", bad) == 1
    assert run(os.path.join(HERE, "check_log.py"), "load", good) == 0
    open(good, "w", encoding="utf-8-sig").write("Синтаксических ошибок не обнаружено!\n")
    assert run(os.path.join(HERE, "check_log.py"), "checkmodules", good, "0") == 0
    assert run(os.path.join(HERE, "check_log.py"), "checkmodules", good, "101") == 1


def test_sessions_and_roles():
    assert run(os.path.join(HERE, "sessions_gate.py")) == 0
    assert run(os.path.join(HERE, "preflight_roles.py")) == 0


if __name__ == "__main__":
    test_config_gate_catches_lost_object()
    test_log_gate()
    test_sessions_and_roles()
    print("OK test_gates")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_gates.py`
Expected: FAIL — `AssertionError: гейт не поймал потерю СС` (скрипта `preflight_config.py` ещё нет, `run` возвращает 2).

- [ ] **Step 3: Реализовать четыре гейта**

`_Rarzrabotki/Python/test/spec_domov/sessions_gate.py`:

```python
# -*- coding: utf-8 -*-
"""Гейт 1: в базе нет Конфигуратора (он блокирует пакетную загрузку, а со старой копией — откатывает её при F7)."""
import collections
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect  # noqa: E402

B = connect()
arr = B.ПолучитьСеансыИнформационнойБазы()
apps = collections.Counter()
designer = []
for i in range(arr.Количество()):
    s = arr.Получить(i)
    app = B.String(s.ИмяПриложения)
    apps[app] += 1
    if app.lower() in ("designer", "конфигуратор", "конфігуратор"):
        designer.append((B.String(s.Пользователь), B.String(s.ИмяКомпьютера), B.String(s.НачалоСеанса)))
print("сеансов:", arr.Количество(), dict(apps))
if designer:
    print("ОТКРЫТ КОНФИГУРАТОР — НЕ ГРУЗИТЬ:", designer)
    sys.exit(1)
print("Конфигуратора нет.")
```

`_Rarzrabotki/Python/test/spec_domov/preflight_config.py`:

```python
# -*- coding: utf-8 -*-
"""Гейт 2: Partial-загрузка с Configuration.xml, где нет объекта базы, УДАЛЯЕТ этот объект (память
config_xml_partial_load_deletes_objects). Правило: база ⊆ файл и файл − база ⊆ EXPECTED_NEW. Все коллекции Metadata.
Запуск: python preflight_config.py [путь к Configuration.xml]  (по умолчанию — основной каталог). Код 0 — грузить можно."""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, connect  # noqa: E402

EXPECTED_NEW = {
    ("Catalog", "А_СпецификацииДомов"),
    ("Enum", "А_ВидыСпецификацийДомов"),
    ("Enum", "А_СтатусыСпецификацийДомов"),
    ("Enum", "А_ВидыСтрокСпецификацииДомов"),
    ("Enum", "А_ИсточникиКарточкиСпецификации"),
    ("Enum", "А_ГруппыЗакупочнойСпецификации"),
    ("Enum", "А_ИсточникиЦеныЗакупки"),
    ("CommonModule", "А_СпецификацииДомовСервер"),
}
EXPECTED_REMOVED = set()

COLL = {"Catalog": "Catalogs", "Document": "Documents", "Enum": "Enums", "InformationRegister": "InformationRegisters",
        "AccumulationRegister": "AccumulationRegisters", "AccountingRegister": "AccountingRegisters",
        "DataProcessor": "DataProcessors", "Report": "Reports", "CommonModule": "CommonModules", "Role": "Roles",
        "EventSubscription": "EventSubscriptions", "CommandGroup": "CommandGroups", "Subsystem": "Subsystems",
        "Constant": "Constants", "CommonForm": "CommonForms", "DocumentJournal": "DocumentJournals",
        "ScheduledJob": "ScheduledJobs", "ExchangePlan": "ExchangePlans", "ChartOfAccounts": "ChartsOfAccounts",
        "ChartOfCharacteristicTypes": "ChartsOfCharacteristicTypes", "BusinessProcess": "BusinessProcesses",
        "Task": "Tasks", "CommonTemplate": "CommonTemplates", "CommonPicture": "CommonPictures",
        "CommonAttribute": "CommonAttributes", "SessionParameter": "SessionParameters",
        "FunctionalOption": "FunctionalOptions", "DefinedType": "DefinedTypes", "FilterCriterion": "FilterCriteria",
        "SettingsStorage": "SettingsStorages", "CommonCommand": "CommonCommands", "StyleItem": "StyleItems",
        "Language": "Languages", "WebService": "WebServices", "HTTPService": "HTTPServices", "XDTOPackage": "XDTOPackages",
        "FunctionalOptionsParameter": "FunctionalOptionsParameters", "CalculationRegister": "CalculationRegisters",
        "ChartOfCalculationTypes": "ChartsOfCalculationTypes", "Sequence": "Sequences", "Style": "Styles",
        "WSReference": "WSReferences", "ExternalDataSource": "ExternalDataSources",
        "DocumentNumerator": "DocumentNumerators"}


def file_objects(path):
    text = io.open(path, encoding="utf-8-sig").read()
    body = text[text.find("<ChildObjects>"):text.find("</ChildObjects>")]
    return {(k, n) for k, n in re.findall(r"<(\w+)>([^<]+)</\1>", body) if k in COLL}


def db_objects(B):
    md, found = B.Metadata, set()
    for kind, prop in COLL.items():
        try:
            coll = getattr(md, prop)
        except Exception:
            continue
        for i in range(coll.Count()):
            found.add((kind, coll.Get(i).Name))
    return found


def main(path):
    B = connect()
    db = db_objects(B)
    kinds = {k for k, _ in db}
    cfg = {x for x in file_objects(path) if x[0] in kinds or x in EXPECTED_NEW}
    lost = sorted(db - cfg - EXPECTED_REMOVED)
    extra = sorted(cfg - db)
    unexpected = sorted(set(extra) - EXPECTED_NEW)
    print(f"Configuration.xml: {path}")
    print(f"объектов в базе {len(db)}, в файле {len(cfg)}")
    print(f"новых П1 в файле, но не в базе: {[n for _, n in extra if (_, n) in EXPECTED_NEW]}")
    print(f"новых П1 уже в базе: {[n for k, n in sorted(EXPECTED_NEW & db)]}")
    fails = [f"УДАЛИТ из базы: {k}.{n}" for k, n in lost] + [f"посторонний в файле: {k}.{n}" for k, n in unexpected]
    print(f"ПРОВАЛОВ: {len(fails)}")
    for f in fails[:50]:
        print("   ", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(BB_MAIN, "Configuration.xml")))
```

`_Rarzrabotki/Python/test/spec_domov/preflight_roles.py`:

```python
# -*- coding: utf-8 -*-
"""Гейт 3: права роли в Rights.xml ОСНОВНОГО каталога = ПравоДоступа в базе (ловит чужие незакоммиченные правки:
наша загрузка роли иначе протащит их в базу или отнимет права). Новые объекты П1 до загрузки пропускаются."""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, connect  # noqa: E402

ROLES = ["ДобавлениеИзменениеДанныхБухгалтерии", "ЧтениеДанныхБухгалтерии", "ЗапускТонкогоКлиента"]
NEW_OBJECTS = {"Catalog.А_СпецификацииДомов"}
COLLECTION = {"Catalog": "Catalogs", "Document": "Documents", "InformationRegister": "InformationRegisters",
              "AccumulationRegister": "AccumulationRegisters", "AccountingRegister": "AccountingRegisters",
              "DataProcessor": "DataProcessors", "Report": "Reports", "Constant": "Constants",
              "DocumentJournal": "DocumentJournals", "ChartOfAccounts": "ChartsOfAccounts",
              "ChartOfCharacteristicTypes": "ChartsOfCharacteristicTypes", "ExchangePlan": "ExchangePlans",
              "BusinessProcess": "BusinessProcesses", "Task": "Tasks", "CalculationRegister": "CalculationRegisters",
              "ChartOfCalculationTypes": "ChartsOfCalculationTypes"}
CHECKED = {"Read", "Insert", "Update", "Delete", "View", "Edit", "InteractiveInsert", "InteractiveDelete",
           "InteractiveSetDeletionMark", "InteractiveClearDeletionMark", "InteractiveDeleteMarked", "InputByString",
           "Posting", "UndoPosting", "InteractivePosting", "InteractiveUndoPosting", "Use", "TotalsControl"}
RE_OBJ = re.compile(r"<object>\s*<name>([^<]+)</name>(.*?)</object>", re.S)
RE_RIGHT = re.compile(r"<right>\s*<name>(\w+)</name>\s*<value>(\w+)</value>")


def file_rights(role):
    text = io.open(os.path.join(BB_MAIN, "Roles", role, "Ext", "Rights.xml"), encoding="utf-8-sig").read()
    out = {}
    for name, body in RE_OBJ.findall(text):
        if name.count(".") == 1 and name.split(".")[0] in COLLECTION:
            out[name] = {r: v == "true" for r, v in RE_RIGHT.findall(body) if r in CHECKED}
    return out


def main():
    B = connect()
    md = B.Metadata
    fails = []
    for role in ROLES:
        md_role = md.Roles.Find(role)
        checked = skipped = 0
        for name, flags in file_rights(role).items():
            kind, obj = name.split(".")
            md_obj = getattr(md, COLLECTION[kind]).Find(obj)
            if md_obj is None:
                if name in NEW_OBJECTS:
                    skipped += 1
                    continue
                fails.append(f"{role}: {name} есть в Rights.xml, объекта нет в базе")
                continue
            for right, want in flags.items():
                try:
                    have = bool(B.ПравоДоступа(right, md_obj, md_role))
                except Exception:
                    continue
                checked += 1
                if have != want:
                    fails.append(f"{role}: {name}.{right} — файл {want}, база {have}")
        print(f"{role}: сверено прав {checked}, новых П1 пропущено {skipped}")
    print(f"РАСХОЖДЕНИЙ: {len(fails)}")
    for f in fails[:40]:
        print("   ", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
```

`_Rarzrabotki/Python/test/spec_domov/check_log.py`:

```python
# -*- coding: utf-8 -*-
"""Разбор логов загрузки (Tee-Object PowerShell 5.1 пишет UTF-16) и /CheckModules.
Запуск: check_log.py load <лог>  |  check_log.py checkmodules <лог> <код выхода 1cv8>. Код 0 — всё чисто."""
import sys

BAD_LOAD = ["Объект удален", "Объект удалён", "Неизвестный объект", "не входит в состав", "Неверное свойство",
            "нельзя добавлять", "Error loading configuration", "Ошибка загрузки"]
OK_CHECK = ["Синтаксических ошибок не обнаружено", "Синтаксичних помилок не виявлено"]


def read_log(path):
    raw = open(path, "rb").read()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1251")


def main(argv):
    kind, path = argv[1], argv[2]
    text = read_log(path)
    if kind == "load":
        hits = [line for line in text.splitlines() if any(b in line for b in BAD_LOAD)]
        for h in hits:
            print("АВАРИЯ ЗАГРУЗКИ:", h)
        return 1 if hits else 0
    if kind == "checkmodules":
        code = int(argv[3])
        ok = code == 0 and any(o in text for o in OK_CHECK)
        if not ok:
            print(text[-5000:])
            print(f"CheckModules: код {code} — модули с ошибками (см. {{Объект(строка,колонка)}} и <<?>> выше)")
        return 0 if ok else 1
    raise SystemExit(f"неизвестный вид лога: {kind}")


if __name__ == "__main__":
    sys.exit(main(sys.argv))
```

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_gates.py`
Expected: `OK test_gates`. Ручная проверка: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/preflight_config.py` → `ПРОВАЛОВ: 0`, «новых П1 в файле» — пустой список.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/sessions_gate.py _Rarzrabotki/Python/test/spec_domov/preflight_config.py \
  _Rarzrabotki/Python/test/spec_domov/preflight_roles.py _Rarzrabotki/Python/test/spec_domov/check_log.py \
  _Rarzrabotki/Python/test/spec_domov/test_gates.py
git commit -m "test(spec_domov): гейты загрузки П1 — сеансы, Configuration.xml, роли, разбор логов

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: `to_main.py` (Rule #4) и `dump_check.py`

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/to_main.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/dump_check.py`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_to_main.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_to_main.py`:

```python
# -*- coding: utf-8 -*-
"""to_main.py: копирует worktree → основной только при неизменном основном (хеш манифеста); --verify сверяет список."""
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE  # noqa: E402

PY = [sys.executable, "-X", "utf8", os.path.join(HERE, "to_main.py")]


def test_verify_synced_list():
    lst = os.path.join(tempfile.mkdtemp(), "l.txt")
    open(lst, "w", encoding="utf-8").write("Configuration.xml\nDocuments/РасчетКомплектаций.xml\nRoles/ЧтениеДанныхБухгалтерии.xml\n")
    r = subprocess.run(PY + ["--verify", lst], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stdout + r.stderr


def test_refuse_unknown_existing():
    """Номенклатура.xml есть в обоих каталогах и РАЗЛИЧАЕТСЯ (worktree отстал), в манифесте её нет → отказ."""
    r = subprocess.run(PY + ["Catalogs/Номенклатура.xml"], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 1 and "не в манифесте" in r.stdout, r.stdout


def test_skip_unchanged():
    """Файл списка, равный основному (корень роли), не копируется и не даёт отказа — иначе список e1a падал бы на Roles/X.xml."""
    r = subprocess.run(PY + ["Roles/ЧтениеДанныхБухгалтерии.xml"], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0 and "без изменений" in r.stdout, r.stdout


def test_expand_forms():
    """Корень формы СС тянет синхронизированный Ext/Form.xml, но не модуль формы, которого нет в манифесте."""
    import to_main
    root = "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента.xml"
    out = to_main.expand([root])
    assert out == [root, "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml"], out


if __name__ == "__main__":
    test_verify_synced_list()
    test_refuse_unknown_existing()
    test_skip_unchanged()
    test_expand_forms()
    print("OK test_to_main")

```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_to_main.py`
Expected: FAIL — `AssertionError` (скрипта `to_main.py` нет, код 2).

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/to_main.py`:

```python
# -*- coding: utf-8 -*-
"""Rule #4 при отставшем worktree (Д1): копия WORKTREE → ОСНОВНОЙ каталог дампа BuhBud.

- файл, который уже равен основному (md5 без CR), не копируется: «без изменений» (корни ролей в списке e1a не правятся);
- файл из манифеста синхронизации: копируется, только если основной не менялся (md5 = манифест); иначе отказ —
  выполнить sync_main_to_wt.py --file <rel>, повторить свой патч-скрипт (идемпотентен) и to_main.py;
- новый файл П1 (нет в манифесте): копируется, если в основном его нет или он уже равен копии worktree;
- после копии md5 сверяется, манифест обновляется.
Запуск:
    python to_main.py <rel> [<rel> …]          # копировать
    python to_main.py --list <файл списка>     # копировать все файлы списка, которые есть в worktree
    python to_main.py --verify <файл списка>   # 0 — каждый файл списка в worktree совпадает с основным
Корневой файл формы в списке тянет её Ext/Form.xml и Ext/Form/Module.bsl (загрузка берёт форму целиком) — только те,
что П1 правит: новые формы П1 и синхронизированные (есть в манифесте). Устаревший модуль формы из worktree не копируется.
"""
import argparse
import json
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE, md5_nocr  # noqa: E402

MANIFEST = os.path.join(HERE, "sync_manifest.json")


def read_list(path):
    return [ln.strip() for ln in open(path, encoding="utf-8-sig") if ln.strip() and not ln.startswith("#")]


def is_new_p1(rel):
    return any(m in rel for m in ("А_СпецификацииДомов", "А_ВидыСпецификацийДомов", "А_СтатусыСпецификацийДомов",
                                  "А_ВидыСтрокСпецификацииДомов", "А_ИсточникиКарточкиСпецификации",
                                  "А_ГруппыЗакупочнойСпецификации", "А_ИсточникиЦеныЗакупки"))


def expand(rels):
    manifest = json.load(open(MANIFEST, encoding="utf-8")) if os.path.exists(MANIFEST) else {}
    out = []
    for rel in rels:
        out.append(rel)
        m = re.match(r"(.*/Forms/[^/]+)\.xml$", rel)
        if m:
            for part in (m.group(1) + "/Ext/Form.xml", m.group(1) + "/Ext/Form/Module.bsl"):
                if os.path.exists(os.path.join(BB_WT, part)) and (part in manifest or is_new_p1(part)):
                    out.append(part)
    return out


def copy(rels):
    manifest = json.load(open(MANIFEST, encoding="utf-8"))
    fails = []
    for rel in rels:
        wt, mn = os.path.join(BB_WT, rel), os.path.join(BB_MAIN, rel)
        if not os.path.exists(wt):
            fails.append(f"{rel}: нет в worktree")
            continue
        if os.path.exists(mn) and md5_nocr(mn) == md5_nocr(wt):
            print(f"  без изменений {rel}")
            continue
        if rel in manifest:
            if os.path.exists(mn) and md5_nocr(mn) != manifest[rel] and md5_nocr(mn) != md5_nocr(wt):
                fails.append(f"{rel}: основной каталог изменён после синхронизации — sync_main_to_wt.py --file {rel}, "
                             f"повторить патч, затем to_main.py")
                continue
        elif not is_new_p1(rel):
            fails.append(f"{rel}: не в манифесте и не новый объект П1 — сначала синхронизировать")
            continue
        elif os.path.exists(mn) and md5_nocr(mn) != md5_nocr(wt):
            print(f"  ПЕРЕЗАПИСЬ нового файла П1 в основном каталоге: {rel}")
        os.makedirs(os.path.dirname(mn), exist_ok=True)
        shutil.copyfile(wt, mn)
        assert md5_nocr(mn) == md5_nocr(wt), rel
        manifest[rel] = md5_nocr(mn)
        print(f"  скопирован {rel}")
    json.dump(manifest, open(MANIFEST, "w", encoding="utf-8"), ensure_ascii=False, indent=1, sort_keys=True)
    for f in fails:
        print("ОТКАЗ:", f)
    return 1 if fails else 0


def verify(rels):
    bad = []
    for rel in rels:
        wt, mn = os.path.join(BB_WT, rel), os.path.join(BB_MAIN, rel)
        if not os.path.exists(mn):
            bad.append(f"{rel}: нет в основном каталоге")
        elif os.path.exists(wt) and md5_nocr(wt) != md5_nocr(mn):
            bad.append(f"{rel}: worktree ≠ основной (to_main.py не выполнен?)")
    for b in bad:
        print("НЕ СОВПАДАЕТ:", b)
    print(f"файлов {len(rels)}, расхождений {len(bad)}")
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("rels", nargs="*")
    ap.add_argument("--list")
    ap.add_argument("--verify")
    a = ap.parse_args()
    if a.verify:
        sys.exit(verify(expand(read_list(a.verify))))
    rels = a.rels + ([r for r in expand(read_list(a.list)) if os.path.exists(os.path.join(BB_WT, r))] if a.list else [])
    sys.exit(copy(rels))
```

`_Rarzrabotki/Python/test/spec_domov/dump_check.py`:

```python
# -*- coding: utf-8 -*-
"""Обратная выгрузка объектов из базы во Temp и сверка с ОСНОВНЫМ каталогом (E2 §4.6).

Частичная выгрузка не включает формы и макеты — их перечислять (`Справочник.X.Форма.Y`).
Сравнение XML — по канонической форме C14N (порядок xmlns и атрибутов, пробелы между тегами, `<a/>` = `<a></a>` не важны).
--canon: файлы, отличающиеся только форматированием, заменяются выгруженной версией в основном каталоге И в worktree
(канон — версия платформы, память bsl_write_lf_crlf_canon). Различие по содержимому:
- у СУЩЕСТВУЮЩИХ объектов (СС, РК, форма СС, план обмена) — код 1 (загрузка не дошла или кеш Конфигуратора);
- у НОВЫХ файлов П1 (сгенерированы meta-compile/gen_forms) платформа сериализует загруженное по-своему (умолчания свойств):
  печатается начало различия, при --canon выгруженная версия становится каноном («ПРИНЯТО КАНОНОМ»), код 0. Что загрузка
  дошла, доказывают COM-тесты этапа (test_e1_metadata, test_e1_priemka, test_e2d_formy), а не побайтная сверка.
Запуск: python dump_check.py --tag e1a --objects "Справочник.А_СпецификацииДомов,Перечисление.А_ВидыСпецификацийДомов" [--canon]
"""
import argparse
import datetime
import difflib
import json
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE, SKILLS, TMP, V8_BIN, md5_nocr, read_text  # noqa: E402

MANIFEST = os.path.join(HERE, "sync_manifest.json")
NEW_MARK = ("А_СпецификацииДомов", "А_ВидыСпецификацийДомов", "А_СтатусыСпецификацийДомов", "А_ВидыСтрокСпецификацииДомов",
            "А_ИсточникиКарточкиСпецификации", "А_ГруппыЗакупочнойСпецификации", "А_ИсточникиЦеныЗакупки")


def normalized(path):
    text, _, _ = read_text(path)
    if path.endswith(".xml"):
        try:
            return ET.canonicalize(text, strip_text=True)
        except Exception:                                   # noqa: BLE001 — не XML: сравнить без пробелов между тегами
            return re.sub(r">\s+<", "><", text).strip()
    return "\n".join(line.rstrip() for line in text.split("\n")).strip()


def first_diff(a, b):
    """Первые строки различия канонических форм (для XML — после разбиения по тегам)."""
    sa, sb = (x.replace("><", ">\n<").split("\n") for x in (a, b))
    return "\n".join(list(difflib.unified_diff(sa, sb, "основной", "база", n=1, lineterm=""))[:12])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", required=True)
    ap.add_argument("--objects", required=True)
    ap.add_argument("--canon", action="store_true")
    a = ap.parse_args()
    out = os.path.join(TMP, f"dump_{a.tag}_{datetime.datetime.now():%Y%m%d_%H%M%S}")
    os.makedirs(out, exist_ok=True)
    cmd = ["powershell.exe", "-NoProfile", "-File", os.path.join(SKILLS, r"db-dump-xml\scripts\db-dump-xml.ps1"),
           "-V8Path", V8_BIN, "-InfoBaseServer", "localhost", "-InfoBaseRef", "bas_industrialbud",
           "-UserName", "cfo", "-Password", "2442", "-ConfigDir", out, "-Mode", "Partial", "-Objects", a.objects]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout[-2000:])
    if r.returncode != 0:
        print(r.stderr[-2000:])
        return 1
    manifest = json.load(open(MANIFEST, encoding="utf-8")) if os.path.exists(MANIFEST) else {}
    same = fmt_only = 0
    content, new_content = [], []

    def adopt(src, rel, main_path, label):
        for dst in (main_path, os.path.join(BB_WT, rel)):
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copyfile(src, dst)
        manifest[rel] = md5_nocr(main_path)
        print(f"  {label}:", rel)

    for root, _, files in os.walk(out):
        for fn in files:
            if fn in ("Configuration.xml", "ConfigDumpInfo.xml"):
                continue
            src = os.path.join(root, fn)
            rel = os.path.relpath(src, out).replace("\\", "/")
            main_path = os.path.join(BB_MAIN, rel)
            if os.path.exists(main_path) and md5_nocr(main_path) == md5_nocr(src):
                same += 1
            elif os.path.exists(main_path) and normalized(main_path) == normalized(src):
                fmt_only += 1
                if a.canon:
                    adopt(src, rel, main_path, "канон")
            elif any(m in rel for m in NEW_MARK):
                new_content.append(rel)
                if os.path.exists(main_path):
                    print(f"   НОВЫЙ ФАЙЛ П1, платформа сериализовала иначе: {rel}\n{first_diff(normalized(main_path), normalized(src))}")
                if a.canon:
                    adopt(src, rel, main_path, "ПРИНЯТО КАНОНОМ (новый файл П1)")
            else:
                content.append(rel)
    if a.canon:
        json.dump(manifest, open(MANIFEST, "w", encoding="utf-8"), ensure_ascii=False, indent=1, sort_keys=True)
    print(f"выгрузка {out}: совпало {same}, только форматирование {fmt_only}, новых файлов П1 иначе {len(new_content)}, "
          f"различие по содержимому {len(content)}")
    for c in content:
        print("   СОДЕРЖИМОЕ РАЗЛИЧАЕТСЯ:", c)
    return 1 if content or ((fmt_only or new_content) and not a.canon) else 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_to_main.py`
Expected: `OK test_to_main`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/dump_check.py --tag t4 --objects "Справочник.СтруктураСебестоимости,Справочник.СтруктураСебестоимости.Форма.ФормаЭлемента,Документ.РасчетКомплектаций"`
Expected: `различие по содержимому 0`, код 0 (E2 §2: основной каталог = база).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/to_main.py _Rarzrabotki/Python/test/spec_domov/dump_check.py _Rarzrabotki/Python/test/spec_domov/test_to_main.py
git commit -m "test(spec_domov): Rule #4 worktree→основной с проверкой манифеста, обратная выгрузка и канон

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: `load_stage.ps1` — рецепт загрузки партии

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/load_stage.ps1` (только ASCII — тогда BOM не нужен, память `feedback_cyrillic_ps1_bom`)

- [ ] **Step 1: Написать падающий тест** — сухой прогон должен пройти все гейты и остановиться до загрузки.

Run (тест-команда): `powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt -Tag dry -DryRun`

- [ ] **Step 2: Запустить — FAIL**

Expected: FAIL — `The argument '...load_stage.ps1' ... does not exist`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/load_stage.ps1`:

```powershell
# Load one batch of P1 into bas_industrialbud: gates -> db-load-xml Partial -> log -> /CheckModules -> db-update -Dynamic + -> gate.
# Loads ONLY from the main dump directory (D1). ASCII-only file: Cyrillic checks live in Python (check_log.py).
param(
    [Parameter(Mandatory=$true)][string]$List,
    [Parameter(Mandatory=$true)][string]$Tag,
    [switch]$WithRoles,
    [switch]$DryRun
)
$V8   = "C:\Program Files\1cv8\8.3.20.1914\bin"
$SK   = "C:\Configuration_downloads\BASERP25\.claude\skills"
$CFG  = "C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
$HERE = Split-Path -Parent $MyInvocation.MyCommand.Path
$LOGD = "C:\Users\SUPPOR~1\AppData\Local\Temp\claude\plan_p1\load_$Tag"
New-Item -ItemType Directory -Force $LOGD | Out-Null
$CONN = @("-V8Path", $V8, "-InfoBaseServer", "localhost", "-InfoBaseRef", "bas_industrialbud", "-UserName", "cfo", "-Password", "2442")
# Summary of a real load (committed): out\<Tag>_load_<yyyyMMdd_HHmm>.txt. Dry run writes nothing.
$SUMMARY = $null
if (-not $DryRun) {
    New-Item -ItemType Directory -Force (Join-Path $HERE "out") | Out-Null
    $SUMMARY = Join-Path (Join-Path $HERE "out") ("{0}_load_{1}.txt" -f $Tag, (Get-Date -Format "yyyyMMdd_HHmm"))
}

function Note([string]$Text) {
    Write-Host $Text
    if ($SUMMARY) { Add-Content -Path $SUMMARY -Value $Text -Encoding UTF8 }
}

function Invoke-Step([string]$Name, [scriptblock]$Body) {
    Note "=== $Name"
    & $Body
    $code = $LASTEXITCODE
    Note "    exit=$code"
    if ($code -ne 0) { Note "STOP at step '$Name' exit=$code. Logs: $LOGD"; exit 1 }
}

Note ("load {0}: list {1}, started {2}" -f $Tag, $List, (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))

Invoke-Step "sessions"        { python -X utf8 "$HERE\sessions_gate.py" }
Invoke-Step "main = worktree" { python -X utf8 "$HERE\to_main.py" --verify $List }
Invoke-Step "preflight config" { python -X utf8 "$HERE\preflight_config.py" }
if ($WithRoles) { Invoke-Step "preflight roles" { python -X utf8 "$HERE\preflight_roles.py" } }
if ($DryRun) { Write-Host "DRY RUN OK ($Tag): gates passed, nothing loaded"; exit 0 }

$files = (Get-Content $List -Encoding UTF8 | Where-Object { $_.Trim() -and -not $_.StartsWith("#") }) -join ","
Invoke-Step "db-load-xml Partial" {
    & powershell.exe -NoProfile -File "$SK\db-load-xml\scripts\db-load-xml.ps1" @CONN -ConfigDir $CFG -Mode Partial -Files $files -StrictLog *>&1 | Tee-Object -FilePath "$LOGD\load.log"
}
Invoke-Step "load log" { python -X utf8 "$HERE\check_log.py" load "$LOGD\load.log" }
Invoke-Step "CheckModules" {
    $log = "$LOGD\checkmodules.log"
    $p = Start-Process -FilePath "$V8\1cv8.exe" -ArgumentList @("DESIGNER", "/S", "`"localhost/bas_industrialbud`"", "/N`"cfo`"", "/P`"2442`"", "/CheckModules", "-ThinClient", "-Server", "-ExternalConnection", "/Out", "`"$log`"", "/DisableStartupDialogs") -NoNewWindow -Wait -PassThru
    python -X utf8 "$HERE\check_log.py" checkmodules $log $p.ExitCode
}
Invoke-Step "db-update -Dynamic +" {
    & powershell.exe -NoProfile -File "$SK\db-update\scripts\db-update.ps1" @CONN -Dynamic "+" *>&1 | Tee-Object -FilePath "$LOGD\update.log"
}
Invoke-Step "preflight config after" { python -X utf8 "$HERE\preflight_config.py" }
Note "OK load $Tag. Logs: $LOGD"
exit 0
```

Создать также `_Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt` (нужен сухому прогону) — содержимое из раздела «Списки Partial-загрузки» выше (16 строк). Новых файлов П1 ещё нет: `to_main.py --verify` для них проверит наличие в основном каталоге — поэтому сухой прогон на этом шаге запускать со списком только из существующих объектов:

`_Rarzrabotki/Python/test/spec_domov/load_lists/dry.txt`:
```
Configuration.xml
Catalogs/СтруктураСебестоимости.xml
Documents/РасчетКомплектаций.xml
```

- [ ] **Step 4: Запустить — PASS**

Run: `powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/dry.txt -Tag dry -DryRun -WithRoles`
Expected: шаги `sessions`, `main = worktree`, `preflight config`, `preflight roles` — каждый код 0; последняя строка `DRY RUN OK (dry): gates passed, nothing loaded`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 _Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt _Rarzrabotki/Python/test/spec_domov/load_lists/dry.txt
git commit -m "test(spec_domov): load_stage.ps1 — гейты, Partial-загрузка, разбор лога, CheckModules, db-update

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Vendor-копии запросов Z6 и архив эталонного снимка

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/vendor/f31_buh_moves.py` (копия `MAIN/_Rarzrabotki/Python/test/ekonomiya_ss_irc/f31_buh_moves.py`)
- Create: `_Rarzrabotki/Python/test/spec_domov/vendor/f33_buh_receipts.py` (копия `…/ekonomiya_ss_irc/f33_buh_receipts.py`)
- Create: `_Rarzrabotki/Python/test/spec_domov/baseline/z6_raw_20261003_234112.json.gz`
- Create: `_Rarzrabotki/Python/test/spec_domov/make_baseline.py`

- [ ] **Step 1: Написать падающий тест** — `make_baseline.py --check` проверяет vendor и архив.

`_Rarzrabotki/Python/test/spec_domov/make_baseline.py`:

```python
# -*- coding: utf-8 -*-
"""Этап 0 не должен зависеть от untracked-файлов основного каталога и от Temp (E2 §5.1, §9 п. 6):
vendor — копии текстов запросов f31/f33 (z6_collect читает из них TEXT), baseline — gzip снимка Z6 03.10.2026 23:41:12.
Запуск: python make_baseline.py           # создать
        python make_baseline.py --check   # 0 — всё на месте и снимок тот самый"""
import gzip
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, MAIN  # noqa: E402

SRC_RES = os.path.join(MAIN, r"_Rarzrabotki\Python\test\ekonomiya_ss_irc")
SRC_RAW = r"C:\Users\SUPPOR~1\AppData\Local\Temp\claude\z6_tmp\calc\raw.json"
VENDOR = os.path.join(HERE, "vendor")
BASE = os.path.join(HERE, "baseline", "z6_raw_20261003_234112.json.gz")
FILES = ["f31_buh_moves.py", "f33_buh_receipts.py"]


def check():
    for f in FILES:
        p = os.path.join(VENDOR, f)
        assert os.path.exists(p), f"нет {p}"
        assert 'TEXT = """' in open(p, encoding="utf-8").read(), f"в {f} нет TEXT"
    assert os.path.exists(BASE), f"нет {BASE}"
    raw = json.loads(gzip.open(BASE, "rt", encoding="utf-8").read())
    assert raw["stamp_full"] == "03.10.2026 23:41:12", raw["stamp_full"]
    assert len(raw["moves"]) == 4800, len(raw["moves"])
    print("OK baseline:", raw["stamp_full"], "движений", len(raw["moves"]))
    return 0


def make():
    os.makedirs(VENDOR, exist_ok=True)
    for f in FILES:
        shutil.copyfile(os.path.join(SRC_RES, f), os.path.join(VENDOR, f))
    os.makedirs(os.path.dirname(BASE), exist_ok=True)
    with open(SRC_RAW, "rb") as fi, gzip.open(BASE, "wb", compresslevel=9) as fo:
        shutil.copyfileobj(fi, fo)
    return check()


if __name__ == "__main__":
    sys.exit(check() if "--check" in sys.argv else make())
```

- [ ] **Step 2: Запустить проверку — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/make_baseline.py --check`
Expected: FAIL — `AssertionError: нет …vendor\f31_buh_moves.py`.

- [ ] **Step 3: Создать архив**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/make_baseline.py`
Expected: `OK baseline: 03.10.2026 23:41:12 движений 4800`; файл `baseline/z6_raw_20261003_234112.json.gz` ≈ 0,5 МБ.
Если `SRC_RAW` уже удалён из Temp — остановиться и сообщить: этап 0 тогда сравнивает только с опубликованным `data/proekt_irc2026_zakupki.json` (без построчного diff движений).

- [ ] **Step 4: Проверка — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/make_baseline.py --check`
Expected: `OK baseline: …`, код 0.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/make_baseline.py _Rarzrabotki/Python/test/spec_domov/vendor/f31_buh_moves.py \
  _Rarzrabotki/Python/test/spec_domov/vendor/f33_buh_receipts.py _Rarzrabotki/Python/test/spec_domov/baseline/z6_raw_20261003_234112.json.gz
git commit -m "test(spec_domov): vendor-копии запросов Z6 и архив снимка 03.10.2026 23:41:12 для этапа 0

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Этап 0 — снимок контрольных сумм (`e0_snimok.py`)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/e0_snimok.py`
- Create (результат): `_Rarzrabotki/Python/test/spec_domov/out/e0_snimok_<YYYYMMDD_HHMM>.json`

- [ ] **Step 1: Тест** — сам скрипт и есть тест: код 0 только если тождество Т3 замкнуто, независимые числа совпали и отличия от Z6 объяснены документами.

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/e0_snimok.py`
Expected: FAIL — `No such file or directory`.

- [ ] **Step 3: Реализовать** (основа — проверенный пробный прогон `plan_p1/E2/e0_trial.py`, E2 §5.2)

`_Rarzrabotki/Python/test/spec_domov/e0_snimok.py`:

```python
# -*- coding: utf-8 -*-
"""Этап 0 (§10.1): снимок контрольных сумм проекта МД IRC 2026 — ТОЛЬКО ЧТЕНИЕ bas_industrialbud.

Скрипты Z6 переиспользуются без правки оригиналов: копии в RUN, пути заменены с assert count == 1.
Выход: out/e0_snimok_<stamp>.json (сводка, коммитится); RUN (Temp) — полный снимок raw.json и расчёт.
Код 0: тождество Т3 = 0,00; независимые числа q_t3 совпали; Δ против Z6 = 0 или каждое Δ объяснено регистраторами."""
import datetime
import gzip
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, OUT, TMP, WT  # noqa: E402

KB = os.path.join(WT, r"_Rarzrabotki\notebook\knowledge_Бухгалтерия_РасчетКомплектаций\Разработка связей")
Z6S = os.path.join(KB, r"research\2026-10-03_mnogourovnevaya_specifikaciya\Z6_scripts")
PROTO = os.path.join(KB, r"scripts\03_prototip_svyazi_30m")
PUBLISHED = os.path.join(KB, r"data\proekt_irc2026_zakupki.json")
BASE_GZ = os.path.join(HERE, "baseline", "z6_raw_20261003_234112.json.gz")
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RUN = os.path.join(TMP, "e0", STAMP)
CALC = os.path.join(RUN, "calc")
PY = [sys.executable, "-X", "utf8"]
KEYS = ["need_15_uah_net", "need_30_uah_net", "purchased_uah_net", "purchased_uah_vat", "other_in_uah_net",
        "other_out_uah_net", "written_off_uah_net", "written_off_15_uah_net", "written_off_30_uah_net", "lies_uah_net",
        "lies_houses_uah_net", "lies_project_uah_net", "lies_ostatki_now_uah_net", "to_buy_30_uah_last_vat",
        "to_buy_30_uah_ss_vat", "to_buy_30_uah_last_net", "to_buy_30_cards", "surplus_uah_net", "surplus_cards"]


def patched(src, dst, reps):
    s = open(src, encoding="utf-8").read()
    for old, new in reps:
        assert s.count(old) == 1, (os.path.basename(src), old[:70], s.count(old))
        s = s.replace(old, new)
    open(dst, "w", encoding="utf-8").write(s)
    return dst


def run(script, env_extra):
    r = subprocess.run(PY + [script], capture_output=True, text=True, encoding="utf-8", env={**os.environ, **env_extra})
    open(script + ".log", "w", encoding="utf-8").write(r.stdout + "\n--- stderr ---\n" + r.stderr)
    if r.returncode != 0:
        print(r.stdout[-3000:], r.stderr[-3000:])
        raise SystemExit(f"FAIL {os.path.basename(script)} rc={r.returncode}")
    print("ok", os.path.basename(script))


def main():
    os.makedirs(CALC, exist_ok=True)
    base_raw = os.path.join(RUN, "z6_baseline_raw.json")
    with gzip.open(BASE_GZ, "rb") as fi, open(base_raw, "wb") as fo:
        shutil.copyfileobj(fi, fo)
    for f in ("base.py", "mapping.py", "views.py"):
        shutil.copy2(os.path.join(PROTO, f), CALC)
    collect = patched(os.path.join(Z6S, "z6_collect.py"), os.path.join(RUN, "z6_collect_e0.py"), [
        ('RES = r"C:\\Configuration_downloads\\BASERP25\\_Rarzrabotki\\Python\\test\\ekonomiya_ss_irc"',
         'RES = %r' % os.path.join(HERE, "vendor")),
        ('raw["stamp_1c"] = sv(B.ТекущаяДата())', 'raw["stamp_1c"] = sv(B.ТекущаяДатаСеанса())'),
    ])
    run(collect, {"Z6_RAW_DIR": CALC})
    calc = patched(os.path.join(Z6S, "z6_calc.py"), os.path.join(RUN, "z6_calc_e0.py"), [
        ('HERE = os.path.dirname(os.path.abspath(__file__))', 'HERE = %r' % Z6S),
        ('OUT_JSON = os.path.join(DATA, "proekt_irc2026_zakupki.json")', 'OUT_JSON = %r' % os.path.join(RUN, "e0_calc.json")),
        ('OUT_TXT = os.path.join(HERE, "z6_out.txt")', 'OUT_TXT = %r' % os.path.join(RUN, "e0_calc.txt")),
        ('_VER = os.path.join(HERE, "proverka", "q_t4_out.json")', '_VER = %r' % os.path.join(RUN, "no_verification.json")),
    ])
    run(calc, {"Z6_RAW_DIR": CALC, "Z1_RAW": base_raw, "Z3_RAW": os.path.join(RUN, "no_z3.json")})
    shutil.copy2(os.path.join(Z6S, "h.py"), os.path.join(RUN, "h.py"))
    t3 = patched(os.path.join(Z6S, "q_t3_proverka.py"), os.path.join(RUN, "q_t3_e0.py"), [
        ('os.path.join(os.path.dirname(os.path.abspath(__file__)), "q_t3_out.json")', '%r' % os.path.join(RUN, "q_t3_out.json")),
    ])
    run(t3, {})

    new = json.load(open(os.path.join(RUN, "e0_calc.json"), encoding="utf-8"))
    old = json.load(open(PUBLISHED, encoding="utf-8"))
    t3o = json.load(open(os.path.join(RUN, "q_t3_out.json"), encoding="utf-8"))
    T = new["totals"]
    vs = {k: {"z6": old["totals"].get(k), "now": T.get(k),
              "delta": round((T.get(k) or 0) - (old["totals"].get(k) or 0), 2)} for k in KEYS}
    closing = round(T["purchased_uah_net"] + T["other_in_uah_net"] - T["written_off_uah_net"]
                    - T["other_out_uah_net"] - T["lies_uah_net"], 2)
    indep = {"ptu_net": [t3o["ptu"][0]["Нетто"], T["purchased_by_kind"].get("ПТУ", {}).get("uah_net")],
             "stock_project": [t3o["stock_project"][0]["Сумма"], T["lies_project_uah_net"]],
             "written_off_sum_houses": [round(sum(x["Сумма"] for x in t3o["spys"]), 2), T["written_off_uah_net"]]}
    diff = new["snapshot"].get("diff_vs_z1") or {}
    ok = {"closing_zero": abs(closing) < 0.01,
          "independent_match": all(abs((x or 0) - (y or 0)) < 0.01 for x, y in indep.values()),
          "unchanged_or_explained": all(v["delta"] == 0 for v in vs.values()) or bool(diff.get("docs"))}
    summary = {"stamp": new["snapshot"]["datetime"], "stamp_1c": new["snapshot"].get("stamp_1c"),
               "base": 'Srvr="localhost";Ref="bas_industrialbud" (лише читання)',
               "baseline": {"z6_stamp": old["snapshot"]["datetime"], "raw": os.path.relpath(BASE_GZ, HERE)},
               "control_sums": {k: T.get(k) for k in KEYS}, "by_house": T["by_house"], "vs_z6": vs,
               "closing_t3": closing, "independent_q_t3": indep, "moves_diff_vs_z6": diff, "run_dir": RUN, "ok": ok}
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e0_snimok_{STAMP[:13]}.json")
    json.dump(summary, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print(json.dumps({"stamp": summary["stamp"], "delta": {k: v for k, v in vs.items() if v["delta"]},
                      "closing": closing, "independent": indep, "ok": ok}, ensure_ascii=False, indent=1))
    print("->", out)
    return 0 if all(ok.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/e0_snimok.py`
Expected: три строки `ok z6_collect_e0.py`, `ok z6_calc_e0.py`, `ok q_t3_e0.py`; `"closing": 0.0`; `independent` — пары равны (ПТУ 7 543 375,28; склад проекта 3 029 926,67; списано 4 332 862,86 на 04.10); `ok` — все `true`; код 0. Если `delta` не пуст — в `moves_diff_vs_z6.docs` перечислены регистраторы с датой и Δ (это и есть «расхождения объяснены документами после 03.10.2026 23:41»).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/e0_snimok.py _Rarzrabotki/Python/test/spec_domov/out/e0_snimok_*.json
git commit -m "feat(spec_domov): этап 0 — снимок контрольных сумм IRC 2026 против Z6 (только чтение)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Этап 0 — сверка конфигурации, обмена и ролей (`e0_kontrol.py`), фиксация решений

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/e0_kontrol.py`
- Create (результат): `_Rarzrabotki/Python/test/spec_domov/out/e0_kontrol_<YYYYMMDD_HHMM>.json`

- [ ] **Step 1: Тест** — скрипт возвращает 0, только если основной каталог = база (состав и выгрузка СС/РК/плана обмена), новые имена П1 свободны и СС 003 имеет 214 карточек.

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/e0_kontrol.py`
Expected: FAIL — `No such file or directory`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/e0_kontrol.py`:

```python
# -*- coding: utf-8 -*-
"""Этап 0 (§10.1): сверка дампа с живой конфигурацией, планы обмена, регистрация СС, роли и пользователи — только чтение.
Выход out/e0_kontrol_<stamp>.json; раздел exchange.changes — справочно (состояние регистрации СС на день этапа 0). Приёмки
Task 28 и 44 сверяют регистрацию НЕ с ним, а начало и конец своего прогона: между днями таблицу меняют обмен и правки
пользователей (память feedback_invariantnye_testy_ne_baseline)."""
import collections
import datetime
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import preflight_config  # noqa: E402
from common import BB_MAIN, BB_WT, HERE, OUT, connect, q  # noqa: E402

STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M")
NEW_NAMES = [n for _, n in preflight_config.EXPECTED_NEW]


def main():
    B = connect()
    md = B.Metadata
    res = {"stamp": STAMP}
    db = preflight_config.db_objects(B)
    for nm, root in (("main", BB_MAIN), ("worktree", BB_WT)):
        f = {x for x in preflight_config.file_objects(os.path.join(root, "Configuration.xml")) if x[0] in {k for k, _ in db}}
        res[f"configuration_{nm}_vs_db"] = {"file": len(f), "db": len(db),
                                           "db_minus_file": sorted(f"{k}.{n}" for k, n in db - f),
                                           "file_minus_db": sorted(f"{k}.{n}" for k, n in f - db)}
    r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "dump_check.py"), "--tag", "e0", "--objects",
                        "Справочник.СтруктураСебестоимости,Справочник.СтруктураСебестоимости.Форма.ФормаЭлемента,"
                        "Документ.РасчетКомплектаций,ПланОбмена.ОбменУправлениеПредприятиемБухгалтерия20"],
                       capture_output=True, text=True, encoding="utf-8")
    res["dump_vs_main"] = {"exit": r.returncode, "tail": r.stdout[-600:]}
    res["new_names_free"] = {n: (md.Catalogs.Find(n) is None and md.Enums.Find(n) is None and md.CommonModules.Find(n) is None)
                             for n in NEW_NAMES}
    probe = {"СтруктураСебестоимости": md.Catalogs.Find("СтруктураСебестоимости"),
             "РасчетКомплектаций": md.Documents.Find("РасчетКомплектаций")}
    plans = {k: [] for k in probe}
    for i in range(md.ExchangePlans.Count()):
        ep = md.ExchangePlans.Get(i)
        for k, m in probe.items():
            if ep.Content.Contains(m):
                plans[k].append(ep.Name)
    chg = q(B, """ВЫБРАТЬ
	Изм.Ссылка.Код КАК Код,
	ПРЕДСТАВЛЕНИЕ(Изм.Узел) КАК Узел,
	ЕСТЬNULL(Изм.НомерСообщения, -1) КАК НомерСообщения
ИЗ
	Справочник.СтруктураСебестоимости.Изменения КАК Изм
УПОРЯДОЧИТЬ ПО
	Код""")
    res["exchange"] = {"plans_with": plans,
                       "changes": [{"code": c["Код"], "node": c["Узел"], "msg": int(c["НомерСообщения"])} for c in chg]}
    cards = q(B, """ВЫБРАТЬ
	КОЛИЧЕСТВО(Компл.НомерСтроки) КАК Строк,
	СУММА(ВЫБОР КОГДА Компл.Номенклатура = ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка) ТОГДА 0 ИНАЧЕ 1 КОНЕЦ) КАК СКарточкой
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код""", Код="000000003")[0]
    res["ss003_cards"] = {"rows": int(cards["Строк"]), "with_card": int(cards["СКарточкой"])}
    roles = collections.Counter()
    users = B.ПользователиИнформационнойБазы.ПолучитьПользователей()
    for i in range(users.Количество()):
        u = users.Получить(i)
        names = {B.String(role_md.Имя) for role_md in u.Роли}    # КоллекцияРолейПользователя: только перебор, без Получить()
        key = ("ПолныеПрава" if "ПолныеПрава" in names else
               "+".join(sorted(names & {"ДобавлениеИзменениеДанныхБухгалтерии", "ЧтениеДанныхБухгалтерии", "ЗапускТонкогоКлиента"})) or "прочие")
        roles[key] += 1
    res["users_by_roles"] = dict(roles)
    write_ss = []
    for i in range(md.Roles.Count()):
        role = md.Roles.Get(i)
        if B.ПравоДоступа("Изменение", probe["СтруктураСебестоимости"], role):
            write_ss.append(role.Name)
    res["roles_write_ss"] = write_ss
    res["decisions"] = {"Д1": "правки в worktree на синхронизированных копиях, загрузка из основного каталога",
                        "Д2": "права (б), решение пользователя 04.10.2026: ЗапускТонкогоКлиента — полный набор (как у СС, правят все 24); ДобавлениеИзменениеДанныхБухгалтерии — полный набор; ЧтениеДанныхБухгалтерии — чтение",
                        "Д3": "«паков» не сводится к «пак»", "Т9": "Деталь без карточки — только в Чернетке",
                        "Т10": "зелёное «единственная сирота этапа»", "Т11": "узел-позиция ищется и по похожему тексту"}
    res["ok"] = {"main_eq_db": not res["configuration_main_vs_db"]["db_minus_file"]
                 and not res["configuration_main_vs_db"]["file_minus_db"],
                 "dump_eq_main": r.returncode == 0, "names_free": all(res["new_names_free"].values()),
                 "ss003_214": res["ss003_cards"] == {"rows": 214, "with_card": 214},
                 "rk_not_in_exchange": plans["РасчетКомплектаций"] == []}
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e0_kontrol_{STAMP}.json")
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({k: res[k] for k in ("ok", "users_by_roles", "roles_write_ss", "ss003_cards")}, ensure_ascii=False, indent=1))
    print("->", out)
    return 0 if all(res["ok"].values()) else 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/e0_kontrol.py`
Expected: `ok` — все `true`; `ss003_cards` = `{"rows": 214, "with_card": 214}`; `roles_write_ss` содержит `ЗапускТонкогоКлиента` и `ПолныеПрава`; `exchange.plans_with.СтруктураСебестоимости` = `["ОбменУправлениеПредприятиемБухгалтерия20"]`; код 0. Worktree-раздел (`configuration_worktree_vs_db`) после Task 2 тоже без расхождений.

- [ ] **Step 5: Коммит (приёмка этапа 0)**

```bash
git add _Rarzrabotki/Python/test/spec_domov/e0_kontrol.py _Rarzrabotki/Python/test/spec_domov/out/e0_kontrol_*.json
git commit -m "feat(spec_domov): этап 0 — основной каталог = база, обмен СС, роли; решения Д1–Д3, Т9–Т11

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Фаза 1A — метаданные этапа 1 (загрузка 1A)

### Task 9: COM-тест метаданных этапа 1 (падает до загрузки)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/test_e1_metadata.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e1_metadata.py`:

```python
# -*- coding: utf-8 -*-
"""Приёмка метаданных этапа 1 через COM (после загрузки 1A): состав, uk-синонимы, обмен, права, навигация, РК не проводится.
До загрузки падает: «перечисление А_ВидыСпецификацийДомов не найдено в базе». Списки ниже — единый источник для test_e1_files.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect  # noqa: E402

UK_ENUM = {
    "А_ВидыСпецификацийДомов": {"Дом": "Будинок", "Узел": "Вузол", "Закупочная": "Закупівельна"},
    "А_СтатусыСпецификацийДомов": {"Черновик": "Чернетка", "Действует": "Діє", "Закрыта": "Закрита"},
    "А_ВидыСтрокСпецификацииДомов": {"Деталь": "Деталь", "Узел": "Вузол",
                                     "ГруппаПоОбщемуНазванию": "Група за загальною назвою"},
    "А_ИсточникиКарточкиСпецификации": {"ЕдинственнаяКарточкаОН": "Єдина картка ОН",
                                        "СписаннаяНаДомаТипа": "Списувалась на будинки типу",
                                        "ИзДокументаКомплектации": "З документа комплектації",
                                        "КопияИзДругойСпецификации": "Скопійовано з іншої специфікації",
                                        "Вручную": "Вручну"},
    "А_ГруппыЗакупочнойСпецификации": {"Эталон": "Еталон", "ЭталонНеВыдавался": "Еталон — не видавалось",
                                       "ВнеЭталона": "Поза еталоном", "МШП": "МШП"},
    "А_ИсточникиЦеныЗакупки": {"ПоследняяЗакупка": "Остання закупівля",
                               "ПредыдущаяЗакупка": "Попередня закупівля (помилка одиниці)",
                               "ЦенаНормы": "Ціна норми (СС)", "СебестоимостьФакта": "Собівартість факту",
                               "Вручную": "Вручну"},
}
ENUMS = {name: list(values) for name, values in UK_ENUM.items()}
HEADER = ["ВидСпецификации", "СтруктураСебестоимости", "Основание", "Статус", "ДатаНачала", "КоличествоВыпуска",
          "ЕдиницаВыпуска", "НоменклатураВузла", "ДатаИзменения", "ДатаЦенНормы", "ДатаФакта", "Комментарий"]
TS = {
    "Состав": ["КлючСтроки", "ВидСтроки", "ЭтапСС", "ПозицияСС", "Номенклатура", "ИсточникКарточки", "Проверено",
               "НеПоставляется", "ОбщееНазвание", "Узел", "ПокупаетсяГотовым", "Количество", "ЕдиницаСтроки",
               "Коэффициент", "ИсточникКоэффициента", "ЦенаНормы", "Комментарий"],
    "Дома": ["Подразделение", "Комментарий"],
    "СоставОснования": ["КлючСтроки", "ВидСтроки", "ЭтапСС", "ПозицияСС", "Номенклатура", "ОбщееНазвание", "Узел",
                        "ПокупаетсяГотовым", "НеПоставляется", "Количество", "ЕдиницаСтроки", "Коэффициент", "ЦенаНормы"],
    "ДомаФакта": ["Подразделение", "ФактСумма", "Комментарий"],
    "ЗакупочныйСостав": ["КлючСтроки", "Номенклатура", "ГруппаЗакупки", "ОбщееНазвание", "Узел", "ПутьУзла", "ЭтапСС",
                         "ПозицияСС", "НормаЭталона", "СреднееФакта", "ДомовСФактом", "КоэффициентВариации",
                         "Количество", "ИзмененоВручную", "ПотребуетРешения", "Проверено", "ЦенаБезНДС", "ЦенаСНДС",
                         "ИсточникЦены", "ДокументЦены", "ЦенаНормы", "Комментарий"],
    "ФактПоДомам": ["Подразделение", "Номенклатура", "Списано", "СписаноСумма", "Лежит", "ЛежитСумма", "ИзСборки"],
}
RK_TS = {
    "НормаДома": ["КлючЛиста", "КлючСтроки", "КлючЭкземпляра", "Спецификация", "ВидСтроки", "Номенклатура",
                  "ОбщееНазвание", "ЭтапСС", "ПозицияСС", "ЕдиницаСтроки", "КоличествоЕдСтроки", "Количество",
                  "ЦенаНормы", "СуммаНормы", "ПутьУзла", "Уровень"],
    "ЭкземплярыУзлов": ["КлючЭкземпляра", "КлючЭкземпляраРодителя", "Узел", "НоменклатураВузла", "ПокупаетсяГотовым",
                        "ПутьУзла", "Уровень", "Количество", "СуммаНормы"],
}
ROLE_RIGHTS = {
    "ДобавлениеИзменениеДанныхБухгалтерии": ["Read", "Insert", "Update", "View", "InteractiveInsert", "Edit",
                                             "InteractiveSetDeletionMark", "InteractiveClearDeletionMark", "InputByString"],
    "ЧтениеДанныхБухгалтерии": ["Read", "View", "InputByString"],
    "ЗапускТонкогоКлиента": ["Read", "Insert", "Update", "View", "InteractiveInsert", "Edit",
                                             "InteractiveSetDeletionMark", "InteractiveClearDeletionMark", "InputByString"],
}


def names(coll):
    return [coll.Get(i).Name for i in range(coll.Count())]


def main():
    B = connect()
    md = B.Metadata
    fails = []

    def check(cond, text):
        if not cond:
            fails.append(text)

    for en, values in UK_ENUM.items():
        e = md.Enums.Find(en)
        if e is None:
            raise SystemExit(f"FAIL: перечисление {en} не найдено в базе")
        check(names(e.EnumValues) == list(values), f"{en}: значения {names(e.EnumValues)}")
        for i in range(e.EnumValues.Count()):
            v = e.EnumValues.Get(i)
            check(str(v.Synonym) == values.get(v.Name), f"{en}.{v.Name}: синоним uk «{v.Synonym}»")
    cat = md.Catalogs.Find("А_СпецификацииДомов")
    if cat is None:
        raise SystemExit("FAIL: справочник А_СпецификацииДомов не найден")
    check(str(cat.Synonym) == "Специфікації будинків і вузлів", f"синоним справочника «{cat.Synonym}»")
    check(bool(cat.Hierarchical), "справочник не иерархический")
    check(names(cat.Attributes) == HEADER, f"шапка {names(cat.Attributes)}")
    for i in range(cat.Attributes.Count()):
        check(str(cat.Attributes.Get(i).Synonym) != "", f"пустой uk-синоним {cat.Attributes.Get(i).Name}")
    check(names(cat.TabularSections) == list(TS), f"ТЧ {names(cat.TabularSections)}")
    for ts, cols in TS.items():
        t = cat.TabularSections.Find(ts)
        check(t is not None and names(t.Attributes) == cols, f"ТЧ {ts}: {names(t.Attributes) if t else None}")
        if t is not None:
            check(str(t.Synonym) != "", f"ТЧ {ts}: пустой uk-синоним")
    ss = md.Catalogs.Find("СтруктураСебестоимости")
    a = ss.Attributes.Find("ЭталоннаяСпецификация")
    check(a is not None and str(a.Synonym) == "Еталонна специфікація", "СС.ЭталоннаяСпецификация")
    rk = md.Documents.Find("РасчетКомплектаций")
    for n, syn in (("СпецификацияДома", "Специфікація будинку"), ("ДатаНормы", "Норма від"),
                   ("Спецификация", "Структура собівартості (СС)")):
        at = rk.Attributes.Find(n)
        check(at is not None and str(at.Synonym) == syn, f"РК.{n}: {at.Synonym if at else None}")
    for ts, cols in RK_TS.items():
        t = rk.TabularSections.Find(ts)
        check(t is not None and names(t.Attributes) == cols, f"РК.{ts}: {names(t.Attributes) if t else None}")
    ost = rk.TabularSections.Find("ТабличнаяЧастьОстатков")
    check(names(ost.Attributes)[-2:] == ["ПозицияСС", "НеИзНормы"], f"РК.ТабличнаяЧастьОстатков {names(ost.Attributes)[-2:]}")
    check(B.String(rk.Проведение) == B.String(B.Метаданные.СвойстваОбъектов.Проведение.Запретить), "РК стал проводимым")
    new_objs = [cat] + [md.Enums.Find(e) for e in UK_ENUM]
    for i in range(md.ExchangePlans.Count()):
        ep = md.ExchangePlans.Get(i)
        for o in new_objs:
            check(not ep.Content.Contains(o), f"{o.Name} в плане обмена {ep.Name}")
    for role, rights in ROLE_RIGHTS.items():
        r = md.Roles.Find(role)
        check(bool(B.ПравоДоступа("Чтение", cat, r)), f"{role}: нет Чтение")
        check(bool(B.ПравоДоступа("Изменение", cat, r)) == ("Update" in rights), f"{role}: Изменение")
    sub = md.Subsystems.Find("Склад").Subsystems.Find("А_СписанняБудинків")
    check(bool(sub.Content.Contains(cat)), "справочника нет в подсистеме А_СписанняБудинків")
    print(f"ПРОВАЛОВ: {len(fails)}")
    for f in fails:
        print("   ", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_metadata.py`
Expected: FAIL — `FAIL: перечисление А_ВидыСпецификацийДомов не найдено в базе`.

- [ ] **Step 3: Реализация** — задачи 10–16 (файлы метаданных); загрузка — Task 17.

- [ ] **Step 4: Запустить — PASS** — в Task 17 после загрузки 1A.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/test_e1_metadata.py
git commit -m "test(spec_domov): COM-приёмка метаданных этапа 1 (падает до загрузки 1A)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Статический тест файлов этапа 1A

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/test_e1_files.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e1_files.py`:

```python
# -*- coding: utf-8 -*-
"""Статическая приёмка файлов этапа 1A в дампе WORKTREE (до загрузки). Разделы: enums, catalog, existing, register.
Запуск: python test_e1_files.py [--only enums|catalog|existing|register]"""
import argparse
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT  # noqa: E402
from test_e1_metadata import HEADER, RK_TS, ROLE_RIGHTS, TS, UK_ENUM  # noqa: E402

RE_ATTR_NAME = re.compile(r"<Attribute uuid=\"[^\"]+\">\s*<Properties>\s*<Name>([^<]+)</Name>")


def load(rel):
    p = os.path.join(BB_WT, rel)
    assert os.path.exists(p), f"нет файла {rel}"
    raw = open(p, "rb").read()
    assert raw.startswith(b"\xef\xbb\xbf"), f"{rel}: нет BOM"
    assert raw.replace(b"\r\n", b"").count(b"\n") == 0, f"{rel}: не CRLF"
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")
    xml.dom.minidom.parseString(text.encode("utf-8"))
    return text


def clean_new(rel, text):
    assert " />" not in text and "d5p1" not in text, f"{rel}: шум генератора"
    ru, uk = text.count("<v8:lang>ru</v8:lang>"), text.count("<v8:lang>uk</v8:lang>")
    assert ru == uk, f"{rel}: ru {ru} ≠ uk {uk}"


def ts_block(text, ts):
    for m in re.finditer(r"\t\t\t<TabularSection uuid=.*?\n\t\t\t</TabularSection>\n", text, re.S):
        if f"<Name>{ts}</Name>" in m.group(0).split("<ChildObjects>")[0]:
            return m.group(0)
    raise AssertionError(f"нет ТЧ {ts}")


def attr_block(text, name):
    for m in re.finditer(r"<Attribute uuid=\"[^\"]+\">.*?</Attribute>", text, re.S):
        if f"<Name>{name}</Name>" in m.group(0):
            return m.group(0)
    raise AssertionError(f"нет реквизита {name}")


def check_enums():
    for name, values in UK_ENUM.items():
        rel = f"Enums/{name}.xml"
        t = load(rel)
        clean_new(rel, t)
        got = re.findall(r"<EnumValue uuid=\"[^\"]+\">\s*<Properties>\s*<Name>([^<]+)</Name>", t)
        assert got == list(values), (name, got)
        for uk in values.values():
            assert f"<v8:content>{uk}</v8:content>" in t, (name, uk)
    print("enums OK")


def check_catalog():
    rel = "Catalogs/А_СпецификацииДомов.xml"
    t = load(rel)
    clean_new(rel, t)
    for tag, val in (("Hierarchical", "true"), ("HierarchyType", "HierarchyFoldersAndItems"), ("CodeLength", "9"),
                     ("DescriptionLength", "150"), ("CheckUnique", "true"), ("Autonumbering", "true"),
                     ("DataLockControlMode", "Managed")):
        assert f"<{tag}>{val}</{tag}>" in t, (tag, val)
    assert "<Owners/>" in t
    head = t.split("<TabularSection", 1)[0]
    assert RE_ATTR_NAME.findall(head) == HEADER, RE_ATTR_NAME.findall(head)
    for ts, cols in TS.items():
        assert RE_ATTR_NAME.findall(ts_block(t, ts)) == cols, ts
    assert t.count("Enum.А_ВидыСпецификацийДомов.EnumValue.Узел") == 3
    assert t.count("<v8:Type>v8:UUID</v8:Type>") == 3
    assert "<v8:content>Специфікації будинків і вузлів</v8:content>" in t
    assert "<v8:Length>200</v8:Length>" in attr_block(ts_block(t, "Состав"), "ПозицияСС")
    q = attr_block(ts_block(t, "Состав"), "Количество")
    assert "<v8:Digits>15</v8:Digits>" in q and "<v8:FractionDigits>6</v8:FractionDigits>" in q
    k = attr_block(ts_block(t, "Состав"), "Коэффициент")
    assert "<v8:FractionDigits>9</v8:FractionDigits>" in k
    c = attr_block(ts_block(t, "Состав"), "ЦенаНормы")
    assert "<v8:Digits>16</v8:Digits>" in c and "<v8:FractionDigits>6</v8:FractionDigits>" in c
    print("catalog OK")


def check_existing():
    ss = load("Catalogs/СтруктураСебестоимости.xml")
    a = attr_block(ss, "ЭталоннаяСпецификация")
    assert "<v8:Type>cfg:CatalogRef.А_СпецификацииДомов</v8:Type>" in a and "Еталонна специфікація" in a
    assert ss.index("<Name>ЭталоннаяСпецификация</Name>") < ss.index("\t\t\t<TabularSection")
    assert ss.count("<Name>ЭталоннаяСпецификация</Name>") == 1
    rk = load("Documents/РасчетКомплектаций.xml")
    for n in ("СпецификацияДома", "ДатаНормы"):
        assert rk.count(f"<Name>{n}</Name>") == 1 and rk.index(f"<Name>{n}</Name>") < rk.index("<Form>ФормаДокумента</Form>")
    for ts, cols in RK_TS.items():
        assert RE_ATTR_NAME.findall(ts_block(rk, ts)) == cols, ts
    assert rk.index("<Name>НормаДома</Name>") < rk.index("<Template>МакетПланФакт</Template>")
    assert RE_ATTR_NAME.findall(ts_block(rk, "ТабличнаяЧастьОстатков"))[-2:] == ["ПозицияСС", "НеИзНормы"]
    sp = attr_block(rk, "Спецификация")
    assert "Структура собівартості (СС)" in sp and "Структура себестоимости (СС)" in sp and "(норми)" not in sp
    assert "<Posting>Deny</Posting>" in rk
    assert rk.count("\t\t\t<TabularSection uuid=") == 15 and rk.count("<Name>ПричиниЦеху</Name>") == 1
    print("existing OK")


def check_register():
    cfg = load("Configuration.xml")
    for tag, name in [("Catalog", "А_СпецификацииДомов")] + [("Enum", e) for e in UK_ENUM]:
        assert cfg.count(f"\t\t\t<{tag}>{name}</{tag}>\n") == 1, (tag, name)
    for role, rights in ROLE_RIGHTS.items():
        t = load(f"Roles/{role}/Ext/Rights.xml")
        m = re.search(r"<object>\s*<name>Catalog\.А_СпецификацииДомов</name>(.*?)</object>", t, re.S)
        assert m, role
        assert re.findall(r"<name>(\w+)</name>\s*<value>true</value>", m.group(1)) == rights, role
    sub = load("Subsystems/Склад/Subsystems/А_СписанняБудинків.xml")
    assert sub.count(">Catalog.А_СпецификацииДомов</xr:Item>") == 1
    ci = load("Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml")
    assert ci.count('"Catalog.А_СпецификацииДомов.StandardCommand.OpenList"') == 1
    print("register OK")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--only")
    a = ap.parse_args()
    parts = {"enums": check_enums, "catalog": check_catalog, "existing": check_existing, "register": check_register}
    for key, fn in parts.items():
        if a.only in (None, key):
            fn()
    print("OK test_e1_files", a.only or "all")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py`
Expected: FAIL — `AssertionError: нет файла Enums/А_ВидыСпецификацийДомов.xml`.

- [ ] **Step 3: Реализация** — задачи 11–15.
- [ ] **Step 4: PASS** — по разделам в задачах 12–15.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/test_e1_files.py
git commit -m "test(spec_domov): статическая приёмка файлов этапа 1A (перечисления, справочник, СС/РК, регистрация)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Постпатч после `meta-compile` (`postpatch.py`)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/postpatch.py`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_postpatch.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_postpatch.py`:

```python
# -*- coding: utf-8 -*-
"""postpatch.fix: cfg: вместо d5p1:, '/>' вместо ' />', uk к каждому ru, идемпотентность, отказ без перевода."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import postpatch  # noqa: E402
from common import read_text, write_text  # noqa: E402

HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" '
        'xmlns:app="http://v8.1c.ru/8.2/managed-application/core" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
        'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">\n')


def syn(ru, t):
    return (f"{t}<Synonym>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{ru}</v8:content>\n"
            f"{t}\t</v8:item>\n{t}</Synonym>\n")


def attr(name, ru):
    return (f'\t\t\t<Attribute uuid="{name}-u">\n\t\t\t\t<Properties>\n\t\t\t\t\t<Name>{name}</Name>\n' + syn(ru, "\t\t\t\t\t")
            + '\t\t\t\t\t<Type>\n\t\t\t\t\t\t<v8:Type xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config">'
              'd5p1:CatalogRef.А_СпецификацииДомов</v8:Type>\n\t\t\t\t\t</Type>\n\t\t\t\t\t<ChoiceParameters />\n'
              '\t\t\t\t</Properties>\n\t\t\t</Attribute>\n')


def test_enum_like():
    p = os.path.join(tempfile.mkdtemp(), "e.xml")
    write_text(p, HEAD + '\t<Enum uuid="1">\n\t\t<Properties>\n\t\t\t<Name>А_Тест</Name>\n' + syn("Узел", "\t\t\t")
               + "\t\t\t<Comment />\n\t\t</Properties>\n\t</Enum>\n</MetaDataObject>\n")
    postpatch.fix(p)
    text, bom, crlf = read_text(p)
    assert bom and crlf and " />" not in text
    assert "<v8:lang>uk</v8:lang>\n\t\t\t\t\t<v8:content>Вузол</v8:content>" in text, text
    postpatch.fix(p)
    assert read_text(p)[0].count("<v8:lang>uk</v8:lang>") == 1, "не идемпотентен"


def test_catalog_choice_and_cfg():
    p = os.path.join(tempfile.mkdtemp(), "c.xml")
    body = attr("Узел", "Узел") + attr("Основание", "Создано из") + attr("Узел", "Узел") + attr("Узел", "Узел строителя")
    write_text(p, HEAD + '\t<Catalog uuid="2">\n\t\t<Properties>\n\t\t\t<CheckUnique>false</CheckUnique>\n'
               "\t\t</Properties>\n\t\t<ChildObjects>\n" + body + "\t\t</ChildObjects>\n\t</Catalog>\n</MetaDataObject>\n")
    postpatch.fix(p, catalog=True)
    text = read_text(p)[0]
    assert "d5p1" not in text and text.count("<v8:Type>cfg:CatalogRef.А_СпецификацииДомов</v8:Type>") == 4
    assert text.count("Enum.А_ВидыСпецификацийДомов.EnumValue.Узел") == 3
    assert "<CheckUnique>true</CheckUnique>" in text


def test_missing_translation():
    p = os.path.join(tempfile.mkdtemp(), "m.xml")
    write_text(p, HEAD + '\t<Enum uuid="1">\n\t\t<Properties>\n' + syn("Слово без перевода", "\t\t\t")
               + "\t\t</Properties>\n\t</Enum>\n</MetaDataObject>\n")
    try:
        postpatch.fix(p)
    except AssertionError as e:
        assert "Слово без перевода" in str(e)
        return
    raise AssertionError("fix должен отказать без uk-перевода")


if __name__ == "__main__":
    test_enum_like()
    test_catalog_choice_and_cfg()
    test_missing_translation()
    print("OK test_postpatch")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_postpatch.py`
Expected: FAIL — `ModuleNotFoundError: No module named 'postpatch'`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/postpatch.py`:

```python
# -*- coding: utf-8 -*-
"""Нормализация файлов после meta-compile к канону дампа BASEBuh (E1 §7.1, §7.5):
1) <v8:Type xmlns:d5p1=…>d5p1: → <v8:Type>cfg:;  2) ' />' → '/>';
3) к каждому ru-синониму — uk-синоним по словарю UK (язык сеанса базы украинский: без uk подпись пустая,
   память uk_lang_sinonim_pustoy); нет перевода — AssertionError со списком;
4) справочник: CheckUnique=true; ChoiceParameters «только вид Вузол» у трёх реквизитов Узел (Состав, СоставОснования,
   ЗакупочныйСостав; образец — Контрагенты.ГоловнойКонтрагент).
Пишет UTF-8 BOM + CRLF. Идемпотентен."""
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import read_text, write_text  # noqa: E402

UK = {
    # перечисления: синонимы объектов
    "Виды спецификаций домов": "Види специфікацій будинків",
    "Статусы спецификаций домов": "Статуси специфікацій будинків",
    "Виды строк спецификации домов": "Види рядків специфікації будинків",
    "Источники карточки спецификации": "Джерела картки специфікації",
    "Группы закупочной спецификации": "Групи закупівельної специфікації",
    "Источники цены закупки": "Джерела ціни закупівлі",
    # значения перечислений
    "Дом": "Будинок", "Узел": "Вузол", "Закупочная": "Закупівельна",
    "Черновик": "Чернетка", "Действует": "Діє", "Закрыта": "Закрита",
    "Деталь": "Деталь", "Группа по общему названию": "Група за загальною назвою",
    "Единственная карточка ОН": "Єдина картка ОН", "Списывалась на дома типа": "Списувалась на будинки типу",
    "Из документа комплектации": "З документа комплектації",
    "Скопировано из другой спецификации": "Скопійовано з іншої специфікації", "Вручную": "Вручну",
    "Эталон": "Еталон", "Эталон — не выдавалось": "Еталон — не видавалось", "Вне эталона": "Поза еталоном", "МШП": "МШП",
    "Последняя закупка": "Остання закупівля", "Предыдущая закупка (ошибка единицы)": "Попередня закупівля (помилка одиниці)",
    "Цена нормы (СС)": "Ціна норми (СС)", "Себестоимость факта": "Собівартість факту",
    # справочник и шапка
    "Спецификации домов и узлов": "Специфікації будинків і вузлів",
    "Вид": "Вид", "Тип дома (СС)": "Тип будинку (СС)", "Создано из": "Створено з", "Статус": "Статус",
    "Действует с": "Діє з", "Количество выпуска": "Кількість випуску", "Единица выпуска": "Одиниця випуску",
    "Карточка узла (полуфабрикат / готовый)": "Картка вузла (напівфабрикат / готовий)", "Изменено": "Змінено",
    "Цены нормы на": "Ціни норми на", "Факт и цены на": "Факт і ціни на", "Комментарий": "Коментар",
    # табличные части
    "Состав": "Склад", "Дома": "Будинки", "Состав основания": "Склад основи", "Дома факта": "Будинки факту",
    "Закупочный состав": "Закупівельний склад", "Факт по домам": "Факт по будинках",
    # реквизиты ТЧ
    "Ключ строки": "Ключ рядка", "Вид строки": "Вид рядка", "Этап СС": "Етап СС", "Позиция СС": "Позиція СС",
    "Карточка": "Картка", "Откуда карточка": "Звідки картка", "Проверено": "Перевірено",
    "Не поставляется": "Не постачається", "Общее название": "Загальна назва", "Покупается готовым": "Купується готовим",
    "Количество": "Кількість", "Единица строки": "Одиниця рядка",
    "1 ед. карточки = k ед. строки": "1 од. картки = k од. рядка", "Источник коэффициента": "Джерело коефіцієнта",
    "Цена нормы, с НДС": "Ціна норми, з ПДВ", "Примечание": "Примітка", "Факт, без НДС": "Факт, без ПДВ",
    "Группа": "Група", "Узел строителя": "Вузол будівельника", "Путь узлов": "Шлях вузлів",
    "Норма эталона на дом": "Норма еталону на будинок", "Среднее факта на дом": "Середнє факту на будинок",
    "Домов с фактом": "Будинків з фактом", "Коэффициент вариации": "Коефіцієнт варіації",
    "Количество на дом": "Кількість на будинок", "Изменено вручную": "Змінено вручну",
    "Требует решения": "Потребує рішення", "Цена, без НДС": "Ціна, без ПДВ", "Цена, с НДС": "Ціна, з ПДВ",
    "Цена из": "Ціна з", "Документ цены": "Документ ціни", "Цена нормы (СС), с НДС": "Ціна норми (СС), з ПДВ",
    "Списано": "Списано", "Списано, без НДС": "Списано, без ПДВ", "Лежит на складе дома": "Лежить на складі будинку",
    "Лежит на складе дома, без НДС": "Лежить на складі будинку, без ПДВ", "Из полуфабриката": "З напівфабрикату",
    # общий модуль
    "Спецификации домов (сервер)": "Специфікації будинків (сервер)",
}

RE_RU = re.compile(r"(?P<ind>\t*)<v8:item>\n(?P=ind)\t<v8:lang>ru</v8:lang>\n"
                   r"(?P=ind)\t<v8:content>(?P<ru>[^<]*)</v8:content>\n(?P=ind)</v8:item>\n")
CHOICE = ("<ChoiceParameters>\n{t}\t<app:item name=\"Отбор.ВидСпецификации\">\n{t}\t\t<app:value xsi:type=\"xr:DesignTimeRef\">"
          "Enum.А_ВидыСпецификацийДомов.EnumValue.Узел</app:value>\n{t}\t</app:item>\n{t}</ChoiceParameters>")


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def unesc(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def add_uk(text):
    out, pos, missing = [], 0, []
    for m in RE_RU.finditer(text):
        ind = m.group("ind")
        if text.startswith(f"{ind}<v8:item>\n{ind}\t<v8:lang>uk</v8:lang>", m.end()):
            continue
        uk = UK.get(unesc(m.group("ru")))
        if uk is None:
            missing.append(unesc(m.group("ru")))
            continue
        out.append(text[pos:m.end()])
        out.append(f"{ind}<v8:item>\n{ind}\t<v8:lang>uk</v8:lang>\n{ind}\t<v8:content>{esc(uk)}</v8:content>\n{ind}</v8:item>\n")
        pos = m.end()
    out.append(text[pos:])
    assert not missing, f"нет uk-перевода для: {sorted(set(missing))}"
    return "".join(out)


def add_choice_node(text):
    count = [0]

    def repl(m):
        block = m.group(0)
        if "<Name>Узел</Name>" not in block or "<ChoiceParameters/>" not in block:
            return block
        t = re.search(r"\n(\t+)<ChoiceParameters/>", block).group(1)
        count[0] += 1
        return block.replace("<ChoiceParameters/>", CHOICE.format(t=t), 1)

    text = re.sub(r"<Attribute uuid=\"[^\"]+\">.*?</Attribute>", repl, text, flags=re.S)
    assert count[0] == 3, f"ChoiceParameters для Узел: {count[0]} (ожидалось 3)"
    return text


def fix(path, catalog=False):
    text, _, _ = read_text(path)
    text = text.replace('<v8:Type xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config">d5p1:', "<v8:Type>cfg:")
    assert "d5p1" not in text, path
    text = text.replace(" />", "/>")
    if catalog:
        assert text.count("<CheckUnique>") == 1, "CheckUnique"
        text = text.replace("<CheckUnique>false</CheckUnique>", "<CheckUnique>true</CheckUnique>")
        if "Enum.А_ВидыСпецификацийДомов.EnumValue.Узел" not in text:
            text = add_choice_node(text)
    text = add_uk(text)
    assert "xmlns:app=" in text and "xmlns:xr=" in text and "xmlns:xsi=" in text, path
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
```

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_postpatch.py`
Expected: `OK test_postpatch`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/postpatch.py _Rarzrabotki/Python/test/spec_domov/test_postpatch.py
git commit -m "feat(spec_domov): постпатч meta-compile — cfg:, '/>', uk-синонимы, ChoiceParameters узла

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Шесть перечислений спецификации (`gen_meta.py`)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/gen_meta.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/meta/e1_enums.json`
- Create (генерируются): `_Rarzrabotki/BASEBuh/Enums/А_ВидыСпецификацийДомов.xml`, `…/А_СтатусыСпецификацийДомов.xml`, `…/А_ВидыСтрокСпецификацииДомов.xml`, `…/А_ИсточникиКарточкиСпецификации.xml`, `…/А_ГруппыЗакупочнойСпецификации.xml`, `…/А_ИсточникиЦеныЗакупки.xml`
- Test: `test_e1_files.py --only enums`

- [ ] **Step 1: Тест уже написан** (Task 10, раздел `enums`).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only enums`
Expected: FAIL — `нет файла Enums/А_ВидыСпецификацийДомов.xml`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/meta/e1_enums.json`:

```json
[
  {"type": "Enum", "name": "А_ВидыСпецификацийДомов", "synonym": "Виды спецификаций домов",
   "values": [{"name": "Дом", "synonym": "Дом"}, {"name": "Узел", "synonym": "Узел"},
              {"name": "Закупочная", "synonym": "Закупочная"}]},
  {"type": "Enum", "name": "А_СтатусыСпецификацийДомов", "synonym": "Статусы спецификаций домов",
   "values": [{"name": "Черновик", "synonym": "Черновик"}, {"name": "Действует", "synonym": "Действует"},
              {"name": "Закрыта", "synonym": "Закрыта"}]},
  {"type": "Enum", "name": "А_ВидыСтрокСпецификацииДомов", "synonym": "Виды строк спецификации домов",
   "values": [{"name": "Деталь", "synonym": "Деталь"}, {"name": "Узел", "synonym": "Узел"},
              {"name": "ГруппаПоОбщемуНазванию", "synonym": "Группа по общему названию"}]},
  {"type": "Enum", "name": "А_ИсточникиКарточкиСпецификации", "synonym": "Источники карточки спецификации",
   "values": [{"name": "ЕдинственнаяКарточкаОН", "synonym": "Единственная карточка ОН"},
              {"name": "СписаннаяНаДомаТипа", "synonym": "Списывалась на дома типа"},
              {"name": "ИзДокументаКомплектации", "synonym": "Из документа комплектации"},
              {"name": "КопияИзДругойСпецификации", "synonym": "Скопировано из другой спецификации"},
              {"name": "Вручную", "synonym": "Вручную"}]},
  {"type": "Enum", "name": "А_ГруппыЗакупочнойСпецификации", "synonym": "Группы закупочной спецификации",
   "values": [{"name": "Эталон", "synonym": "Эталон"}, {"name": "ЭталонНеВыдавался", "synonym": "Эталон — не выдавалось"},
              {"name": "ВнеЭталона", "synonym": "Вне эталона"}, {"name": "МШП", "synonym": "МШП"}]},
  {"type": "Enum", "name": "А_ИсточникиЦеныЗакупки", "synonym": "Источники цены закупки",
   "values": [{"name": "ПоследняяЗакупка", "synonym": "Последняя закупка"},
              {"name": "ПредыдущаяЗакупка", "synonym": "Предыдущая закупка (ошибка единицы)"},
              {"name": "ЦенаНормы", "synonym": "Цена нормы (СС)"},
              {"name": "СебестоимостьФакта", "synonym": "Себестоимость факта"},
              {"name": "Вручную", "synonym": "Вручную"}]}
]
```

`_Rarzrabotki/Python/test/spec_domov/gen_meta.py`:

```python
# -*- coding: utf-8 -*-
"""Новые объекты П1 навыком meta-compile: генерация в scratch-каталог с заглушкой Configuration.xml версии 2.13
(meta-compile переписывает Configuration.xml целиком — E1 §1.3), постпатч, копия XML объекта в дамп WORKTREE.
Configuration.xml дампа не трогается: регистрация — register.py. Пустые модули (только BOM) не копируются.
ОПАСНО перегенерировать уже загруженный объект: новые UUID = другой объект для платформы. Существующий файл в worktree
не перезаписывается (--force — только до первой загрузки).
Запуск: python gen_meta.py meta/e1_enums.json [--force]"""
import argparse
import datetime
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import postpatch  # noqa: E402
from common import BB_WT, HERE, SKILLS, TMP, write_text  # noqa: E402

STUB = ('<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" version="2.13">'
        "<Configuration><ChildObjects/></Configuration></MetaDataObject>\n")
FOLDER = {"Catalog": "Catalogs", "Enum": "Enums", "CommonModule": "CommonModules"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()
    path = a.json if os.path.isabs(a.json) else os.path.join(HERE, a.json)
    defs = json.load(open(path, encoding="utf-8"))
    defs = defs if isinstance(defs, list) else [defs]
    scratch = os.path.join(TMP, "gen_meta", datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(scratch, exist_ok=True)
    write_text(os.path.join(scratch, "Configuration.xml"), STUB)
    r = subprocess.run(["powershell.exe", "-NoProfile", "-File", os.path.join(SKILLS, r"meta-compile\scripts\meta-compile.ps1"),
                        "-JsonPath", path, "-OutputDir", scratch], capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout[-2000:])
    if r.returncode != 0:
        print(r.stderr[-2000:])
        return 1
    for d in defs:
        rel = f"{FOLDER[d['type']]}/{d['name']}.xml"
        src, dst = os.path.join(scratch, rel), os.path.join(BB_WT, rel)
        assert os.path.exists(src), f"meta-compile не создал {rel}"
        if os.path.exists(dst) and not a.force:
            print("  ПРОПУСК (уже есть — UUID не перегенерируем):", rel)
            continue
        postpatch.fix(src, catalog=(d["type"] == "Catalog"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(src, dst)
        print("  создан", rel)
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/gen_meta.py meta/e1_enums.json`
Expected: 6 строк `создан Enums/…xml`. Если `postpatch` падает «нет uk-перевода для: [...]» — `meta-compile` сгенерировал ru-синоним, которого нет в `UK`: добавить строку в `postpatch.UK` и повторить с `--force`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only enums`
Expected: `enums OK`, `OK test_e1_files enums`.
Run: `powershell.exe -NoProfile -File C:/Configuration_downloads/BASERP25/.claude/skills/meta-validate/scripts/meta-validate.ps1 -ObjectPath "_Rarzrabotki/BASEBuh/Enums/А_ВидыСпецификацийДомов.xml"` (и для остальных пяти)
Expected: без ошибок, кроме известной ложной «is not a valid 1C identifier» на именах с і/ї/є (в именах перечислений их нет — ошибок 0).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/gen_meta.py _Rarzrabotki/Python/test/spec_domov/meta/e1_enums.json \
  "_Rarzrabotki/BASEBuh/Enums/А_ВидыСпецификацийДомов.xml" "_Rarzrabotki/BASEBuh/Enums/А_СтатусыСпецификацийДомов.xml" \
  "_Rarzrabotki/BASEBuh/Enums/А_ВидыСтрокСпецификацииДомов.xml" "_Rarzrabotki/BASEBuh/Enums/А_ИсточникиКарточкиСпецификации.xml" \
  "_Rarzrabotki/BASEBuh/Enums/А_ГруппыЗакупочнойСпецификации.xml" "_Rarzrabotki/BASEBuh/Enums/А_ИсточникиЦеныЗакупки.xml"
git commit -m "feat(BASEBuh): 6 перечислений многоуровневой спецификации домов (§4.10)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 13: Справочник `А_СпецификацииДомов` (метаданные §4.1–4.5)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/meta/e1_catalog.json`
- Create (генерируется): `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml`
- Test: `test_e1_files.py --only catalog`

- [ ] **Step 1: Тест уже написан** (Task 10, раздел `catalog`).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only catalog`
Expected: FAIL — `нет файла Catalogs/А_СпецификацииДомов.xml`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/meta/e1_catalog.json`:

```json
{
  "type": "Catalog", "name": "А_СпецификацииДомов", "synonym": "Спецификации домов и узлов",
  "hierarchical": true, "codeLength": 9, "codeType": "String", "descriptionLength": 150,
  "checkUnique": true, "autonumbering": true, "dataLockControlMode": "Managed",
  "attributes": [
    {"name": "ВидСпецификации", "type": "EnumRef.А_ВидыСпецификацийДомов", "synonym": "Вид"},
    {"name": "СтруктураСебестоимости", "type": "CatalogRef.СтруктураСебестоимости", "synonym": "Тип дома (СС)", "flags": ["index"]},
    {"name": "Основание", "type": "CatalogRef.А_СпецификацииДомов", "synonym": "Создано из"},
    {"name": "Статус", "type": "EnumRef.А_СтатусыСпецификацийДомов", "synonym": "Статус"},
    {"name": "ДатаНачала", "type": "Date", "synonym": "Действует с"},
    {"name": "КоличествоВыпуска", "type": "Number(15,3,nonneg)", "synonym": "Количество выпуска"},
    {"name": "ЕдиницаВыпуска", "type": "CatalogRef.КлассификаторЕдиницИзмерения", "synonym": "Единица выпуска"},
    {"name": "НоменклатураВузла", "type": "CatalogRef.Номенклатура", "synonym": "Карточка узла (полуфабрикат / готовый)"},
    {"name": "ДатаИзменения", "type": "DateTime", "synonym": "Изменено"},
    {"name": "ДатаЦенНормы", "type": "Date", "synonym": "Цены нормы на"},
    {"name": "ДатаФакта", "type": "DateTime", "synonym": "Факт и цены на"},
    {"name": "Комментарий", "type": "String(1000)", "synonym": "Комментарий"}
  ],
  "tabularSections": [
    {"name": "Состав", "attributes": [
      {"name": "КлючСтроки", "type": "v8:UUID", "synonym": "Ключ строки"},
      {"name": "ВидСтроки", "type": "EnumRef.А_ВидыСтрокСпецификацииДомов", "synonym": "Вид строки"},
      {"name": "ЭтапСС", "type": "CatalogRef.ЭтапыРабот", "synonym": "Этап СС"},
      {"name": "ПозицияСС", "type": "String(200)", "synonym": "Позиция СС"},
      {"name": "Номенклатура", "type": "CatalogRef.Номенклатура", "synonym": "Карточка"},
      {"name": "ИсточникКарточки", "type": "EnumRef.А_ИсточникиКарточкиСпецификации", "synonym": "Откуда карточка"},
      {"name": "Проверено", "type": "Boolean", "synonym": "Проверено"},
      {"name": "НеПоставляется", "type": "Boolean", "synonym": "Не поставляется"},
      {"name": "ОбщееНазвание", "type": "CatalogRef.ОбщиеНазванияНоменклатуры", "synonym": "Общее название"},
      {"name": "Узел", "type": "CatalogRef.А_СпецификацииДомов", "synonym": "Узел"},
      {"name": "ПокупаетсяГотовым", "type": "Boolean", "synonym": "Покупается готовым"},
      {"name": "Количество", "type": "Number(15,6,nonneg)", "synonym": "Количество"},
      {"name": "ЕдиницаСтроки", "type": "CatalogRef.КлассификаторЕдиницИзмерения", "synonym": "Единица строки"},
      {"name": "Коэффициент", "type": "Number(15,9,nonneg)", "synonym": "1 ед. карточки = k ед. строки"},
      {"name": "ИсточникКоэффициента", "type": "EnumRef.А_ИсточникиКоэффициентаПересчетаСС", "synonym": "Источник коэффициента"},
      {"name": "ЦенаНормы", "type": "Number(16,6,nonneg)", "synonym": "Цена нормы, с НДС"},
      {"name": "Комментарий", "type": "String(250)", "synonym": "Примечание"}
    ]},
    {"name": "Дома", "attributes": [
      {"name": "Подразделение", "type": "CatalogRef.Подразделения", "synonym": "Дом"},
      {"name": "Комментарий", "type": "String(250)", "synonym": "Примечание"}
    ]},
    {"name": "СоставОснования", "attributes": [
      {"name": "КлючСтроки", "type": "v8:UUID", "synonym": "Ключ строки"},
      {"name": "ВидСтроки", "type": "EnumRef.А_ВидыСтрокСпецификацииДомов", "synonym": "Вид строки"},
      {"name": "ЭтапСС", "type": "CatalogRef.ЭтапыРабот", "synonym": "Этап СС"},
      {"name": "ПозицияСС", "type": "String(200)", "synonym": "Позиция СС"},
      {"name": "Номенклатура", "type": "CatalogRef.Номенклатура", "synonym": "Карточка"},
      {"name": "ОбщееНазвание", "type": "CatalogRef.ОбщиеНазванияНоменклатуры", "synonym": "Общее название"},
      {"name": "Узел", "type": "CatalogRef.А_СпецификацииДомов", "synonym": "Узел"},
      {"name": "ПокупаетсяГотовым", "type": "Boolean", "synonym": "Покупается готовым"},
      {"name": "НеПоставляется", "type": "Boolean", "synonym": "Не поставляется"},
      {"name": "Количество", "type": "Number(15,6,nonneg)", "synonym": "Количество"},
      {"name": "ЕдиницаСтроки", "type": "CatalogRef.КлассификаторЕдиницИзмерения", "synonym": "Единица строки"},
      {"name": "Коэффициент", "type": "Number(15,9,nonneg)", "synonym": "1 ед. карточки = k ед. строки"},
      {"name": "ЦенаНормы", "type": "Number(16,6,nonneg)", "synonym": "Цена нормы, с НДС"}
    ]},
    {"name": "ДомаФакта", "attributes": [
      {"name": "Подразделение", "type": "CatalogRef.Подразделения", "synonym": "Дом"},
      {"name": "ФактСумма", "type": "Number(15,2)", "synonym": "Факт, без НДС"},
      {"name": "Комментарий", "type": "String(250)", "synonym": "Примечание"}
    ]},
    {"name": "ЗакупочныйСостав", "attributes": [
      {"name": "КлючСтроки", "type": "v8:UUID", "synonym": "Ключ строки"},
      {"name": "Номенклатура", "type": "CatalogRef.Номенклатура", "synonym": "Карточка"},
      {"name": "ГруппаЗакупки", "type": "EnumRef.А_ГруппыЗакупочнойСпецификации", "synonym": "Группа"},
      {"name": "ОбщееНазвание", "type": "CatalogRef.ОбщиеНазванияНоменклатуры", "synonym": "Общее название"},
      {"name": "Узел", "type": "CatalogRef.А_СпецификацииДомов", "synonym": "Узел строителя"},
      {"name": "ПутьУзла", "type": "String(200)", "synonym": "Путь узлов"},
      {"name": "ЭтапСС", "type": "CatalogRef.ЭтапыРабот", "synonym": "Этап СС"},
      {"name": "ПозицияСС", "type": "String(200)", "synonym": "Позиция СС"},
      {"name": "НормаЭталона", "type": "Number(15,6)", "synonym": "Норма эталона на дом"},
      {"name": "СреднееФакта", "type": "Number(15,6)", "synonym": "Среднее факта на дом"},
      {"name": "ДомовСФактом", "type": "Number(3,0)", "synonym": "Домов с фактом"},
      {"name": "КоэффициентВариации", "type": "Number(7,3)", "synonym": "Коэффициент вариации"},
      {"name": "Количество", "type": "Number(15,6,nonneg)", "synonym": "Количество на дом"},
      {"name": "ИзмененоВручную", "type": "Boolean", "synonym": "Изменено вручную"},
      {"name": "ПотребуетРешения", "type": "Boolean", "synonym": "Требует решения"},
      {"name": "Проверено", "type": "Boolean", "synonym": "Проверено"},
      {"name": "ЦенаБезНДС", "type": "Number(16,6)", "synonym": "Цена, без НДС"},
      {"name": "ЦенаСНДС", "type": "Number(16,6)", "synonym": "Цена, с НДС"},
      {"name": "ИсточникЦены", "type": "EnumRef.А_ИсточникиЦеныЗакупки", "synonym": "Цена из"},
      {"name": "ДокументЦены", "type": "DocumentRef.ПоступлениеТоваровУслуг", "synonym": "Документ цены"},
      {"name": "ЦенаНормы", "type": "Number(16,6)", "synonym": "Цена нормы (СС), с НДС"},
      {"name": "Комментарий", "type": "String(250)", "synonym": "Примечание"}
    ]},
    {"name": "ФактПоДомам", "attributes": [
      {"name": "Подразделение", "type": "CatalogRef.Подразделения", "synonym": "Дом"},
      {"name": "Номенклатура", "type": "CatalogRef.Номенклатура", "synonym": "Карточка"},
      {"name": "Списано", "type": "Number(15,6)", "synonym": "Списано"},
      {"name": "СписаноСумма", "type": "Number(15,2)", "synonym": "Списано, без НДС"},
      {"name": "Лежит", "type": "Number(15,6)", "synonym": "Лежит на складе дома"},
      {"name": "ЛежитСумма", "type": "Number(15,2)", "synonym": "Лежит на складе дома, без НДС"},
      {"name": "ИзСборки", "type": "Boolean", "synonym": "Из полуфабриката"}
    ]}
  ]
}
```

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/gen_meta.py meta/e1_catalog.json`
Expected: `создан Catalogs/А_СпецификацииДомов.xml`. Пустые `Ext/ObjectModule.bsl`/`ManagerModule.bsl` из scratch в дамп не попадают (копируется только корневой XML).

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only catalog`
Expected: `catalog OK`.
Run: `powershell.exe -NoProfile -File C:/Configuration_downloads/BASERP25/.claude/skills/meta-validate/scripts/meta-validate.ps1 -ObjectPath "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml"`
Expected: без ошибок (ссылка на ещё незарегистрированные перечисления — допустимое предупреждение валидатора до Task 15).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/meta/e1_catalog.json "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml"
git commit -m "feat(BASEBuh): справочник А_СпецификацииДомов — шапка и 6 ТЧ, включая закупочные (§4.1–4.5)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 14: Новые реквизиты СС и РК (`patch_existing.py`, часть метаданных)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/patch_existing.py`
- Modify: `_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости.xml` (после реквизита `ЗапретитьРедактированиеСС`, перед первой `<TabularSection>`)
- Modify: `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml` (перед `<Form>ФормаДокумента</Form>`; перед `<Template>МакетПланФакт</Template>`; конец `ChildObjects` ТЧ `ТабличнаяЧастьОстатков`; синоним `Спецификация`)
- Test: `test_e1_files.py --only existing`

- [ ] **Step 1: Тест уже написан** (Task 10, раздел `existing`).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only existing`
Expected: FAIL — `AssertionError: нет реквизита ЭталоннаяСпецификация`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/patch_existing.py`:

```python
# -*- coding: utf-8 -*-
"""Хирургические вставки П1 в СУЩЕСТВУЮЩИЕ объекты дампа WORKTREE (после sync_main_to_wt.py; Д1).
--etap meta: СС.ЭталоннаяСпецификация; РК: СпецификацияДома, ДатаНормы, ТЧ НормаДома и ЭкземплярыУзлов,
             колонки ТабличнаяЧастьОстатков.ПозицияСС/НеИзНормы, синоним Спецификация (§4.7).
--etap form: поле ЭталоннаяСпецификация (только просмотр) в Группа4 формы СС (Task 27).
Каждая вставка — по уникальному якорю с assert; повторный запуск ничего не дублирует. UUID детерминированные (uuid5).
Блоки XML — по образцу соседних элементов тех же файлов (E3 §1.3, §2; порядок свойств реквизита шапки/ТЧ — E1 §2.3–2.4)."""
import argparse
import os
import re
import sys
import uuid
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, read_text, write_text  # noqa: E402

NS = uuid.UUID("a6f1d1e0-5b8c-4f43-9d0e-2c7b6a91f3d4")


def uid(key):
    return str(uuid.uuid5(NS, key))


def syn(ru, uk, t):
    return (f"{t}<Synonym>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{ru}</v8:content>\n{t}\t</v8:item>\n"
            f"{t}\t<v8:item>\n{t}\t\t<v8:lang>uk</v8:lang>\n{t}\t\t<v8:content>{uk}</v8:content>\n{t}\t</v8:item>\n{t}</Synonym>\n")


def type_xml(spec, t):
    if spec.startswith("cfg:"):
        return f"{t}<v8:Type>{spec}</v8:Type>\n"
    if spec == "uuid":
        return f"{t}<v8:Type>v8:UUID</v8:Type>\n"
    if spec == "bool":
        return f"{t}<v8:Type>xs:boolean</v8:Type>\n"
    if spec == "datetime":
        return (f"{t}<v8:Type>xs:dateTime</v8:Type>\n{t}<v8:DateQualifiers>\n{t}\t<v8:DateFractions>DateTime</v8:DateFractions>\n"
                f"{t}</v8:DateQualifiers>\n")
    if spec.startswith("str:"):
        return (f"{t}<v8:Type>xs:string</v8:Type>\n{t}<v8:StringQualifiers>\n{t}\t<v8:Length>{spec[4:]}</v8:Length>\n"
                f"{t}\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{t}</v8:StringQualifiers>\n")
    if spec.startswith("num:"):
        d, f = spec[4:].split(",")
        return (f"{t}<v8:Type>xs:decimal</v8:Type>\n{t}<v8:NumberQualifiers>\n{t}\t<v8:Digits>{d}</v8:Digits>\n"
                f"{t}\t<v8:FractionDigits>{f}</v8:FractionDigits>\n{t}\t<v8:AllowedSign>Any</v8:AllowedSign>\n{t}</v8:NumberQualifiers>\n")
    raise ValueError(spec)


def fill_value(spec):
    if spec == "bool":
        return '<FillValue xsi:type="xs:boolean">false</FillValue>'
    if spec.startswith("str:"):
        return '<FillValue xsi:type="xs:string"/>'
    return '<FillValue xsi:nil="true"/>'


def attr(key, name, ru, uk, spec, t, kind):
    """kind: header_catalog | header_document | ts (реквизит ТЧ — без FillFromFillingValue/FillValue/Use)."""
    p = t + "\t\t"
    s = f'{t}<Attribute uuid="{uid(key)}">\n{t}\t<Properties>\n{p}<Name>{name}</Name>\n' + syn(ru, uk, p)
    s += f"{p}<Comment/>\n{p}<Type>\n" + type_xml(spec, p + "\t") + f"{p}</Type>\n"
    s += (f"{p}<PasswordMode>false</PasswordMode>\n{p}<Format/>\n{p}<EditFormat/>\n{p}<ToolTip/>\n"
          f"{p}<MarkNegatives>false</MarkNegatives>\n{p}<Mask/>\n{p}<MultiLine>false</MultiLine>\n"
          f"{p}<ExtendedEdit>false</ExtendedEdit>\n{p}<MinValue xsi:nil=\"true\"/>\n{p}<MaxValue xsi:nil=\"true\"/>\n")
    if kind != "ts":
        s += f"{p}<FillFromFillingValue>false</FillFromFillingValue>\n{p}{fill_value(spec)}\n"
    s += (f"{p}<FillChecking>DontCheck</FillChecking>\n{p}<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>\n"
          f"{p}<ChoiceParameterLinks/>\n{p}<ChoiceParameters/>\n{p}<QuickChoice>Auto</QuickChoice>\n"
          f"{p}<CreateOnInput>Auto</CreateOnInput>\n{p}<ChoiceForm/>\n{p}<LinkByType/>\n"
          f"{p}<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>\n")
    if kind == "header_catalog":
        s += f"{p}<Use>ForItem</Use>\n"
    s += f"{p}<Indexing>DontIndex</Indexing>\n{p}<FullTextSearch>Use</FullTextSearch>\n{p}<DataHistory>Use</DataHistory>\n"
    return s + f"{t}\t</Properties>\n{t}</Attribute>\n"


LINE_NUMBER = "".join(f"\t\t\t\t\t\t\t{x}\n" for x in (
    "<xr:LinkByType/>", "<xr:FillChecking>DontCheck</xr:FillChecking>", "<xr:MultiLine>false</xr:MultiLine>",
    "<xr:FillFromFillingValue>false</xr:FillFromFillingValue>", "<xr:CreateOnInput>Auto</xr:CreateOnInput>",
    '<xr:MaxValue xsi:nil="true"/>', "<xr:ToolTip/>", "<xr:ExtendedEdit>false</xr:ExtendedEdit>", "<xr:Format/>",
    "<xr:ChoiceForm/>", "<xr:QuickChoice>Auto</xr:QuickChoice>", "<xr:ChoiceHistoryOnInput>Auto</xr:ChoiceHistoryOnInput>",
    "<xr:EditFormat/>", "<xr:PasswordMode>false</xr:PasswordMode>", "<xr:DataHistory>Use</xr:DataHistory>",
    "<xr:MarkNegatives>false</xr:MarkNegatives>", '<xr:MinValue xsi:nil="true"/>', "<xr:Synonym/>", "<xr:Comment/>",
    "<xr:FullTextSearch>Use</xr:FullTextSearch>", "<xr:ChoiceParameterLinks/>", '<xr:FillValue xsi:nil="true"/>',
    "<xr:Mask/>", "<xr:ChoiceParameters/>"))


def doc_ts(doc, ts, ru, uk, cols):
    s = (f'\t\t\t<TabularSection uuid="{uid(doc + "." + ts)}">\n\t\t\t\t<InternalInfo>\n'
         f'\t\t\t\t\t<xr:GeneratedType name="DocumentTabularSection.{doc}.{ts}" category="TabularSection">\n'
         f'\t\t\t\t\t\t<xr:TypeId>{uid(doc + "." + ts + ".t1")}</xr:TypeId>\n\t\t\t\t\t\t<xr:ValueId>{uid(doc + "." + ts + ".v1")}</xr:ValueId>\n'
         f'\t\t\t\t\t</xr:GeneratedType>\n'
         f'\t\t\t\t\t<xr:GeneratedType name="DocumentTabularSectionRow.{doc}.{ts}" category="TabularSectionRow">\n'
         f'\t\t\t\t\t\t<xr:TypeId>{uid(doc + "." + ts + ".t2")}</xr:TypeId>\n\t\t\t\t\t\t<xr:ValueId>{uid(doc + "." + ts + ".v2")}</xr:ValueId>\n'
         f"\t\t\t\t\t</xr:GeneratedType>\n\t\t\t\t</InternalInfo>\n\t\t\t\t<Properties>\n\t\t\t\t\t<Name>{ts}</Name>\n"
         + syn(ru, uk, "\t\t\t\t\t") +
         "\t\t\t\t\t<Comment/>\n\t\t\t\t\t<ToolTip/>\n\t\t\t\t\t<FillChecking>DontCheck</FillChecking>\n"
         '\t\t\t\t\t<StandardAttributes>\n\t\t\t\t\t\t<xr:StandardAttribute name="LineNumber">\n' + LINE_NUMBER +
         "\t\t\t\t\t\t</xr:StandardAttribute>\n\t\t\t\t\t</StandardAttributes>\n\t\t\t\t</Properties>\n\t\t\t\t<ChildObjects>\n")
    for name, cru, cuk, spec in cols:
        s += attr(f"{doc}.{ts}.{name}", name, cru, cuk, spec, "\t\t\t\t\t", "ts")
    return s + "\t\t\t\t</ChildObjects>\n\t\t\t</TabularSection>\n"


REF_SPEC = "cfg:CatalogRef.А_СпецификацииДомов"
NORMA = [
    ("КлючЛиста", "Ключ листа", "Ключ листа", "str:32"), ("КлючСтроки", "Ключ строки", "Ключ рядка", "uuid"),
    ("КлючЭкземпляра", "Ключ экземпляра", "Ключ екземпляра", "str:32"),
    ("Спецификация", "Спецификация", "Специфікація", REF_SPEC),
    ("ВидСтроки", "Вид строки", "Вид рядка", "cfg:EnumRef.А_ВидыСтрокСпецификацииДомов"),
    ("Номенклатура", "Карточка", "Картка", "cfg:CatalogRef.Номенклатура"),
    ("ОбщееНазвание", "Общее название", "Загальна назва", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры"),
    ("ЭтапСС", "Этап СС", "Етап СС", "cfg:CatalogRef.ЭтапыРабот"), ("ПозицияСС", "Позиция СС", "Позиція СС", "str:200"),
    ("ЕдиницаСтроки", "Единица строки", "Одиниця рядка", "cfg:CatalogRef.КлассификаторЕдиницИзмерения"),
    ("КоличествоЕдСтроки", "Количество в ед. строки", "Кількість в од. рядка", "num:15,6"),
    ("Количество", "Количество", "Кількість", "num:15,6"), ("ЦенаНормы", "Цена нормы, с НДС", "Ціна норми, з ПДВ", "num:16,6"),
    ("СуммаНормы", "Сумма нормы, с НДС", "Сума норми, з ПДВ", "num:15,2"), ("ПутьУзла", "Путь узлов", "Шлях вузлів", "str:200"),
    ("Уровень", "Уровень", "Рівень", "num:2,0"),
]
EKZ = [
    ("КлючЭкземпляра", "Ключ экземпляра", "Ключ екземпляра", "str:32"),
    ("КлючЭкземпляраРодителя", "Ключ экземпляра родителя", "Ключ екземпляра батька", "str:32"),
    ("Узел", "Узел", "Вузол", REF_SPEC), ("НоменклатураВузла", "Карточка узла", "Картка вузла", "cfg:CatalogRef.Номенклатура"),
    ("ПокупаетсяГотовым", "Покупается готовым", "Купується готовим", "bool"), ("ПутьУзла", "Путь узлов", "Шлях вузлів", "str:200"),
    ("Уровень", "Уровень", "Рівень", "num:2,0"), ("Количество", "Узлов на дом", "Вузлів на будинок", "num:15,6"),
    ("СуммаНормы", "Сумма нормы узла, с НДС", "Сума норми вузла, з ПДВ", "num:15,2"),
]


def rep_once(text, old, new, what):
    assert text.count(old) == 1, f"якорь «{what}» найден {text.count(old)} раз"
    return text.replace(old, new, 1)


def patch_ss():
    rel = "Catalogs/СтруктураСебестоимости.xml"
    path = os.path.join(BB_WT, rel)
    text, _, _ = read_text(path)
    if "<Name>ЭталоннаяСпецификация</Name>" in text:
        print("  СС: уже есть")
        return
    block = attr("СС.ЭталоннаяСпецификация", "ЭталоннаяСпецификация", "Эталонная спецификация", "Еталонна специфікація",
                 REF_SPEC, "\t\t\t", "header_catalog")
    first_ts = "\t\t\t<TabularSection uuid="
    i = text.index(first_ts)
    assert text[:i].rstrip("\n").endswith("</Attribute>"), "перед первой ТЧ СС не реквизит"
    assert "<Name>ЗапретитьРедактированиеСС</Name>" in text[:i]
    text = text[:i] + block + text[i:]
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
    print("  СС: + ЭталоннаяСпецификация")


def patch_rk():
    rel = "Documents/РасчетКомплектаций.xml"
    path = os.path.join(BB_WT, rel)
    text, _, _ = read_text(path)
    assert "<Name>ПричиниЦеху</Name>" in text, "РК не из основного каталога (нет ТЧ ПричиниЦеху) — сначала sync_main_to_wt.py"
    doc = "РасчетКомплектаций"
    if "<Name>СпецификацияДома</Name>" not in text:
        hdr = (attr(doc + ".СпецификацияДома", "СпецификацияДома", "Спецификация дома", "Специфікація будинку", REF_SPEC,
                    "\t\t\t", "header_document")
               + attr(doc + ".ДатаНормы", "ДатаНормы", "Норма от", "Норма від", "datetime", "\t\t\t", "header_document"))
        anchor = "\t\t\t<Form>ФормаДокумента</Form>\n"
        text = rep_once(text, anchor, hdr + anchor, "Form ФормаДокумента")
    if "<Name>НормаДома</Name>" not in text:
        ts = (doc_ts(doc, "НормаДома", "Норма дома (снимок)", "Норма будинку (знімок)", NORMA)
              + doc_ts(doc, "ЭкземплярыУзлов", "Узлы дома (снимок)", "Вузли будинку (знімок)", EKZ))
        anchor = "\t\t\t<Template>МакетПланФакт</Template>\n"
        text = rep_once(text, anchor, ts + anchor, "Template МакетПланФакт")
    start = text.index("<Name>ТабличнаяЧастьОстатков</Name>")
    end_ts = text.index("\t\t\t\t</ChildObjects>\n\t\t\t</TabularSection>\n", start)
    if "<Name>НеИзНормы</Name>" not in text[start:end_ts]:
        cols = (attr(doc + ".ТабличнаяЧастьОстатков.ПозицияСС", "ПозицияСС", "Позиция СС", "Позиція СС", "str:200", "\t\t\t\t\t", "ts")
                + attr(doc + ".ТабличнаяЧастьОстатков.НеИзНормы", "НеИзНормы", "Не из нормы", "Не з норми", "bool", "\t\t\t\t\t", "ts"))
        text = text[:end_ts] + cols + text[end_ts:]
    old_syn = ("<Name>Спецификация</Name>\n" + syn("Спецификация", "Специфікація (норми)", "\t\t\t\t\t"))
    if old_syn in text:
        text = rep_once(text, old_syn, "<Name>Спецификация</Name>\n"
                        + syn("Структура себестоимости (СС)", "Структура собівартості (СС)", "\t\t\t\t\t"), "синоним Спецификация")
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
    print("  РК: реквизиты, ТЧ, колонки, синоним — на месте")


def patch_ss_form():
    rel = "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml"
    path = os.path.join(BB_WT, rel)
    text, _, _ = read_text(path)
    if 'name="ЭталоннаяСпецификация"' in text:
        print("  форма СС: уже есть")
        return
    start = rep_start = text.index('<InputField name="КоличествоДомов"')
    end = text.index("\t\t\t\t</InputField>\n", rep_start) + len("\t\t\t\t</InputField>\n")
    assert '<UsualGroup name="Группа4"' in text[:start], "поле КоличествоДомов не в Группа4"
    n = max(int(x) for x in re.findall(r' id="(\d+)"', text.split("<Attributes>")[0])) + 1
    block = (f'\t\t\t\t<InputField name="ЭталоннаяСпецификация" id="{n}">\n'
             "\t\t\t\t\t<DataPath>Объект.ЭталоннаяСпецификация</DataPath>\n\t\t\t\t\t<ReadOnly>true</ReadOnly>\n"
             f'\t\t\t\t\t<ContextMenu name="ЭталоннаяСпецификацияКонтекстноеМеню" id="{n + 1}"/>\n'
             f'\t\t\t\t\t<ExtendedTooltip name="ЭталоннаяСпецификацияРасширеннаяПодсказка" id="{n + 2}"/>\n'
             "\t\t\t\t</InputField>\n")
    text = text[:end] + block + text[end:]
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
    print(f"  форма СС: + ЭталоннаяСпецификация (id {n}–{n + 2}, только просмотр)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--etap", choices=["meta", "form"], required=True)
    a = ap.parse_args()
    if a.etap == "meta":
        patch_ss()
        patch_rk()
    else:
        patch_ss_form()
```

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/patch_existing.py --etap meta`
Expected: `СС: + ЭталоннаяСпецификация`, `РК: реквизиты, ТЧ, колонки, синоним — на месте`. Повторный запуск: `СС: уже есть`, дифф не меняется.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only existing`
Expected: `existing OK`.
Run: `git diff --stat -- "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости.xml" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml"`
Expected: только вставки (СС ≈ +43 строки; РК ≈ +1100 строк), удалений 0, кроме двух строк синонима `Спецификация`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/patch_existing.py "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости.xml" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml"
git commit -m "feat(BASEBuh): СС.ЭталоннаяСпецификация; РК — СпецификацияДома, ДатаНормы, ТЧ снимка нормы (метаданные §4.7)

РК остаётся непроводимым, модули РК не меняются: реквизиты пусты до П3 (И1).

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 15: Регистрация — `Configuration.xml`, роли (Д2), подсистема

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/register.py`
- Modify: `_Rarzrabotki/BASEBuh/Configuration.xml`
- Modify: `_Rarzrabotki/BASEBuh/Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml`, `_Rarzrabotki/BASEBuh/Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml`, `_Rarzrabotki/BASEBuh/Roles/ЗапускТонкогоКлиента/Ext/Rights.xml`
- Modify: `_Rarzrabotki/BASEBuh/Subsystems/Склад/Subsystems/А_СписанняБудинків.xml`, `…/А_СписанняБудинків/Ext/CommandInterface.xml`
- Test: `test_e1_files.py --only register`, `preflight_config.py <Configuration.xml worktree>`

- [ ] **Step 1: Тест уже написан** (Task 10, раздел `register`).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only register`
Expected: FAIL — `AssertionError: ('Catalog', 'А_СпецификацииДомов')`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/register.py`:

```python
# -*- coding: utf-8 -*-
"""Регистрация объектов П1 в дампе WORKTREE хирургически (role-compile и meta-compile переписывают файлы целиком — не
используются, память role_compile_configuration_xml_gotcha). Повторный запуск ничего не дублирует.
--etap 1a: Configuration.xml (справочник + 6 перечислений), роли (Д2), подсистема А_СписанняБудинків.
--etap 1b: Configuration.xml (общий модуль А_СпецификацииДомовСервер)."""
import argparse
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, read_text, write_text  # noqa: E402
from test_e1_metadata import ROLE_RIGHTS, UK_ENUM  # noqa: E402

NEW_1A = [("Catalog", "А_СпецификацииДомов")] + [("Enum", e) for e in UK_ENUM]
NEW_1B = [("CommonModule", "А_СпецификацииДомовСервер")]
SUB = "Subsystems/Склад/Subsystems/А_СписанняБудинків.xml"
SUB_CI = "Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml"
OBJ = "Catalog.А_СпецификацииДомов"


def save(path, text):
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)


def register_config(objects):
    path = os.path.join(BB_WT, "Configuration.xml")
    text, _, _ = read_text(path)
    added = 0
    for tag, name in objects:
        line = f"\t\t\t<{tag}>{name}</{tag}>\n"
        if line in text:
            continue
        last = list(re.finditer(rf"^\t\t\t<{tag}>[^<]+</{tag}>\n", text, re.M))[-1]      # конец своей группы (E1 §1.1)
        text = text[:last.end()] + line + text[last.end():]
        added += 1
    save(path, text)
    print(f"  Configuration.xml: +{added}")


def rights_block(name, rights):
    lines = ["\t<object>", f"\t\t<name>{name}</name>"]
    for r in rights:
        lines += ["\t\t<right>", f"\t\t\t<name>{r}</name>", "\t\t\t<value>true</value>", "\t\t</right>"]
    return "\n".join(lines + ["\t</object>"]) + "\n"


def register_roles():
    for role, rights in ROLE_RIGHTS.items():
        path = os.path.join(BB_WT, "Roles", role, "Ext", "Rights.xml")
        text, _, _ = read_text(path)
        if f"<name>{OBJ}</name>" in text:
            print(f"  роль {role}: уже есть")
            continue
        anchor = "\t<restrictionTemplate>" if "\t<restrictionTemplate>" in text else "</Rights>"
        i = text.index(anchor)
        text = text[:i] + rights_block(OBJ, rights) + text[i:]
        save(path, text)
        print(f"  роль {role}: + {OBJ} ({', '.join(rights)})")


def register_subsystem():
    path = os.path.join(BB_WT, SUB)
    text, _, _ = read_text(path)
    item = f'\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">{OBJ}</xr:Item>\n'
    anchor = '\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">Catalog.СтруктураСебестоимости</xr:Item>\n'
    if item not in text:
        assert text.count(anchor) == 1, "якорь СС в Content"
        save(path, text.replace(anchor, anchor + item, 1))
    path = os.path.join(BB_WT, SUB_CI)
    text, _, _ = read_text(path)
    cmd = (f'\t\t<Command name="{OBJ}.StandardCommand.OpenList">\n\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
           "\t\t</Command>\n")
    anchor = ('\t\t<Command name="Catalog.СтруктураСебестоимости.StandardCommand.OpenList">\n'
              "\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n\t\t</Command>\n")
    if cmd not in text:
        assert text.count(anchor) == 1, "якорь СС в CommandInterface"
        save(path, text.replace(anchor, anchor + cmd, 1))
    print("  подсистема А_СписанняБудинків: справочник в составе и в панели навигации")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--etap", choices=["1a", "1b"], required=True)
    a = ap.parse_args()
    if a.etap == "1a":
        register_config(NEW_1A)
        register_roles()
        register_subsystem()
    else:
        register_config(NEW_1B)
```

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/register.py --etap 1a`
Expected: `Configuration.xml: +7`, две строки `роль …: + Catalog.А_СпецификацииДомов (…)`, строка подсистемы.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py`
Expected: `enums OK`, `catalog OK`, `existing OK`, `register OK`, `OK test_e1_files all`.
Run: `git diff --numstat -- _Rarzrabotki/BASEBuh/Configuration.xml`
Expected: `7	0` (ровно 7 добавленных строк).
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/preflight_config.py _Rarzrabotki/BASEBuh/Configuration.xml`
Expected: `новых П1 в файле, но не в базе: ['А_СпецификацииДомов', 'А_ВидыСпецификацийДомов', …]` (7 имён), `ПРОВАЛОВ: 0`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/register.py _Rarzrabotki/BASEBuh/Configuration.xml \
  "_Rarzrabotki/BASEBuh/Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml" "_Rarzrabotki/BASEBuh/Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml" "_Rarzrabotki/BASEBuh/Roles/ЗапускТонкогоКлиента/Ext/Rights.xml" \
  "_Rarzrabotki/BASEBuh/Subsystems/Склад/Subsystems/А_СписанняБудинків.xml" "_Rarzrabotki/BASEBuh/Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml"
git commit -m "feat(BASEBuh): регистрация справочника и перечислений П1, права (вариант а), навигация «Будинки: списання за СС»

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 16: Статический гейт `check_static.py`

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/check_static.py`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_check_static.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_check_static.py`:

```python
# -*- coding: utf-8 -*-
"""check_static.py ловит типовые грабли BSL до загрузки и пропускает чистый модуль."""
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, write_text  # noqa: E402

GOOD = """#Область Тест

Функция Проба(Знач Текст) Экспорт
	Запрос = Новый Запрос(
	"ВЫБРАТЬ
	|	Т.Ссылка КАК Ссылка
	|ИЗ
	|	Справочник.Номенклатура КАК Т");
	Если Текст = "" Тогда
		Возврат НСтр("ru = 'Пусто'; uk = 'Порожньо'");
	КонецЕсли;
	Возврат Запрос;
КонецФункции

#КонецОбласти
"""
BAD = (GOOD.replace("КонецФункции", "КонецФункция").replace("КАК Т\"", "КАК ПО\"").replace("uk = 'Порожньо'", "")
       .replace("\tЕсли Текст", "\tДокументы = Новый Массив;\n\tЕсли Текст"))


def run(path):
    return subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "check_static.py"), path],
                          capture_output=True, text=True, encoding="utf-8")


if __name__ == "__main__":
    d = tempfile.mkdtemp()
    good, bad = os.path.join(d, "good.bsl"), os.path.join(d, "bad.bsl")
    write_text(good, GOOD)
    write_text(bad, BAD)
    r = run(good)
    assert r.returncode == 0, r.stdout
    r = run(bad)
    assert r.returncode == 1, r.stdout
    for needle in ("КонецФункция", "КАК ПО", "НСтр без uk", "баланс Функция", "«Документы»"):
        assert needle in r.stdout, (needle, r.stdout)
    print("OK test_check_static")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_check_static.py`
Expected: FAIL — `AssertionError` (скрипта нет, код 2).

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/check_static.py`:

```python
# -*- coding: utf-8 -*-
"""Статический гейт П1 (E2 §4.1 П5).
XML: BOM, CRLF, well-formed, без ' />' и d5p1, у новых объектов П1 ru = uk; формы — уникальные id, <Action> без xsi:type.
BSL: BOM, CRLF, баланс Процедура/Функция/Если/Цикл/Попытка (без строк и комментариев), опечатки КонецФункция/КонецПроцедура,
псевдонимы-ключевые слова в запросах (КАК В/И/ПО…), SQLSERVER, Асинх/Ждать (совместимость 8.3.12), присваивание
зарезервированным именам и именам общих модулей (память feedback_vypolnit_reserved), НСтр без uk, неразвёрнутый //@ЗАПРОС,
вызов метода прямо на Новый; в модулях форм — директива перед каждой процедурой.
--tests: в spec_domov/*.py нет CommitTransaction/ЗафиксироватьТранзакцию/.Удалить().
Запуск: python check_static.py <путь|rel дампа>… | --list <список загрузки> | --tests"""
import argparse
import glob
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE  # noqa: E402

RESERVED = {"Ссылка", "Метаданные", "Выполнить", "Знач", "ЭтотОбъект", "Строка", "Число", "Дата", "Тип", "Булево",
            # свойства глобального контекста: неявная локальная с таким именем падает «Поле объекта недоступно для записи»
            "Справочники", "Документы", "Перечисления", "Константы", "Отчеты", "Обработки", "ПланыВидовХарактеристик",
            "ПланыСчетов", "ПланыОбмена", "РегистрыСведений", "РегистрыНакопления", "РегистрыБухгалтерии",
            "ЖурналыДокументов", "БизнесПроцессы", "Задачи", "ПараметрыСеанса", "Символы", "ВнешниеОбработки", "ВнешниеОтчеты"}
FORM_MEMBERS = {"Параметры", "Элементы", "Объект", "Команды", "УсловноеОформление", "Модифицированность"}
KW = ("В", "И", "ПО", "ИЛИ", "НЕ", "ИЗ", "КАК", "ДЛЯ", "ГДЕ", "ВЫБРАТЬ", "ИМЕЮЩИЕ", "ИТОГИ", "МЕЖДУ", "ЕСТЬ",
      "ВЫБОР", "КОГДА", "ТОГДА", "ИНАЧЕ", "КОНЕЦ", "ПОДОБНО", "СОЕДИНЕНИЕ")
L = r"[A-Za-zА-Яа-яЁёІіЇїЄєҐґ_0-9]"
NEW_MARK = ("А_СпецификацииДомов", "А_ВидыСпецификацийДомов", "А_СтатусыСпецификацийДомов", "А_ВидыСтрокСпецификацииДомов",
            "А_ИсточникиКарточкиСпецификации", "А_ГруппыЗакупочнойСпецификации", "А_ИсточникиЦеныЗакупки")


def module_names():
    return {os.path.splitext(f)[0] for f in os.listdir(os.path.join(BB_MAIN, "CommonModules")) if f.endswith(".xml")}


def strip_code(text):
    """Текст модуля без строковых литералов (→ "") и комментариев; переводы строк сохраняются."""
    out, i, n, in_str = [], 0, len(text), False
    while i < n:
        c = text[i]
        if in_str:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
                out.append('""')
            elif c == "\n":
                out.append("\n")
            i += 1
            continue
        if text.startswith("//", i):
            j = text.find("\n", i)
            i = n if j < 0 else j
            continue
        if c == '"':
            in_str = True
            i += 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def count(pattern, code):
    return len(re.findall(pattern, code, re.M))


def check_raw(raw, errs, label):
    if not raw.startswith(b"\xef\xbb\xbf"):
        errs.append(f"{label}: нет BOM")
    if raw.replace(b"\r\n", b"").count(b"\n"):
        errs.append(f"{label}: не CRLF")


def check_bsl(label, raw, is_form, mods):
    errs = []
    check_raw(raw, errs, label)
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")
    code = strip_code(text)
    pairs = [(r"^\s*Процедура\s", r"^\s*КонецПроцедуры\b", "Процедура"), (r"^\s*Функция\s", r"^\s*КонецФункции\b", "Функция"),
             (r"(?<!Иначе)(?<!" + L + r")Если(?!" + L + r")", r"(?<!" + L + r")КонецЕсли(?!" + L + r")", "Если"),
             (r"(?<!" + L + r")Цикл(?!" + L + r")", r"(?<!" + L + r")КонецЦикла(?!" + L + r")", "Цикл"),
             (r"(?<!" + L + r")Попытка(?!" + L + r")", r"(?<!" + L + r")КонецПопытки(?!" + L + r")", "Попытка")]
    for op, cl, nm in pairs:
        a, b = count(op, code), count(cl, code)
        if a != b:
            errs.append(f"{label}: баланс {nm}: {a} ≠ {b}")
    for m in re.finditer(r"КонецФункция\b|КонецПроцедура\b", code):
        errs.append(f"{label}: опечатка {m.group(0)}")
    for m in re.finditer(r"(?<!" + L + r")(Асинх|Ждать)(?!" + L + r")", code):
        errs.append(f"{label}: {m.group(1)} недоступно в режиме совместимости 8.3.12")
    if "SQLSERVER" in text:
        errs.append(f"{label}: SQLSERVER в коде (только localhost)")
    if "//@ЗАПРОС" in text and not label.startswith("src:"):
        errs.append(f"{label}: неразвёрнутый //@ЗАПРОС")
    for m in re.finditer(r"КАК\s+(" + "|".join(KW) + r")(?!" + L + r")", text):
        errs.append(f"{label}: псевдоним-ключевое слово «КАК {m.group(1)}»")
    targets = [m.group(1) for m in re.finditer(r"^[ \t]*(" + L + r"+)[ \t]*=(?!=)", code, re.M)]
    # переменная цикла — тоже неявная локальная («Для Каждого Документы Из …» падает так же, как присваивание)
    targets += [m.group(1) for m in re.finditer(r"^[ \t]*Для[ \t]+(?:Каждого[ \t]+)?(" + L + r"+)[ \t]*(?:Из(?!" + L + r")|=)",
                                                code, re.M | re.I)]
    for name in targets:
        if name in RESERVED or name in mods or (is_form and name in FORM_MEMBERS):
            errs.append(f"{label}: присваивание зарезервированному имени «{name}»")
    for m in re.finditer(r'НСтр\(\s*"((?:[^"]|"")*)"', text):
        if not re.search(r"\buk\s*=", m.group(1)):
            errs.append(f"{label}: НСтр без uk: {m.group(1)[:50]}")
    if re.search(r"Новый\s+" + L + r"+\([^()\n]*\)\.", code):
        errs.append(f"{label}: вызов метода прямо на Новый")
    if is_form:
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if re.match(r"^(Процедура|Функция)\s", line):
                j = i - 1
                while j >= 0 and not lines[j].strip():
                    j -= 1
                if j < 0 or not lines[j].lstrip().startswith("&"):
                    errs.append(f"{label}: нет директивы компиляции перед «{line.strip()[:60]}»")
    return errs


def check_xml(label, raw, new_obj, is_form):
    errs = []
    check_raw(raw, errs, label)
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")
    try:
        xml.dom.minidom.parseString(text.encode("utf-8"))
    except Exception as e:
        errs.append(f"{label}: не XML: {e}")
    if " />" in text or "d5p1" in text:
        errs.append(f"{label}: шум генератора (' />' или d5p1)")
    if new_obj and text.count("<v8:lang>ru</v8:lang>") != text.count("<v8:lang>uk</v8:lang>"):
        errs.append(f"{label}: ru {text.count('<v8:lang>ru</v8:lang>')} ≠ uk {text.count('<v8:lang>uk</v8:lang>')}")
    if is_form:
        if "<Action xsi:type" in text:
            errs.append(f"{label}: <Action> с xsi:type")
        body = text.split("<Attributes>")[0]
        ids = re.findall(r'<(?!Attribute|Command|Column|Parameter)\w+ name="[^"]*" id="(-?\d+)"', body)
        dup = {x for x in ids if ids.count(x) > 1}
        if dup:
            errs.append(f"{label}: повтор id элементов {sorted(dup)}")
        cids = re.findall(r'<Command name="[^"]*" id="(\d+)"', text)
        if len(cids) != len(set(cids)):
            errs.append(f"{label}: повтор id команд")
    return errs


def resolve(p):
    if os.path.isabs(p):
        return p
    wt = os.path.join(BB_WT, p)
    return wt if os.path.exists(wt) else os.path.join(BB_MAIN, p)


def expand(rels):
    out = []
    for r in rels:
        out.append(r)
        m = re.match(r"(.*/Forms/[^/]+)\.xml$", r)
        if m:
            out += [m.group(1) + "/Ext/Form.xml", m.group(1) + "/Ext/Form/Module.bsl"]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="*")
    ap.add_argument("--list")
    ap.add_argument("--tests", action="store_true")
    a = ap.parse_args()
    errs = []
    if a.tests:
        for f in glob.glob(os.path.join(HERE, "*.py")):
            if os.path.basename(f) == "check_static.py":
                continue
            t = open(f, encoding="utf-8").read()
            for bad in ("CommitTransaction", "ЗафиксироватьТранзакцию", ".Удалить()"):
                if bad in t:
                    errs.append(f"{os.path.basename(f)}: запрещено в тестах: {bad}")
    rels = list(a.paths)
    if a.list:
        rels += [ln.strip() for ln in open(a.list, encoding="utf-8-sig") if ln.strip() and not ln.startswith("#")]
    mods = module_names()
    for rel in expand(rels):
        path = resolve(rel)
        if not os.path.exists(path):
            if rel.endswith("Module.bsl") and "/Forms/" in rel:
                continue
            errs.append(f"{rel}: файла нет")
            continue
        raw = open(path, "rb").read()
        new_obj = any(m in rel for m in NEW_MARK)
        if path.endswith(".bsl"):
            errs += check_bsl(rel, raw, "/Forms/" in rel.replace("\\", "/"), mods)
        elif path.endswith(".xml") and "/Roles/" not in rel.replace("\\", "/") and not rel.endswith("Configuration.xml"):
            errs += check_xml(rel, raw, new_obj, rel.endswith("Form.xml"))
        elif path.endswith(".xml"):
            try:
                xml.dom.minidom.parseString(raw.decode("utf-8-sig").encode("utf-8"))
            except Exception as e:
                errs.append(f"{rel}: не XML: {e}")
    for e in errs:
        print("ОШИБКА:", e)
    print(f"проверено файлов: {len(expand(rels))}, ошибок: {len(errs)}")
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
```

Создать также `_Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt` (если не создан в Task 5) — 16 строк из раздела «Списки Partial-загрузки».

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_check_static.py`
Expected: `OK test_check_static`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt --tests`
Expected: `ошибок: 0`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/test_check_static.py _Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt
git commit -m "test(spec_domov): статический гейт XML/BSL до загрузки (баланс блоков, резервные имена, псевдонимы, uk)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 17: Загрузка 1A (метаданные) и приёмка метаданных

**Files:**
- Modify (копии worktree → основной): все файлы `load_lists/e1a.txt`, которые есть в worktree
- Create: `_Rarzrabotki/Python/test/spec_domov/out/e1a_load_<YYYYMMDD_HHMM>.txt`

- [ ] **Step 1: Тест** — `test_e1_metadata.py` (Task 9) и `dump_check.py`.

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_metadata.py`
Expected: FAIL — `перечисление А_ВидыСпецификацийДомов не найдено в базе`.

- [ ] **Step 3: Копия в основной каталог и загрузка**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt`
Expected: `скопирован …` для 6 перечислений, справочника, СС, РК, двух `Rights.xml`, двух файлов подсистемы и `Configuration.xml`; `ОТКАЗ` — нет. При `ОТКАЗ: … основной каталог изменён после синхронизации` — `sync_main_to_wt.py --file <rel>`, повторить `patch_existing.py --etap meta` / `register.py --etap 1a`, затем снова `to_main.py`.

Run: `powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e1a.txt -Tag e1a -WithRoles`
Expected: все шаги с кодом 0, последняя строка `OK load e1a. Logs: …\load_e1a`. Остановка на `db-update -Dynamic +` с «монопольный режим»/блокировкой — по правилу раздела «Как исполнять» (сообщить пользователю, ждать F7).

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_metadata.py`
Expected: `ПРОВАЛОВ: 0`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/dump_check.py --tag e1a --canon --objects "Справочник.А_СпецификацииДомов,Перечисление.А_ВидыСпецификацийДомов,Перечисление.А_СтатусыСпецификацийДомов,Перечисление.А_ВидыСтрокСпецификацииДомов,Перечисление.А_ИсточникиКарточкиСпецификации,Перечисление.А_ГруппыЗакупочнойСпецификации,Перечисление.А_ИсточникиЦеныЗакупки,Справочник.СтруктураСебестоимости,Документ.РасчетКомплектаций"`
Expected: `различие по содержимому 0`; файлы «только форматирование» и новые файлы П1, которые платформа сериализовала иначе («ПРИНЯТО КАНОНОМ»), переписаны выгруженной версией в обоих каталогах. Различие по содержимому у СС или РК — остановиться: загрузка не дошла или кеш Конфигуратора (память `feedback_designer_cache_invalidation`).
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/preflight_config.py`
Expected: `новых П1 уже в базе: [7 имён]`, `ПРОВАЛОВ: 0`.
Сводку шагов загрузки `load_stage.ps1` пишет сам в `out/e1a_load_<stamp>.txt`; итоговый вывод проверок метаданных и гейта дописать туда же (строку итога `dump_check.py` — вручную из вывода выше):
```bash
F=$(ls -t _Rarzrabotki/Python/test/spec_domov/out/e1a_load_*.txt | head -1)
for S in test_e1_metadata.py preflight_config.py; do python -X utf8 _Rarzrabotki/Python/test/spec_domov/$S 2>&1 | tee -a "$F"; done
```

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/out/e1a_load_*.txt _Rarzrabotki/Python/test/spec_domov/sync_manifest.json \
  "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml" \
  "_Rarzrabotki/BASEBuh/Enums/А_ВидыСпецификацийДомов.xml" "_Rarzrabotki/BASEBuh/Enums/А_СтатусыСпецификацийДомов.xml" \
  "_Rarzrabotki/BASEBuh/Enums/А_ВидыСтрокСпецификацииДомов.xml" "_Rarzrabotki/BASEBuh/Enums/А_ИсточникиКарточкиСпецификации.xml" \
  "_Rarzrabotki/BASEBuh/Enums/А_ГруппыЗакупочнойСпецификации.xml" "_Rarzrabotki/BASEBuh/Enums/А_ИсточникиЦеныЗакупки.xml" \
  "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости.xml" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml"

git commit -m "feat(BASEBuh): загрузка 1A в bas_industrialbud — справочник, перечисления, реквизиты СС/РК, права, навигация

Partial из основного каталога, гейты 0, CheckModules чисто, db-update -Dynamic +; канон — выгрузка платформы.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Фаза 1B — код, формы и приёмка этапа 1 (загрузка 1B)

### Task 18: Rule #-1 — запросы к новому справочнику (Python COM до BSL)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/zaprosy/uroven_dereva.sql`, `roditeli.sql`, `doma_v_variantah.sql`, `drugie_zakupochnye.sql`, `specifikacii_dlya_proverki.sql`, `spisok_formy.sql`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e1_zaprosy.py`

Справочник уже в базе (Task 17), поэтому запросы к нему исполняются впервые здесь (E3 §5, §8 п. 10). Тестовые данные пишутся «тихо» (`ОбменДанными.Загрузка`) внутри транзакции и откатываются.

- [ ] **Step 1: Написать тест (тексты запросов ещё не созданы — тест падает)**

`_Rarzrabotki/Python/test/spec_domov/test_e1_zaprosy.py`:

```python
# -*- coding: utf-8 -*-
"""Rule #-1 для запросов этапа 1: каждый файл zaprosy/*.sql исполняется на bas_industrialbud до переноса в BSL.
Данные — тестовые элементы внутри транзакции (откат). Печатает «OK, строк=N» по каждому запросу."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, q, same, tx, zapros  # noqa: E402

PFX = "__ТЕСТ_П1 "


def spec(B, kind, name, ss_ref=None, st="Черновик", unit_ref=None):
    o = B.Справочники.А_СпецификацииДомов.СоздатьЭлемент()
    o.Наименование = PFX + name
    o.ВидСпецификации = getattr(B.Перечисления.А_ВидыСпецификацийДомов, kind)
    o.Статус = getattr(B.Перечисления.А_СтатусыСпецификацийДомов, st)
    o.КоличествоВыпуска = 1
    if ss_ref is not None:
        o.СтруктураСебестоимости = ss_ref
    if unit_ref is not None:
        o.ЕдиницаВыпуска = unit_ref
    return o


def node_row(B, o, node, unit_ref):
    r = o.Состав.Добавить()
    r.КлючСтроки = B.NewObject("УникальныйИдентификатор")
    r.ВидСтроки = B.Перечисления.А_ВидыСтрокСпецификацииДомов.Узел
    r.Узел = node
    r.Количество = 1
    r.ЕдиницаСтроки = unit_ref
    r.Коэффициент = 1


def silent(o):
    o.ОбменДанными.Загрузка = True
    o.Записать()
    return o.Ссылка


def run(B, name, **params):
    rows = q(B, zapros(name), **params)
    print(f"OK {name}, строк={len(rows)}")
    return rows


def main():
    B = connect()
    s9 = B.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000009")
    u = B.Справочники.КлассификаторЕдиницИзмерения.НайтиПоКоду("2938")
    home = q(B, "ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка КАК Ссылка ИЗ Справочник.Подразделения КАК П ГДЕ П.Родитель = &Р УПОРЯДОЧИТЬ ПО П.Код",
             Р=s9.ПодразделениеОбъекта)[0]["Ссылка"]
    with tx(B):
        septik = silent(spec(B, "Узел", "Септик", unit_ref=u))
        zm_o = spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        node_row(B, zm_o, septik, u)
        zm = silent(zm_o)
        h_o = spec(B, "Дом", "дім", s9, "Действует")
        node_row(B, h_o, zm, u)
        h_o.Дома.Добавить().Подразделение = home
        h = silent(h_o)
        v_o = spec(B, "Дом", "варіант", s9, "Действует")
        v_o.Дома.Добавить().Подразделение = home
        v = silent(v_o)
        z1 = silent(spec(B, "Закупочная", "закупівельна діє", s9, "Действует"))
        z2 = silent(spec(B, "Закупочная", "закупівельна чернетка", s9))

        z = B.NewObject("Запрос")
        z.Text = zapros("uroven_dereva")
        z.SetParameter("СпецификацииУровня", arr(B, [h, zm]))
        pak = z.ВыполнитьПакет()
        heads, rows = pak.Получить(0).Выгрузить(), pak.Получить(1).Выгрузить()
        assert heads.Количество() == 2 and rows.Количество() == 2, (heads.Количество(), rows.Количество())
        print("OK uroven_dereva, шапок=2, строк=2")

        par = run(B, "roditeli", Узлы=arr(B, [septik]))
        assert len(par) == 1 and same(B, par[0]["Спецификация"], zm)
        par = run(B, "roditeli", Узлы=arr(B, [zm]))
        assert len(par) == 1 and same(B, par[0]["Спецификация"], h)

        dv = run(B, "doma_v_variantah", Дома=arr(B, [home]), Ссылка=h, СсылкаСС=s9)
        assert len(dv) == 1 and same(B, dv[0]["Спецификация"], v)

        dz = run(B, "drugie_zakupochnye", Ссылка=z2, СсылкаСС=s9)
        assert len(dz) == 1 and same(B, dz[0]["Ссылка"], z1)

        sp = run(B, "specifikacii_dlya_proverki")
        names = {str(r["Наименование"]) for r in sp}
        assert {PFX + "Септик", PFX + "Зовнішні мережі", PFX + "дім"} <= names and PFX + "закупівельна діє" not in names

        sso = s9.ПолучитьОбъект()
        sso.ЭталоннаяСпецификация = h
        sso.ОбменДанными.Загрузка = True
        sso.ДополнительныеСвойства.Вставить("ОтключитьМеханизмРегистрацииОбъектов", True)
        sso.Записать()
        lst = [r for r in run(B, "spisok_formy") if same(B, r["Ссылка"], h)]
        assert len(lst) == 1 and same(B, lst[0]["ЭталонДляСС"], s9), "колонка «Еталон для СС»"
    print("OK test_e1_zaprosy")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_zaprosy.py`
Expected: FAIL — `FileNotFoundError: …zaprosy\uroven_dereva.sql`.

- [ ] **Step 3: Написать тексты запросов** (псевдонимы — не ключевые слова: `Спец`, `СоставСпец`, `ДомаСпец`, `ЭтСС`, `Эталоны`)

`zaprosy/uroven_dereva.sql`:

```
ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка,
	Спец.Код КАК Код,
	Спец.Наименование КАК Наименование,
	Спец.ВидСпецификации КАК ВидСпецификации,
	Спец.Статус КАК Статус,
	Спец.КоличествоВыпуска КАК КоличествоВыпуска,
	Спец.ЕдиницаВыпуска КАК ЕдиницаВыпуска,
	Спец.НоменклатураВузла КАК НоменклатураВузла,
	Спец.СтруктураСебестоимости КАК СтруктураСебестоимости,
	Спец.ДатаИзменения КАК ДатаИзменения
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
ГДЕ
	Спец.Ссылка В (&СпецификацииУровня)
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	СоставСпец.Ссылка КАК Спецификация,
	СоставСпец.НомерСтроки КАК НомерСтроки,
	СоставСпец.КлючСтроки КАК КлючСтроки,
	СоставСпец.ВидСтроки КАК ВидСтроки,
	СоставСпец.ЭтапСС КАК ЭтапСС,
	СоставСпец.ПозицияСС КАК ПозицияСС,
	СоставСпец.Номенклатура КАК Номенклатура,
	СоставСпец.ИсточникКарточки КАК ИсточникКарточки,
	СоставСпец.Проверено КАК Проверено,
	СоставСпец.НеПоставляется КАК НеПоставляется,
	СоставСпец.ОбщееНазвание КАК ОбщееНазвание,
	СоставСпец.Узел КАК Узел,
	СоставСпец.ПокупаетсяГотовым КАК ПокупаетсяГотовым,
	СоставСпец.Количество КАК Количество,
	СоставСпец.ЕдиницаСтроки КАК ЕдиницаСтроки,
	СоставСпец.Коэффициент КАК Коэффициент,
	СоставСпец.ИсточникКоэффициента КАК ИсточникКоэффициента,
	СоставСпец.ЦенаНормы КАК ЦенаНормы,
	СоставСпец.Комментарий КАК Комментарий
ИЗ
	Справочник.А_СпецификацииДомов.Состав КАК СоставСпец
ГДЕ
	СоставСпец.Ссылка В (&СпецификацииУровня)

УПОРЯДОЧИТЬ ПО
	Спецификация,
	НомерСтроки
```

`zaprosy/roditeli.sql`:

```
ВЫБРАТЬ РАЗЛИЧНЫЕ
	СоставСпец.Ссылка КАК Спецификация,
	СоставСпец.Узел КАК ЧерезУзел,
	СоставСпец.Ссылка.ВидСпецификации КАК ВидСпецификации,
	СоставСпец.Ссылка.СтруктураСебестоимости КАК СтруктураСебестоимости,
	СоставСпец.Ссылка.Статус КАК Статус
ИЗ
	Справочник.А_СпецификацииДомов.Состав КАК СоставСпец
ГДЕ
	СоставСпец.Узел В (&Узлы)
	И НЕ СоставСпец.Ссылка.ПометкаУдаления
```

`zaprosy/doma_v_variantah.sql`:

```
ВЫБРАТЬ
	ДомаСпец.Подразделение КАК Подразделение,
	ДомаСпец.Ссылка КАК Спецификация
ИЗ
	Справочник.А_СпецификацииДомов.Дома КАК ДомаСпец
ГДЕ
	ДомаСпец.Подразделение В (&Дома)
	И ДомаСпец.Ссылка <> &Ссылка
	И ДомаСпец.Ссылка.СтруктураСебестоимости = &СсылкаСС
	И ДомаСпец.Ссылка.Статус = ЗНАЧЕНИЕ(Перечисление.А_СтатусыСпецификацийДомов.Действует)
	И НЕ ДомаСпец.Ссылка.ПометкаУдаления
```

`zaprosy/drugie_zakupochnye.sql`:

```
ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
ГДЕ
	Спец.ВидСпецификации = ЗНАЧЕНИЕ(Перечисление.А_ВидыСпецификацийДомов.Закупочная)
	И Спец.СтруктураСебестоимости = &СсылкаСС
	И Спец.Статус = ЗНАЧЕНИЕ(Перечисление.А_СтатусыСпецификацийДомов.Действует)
	И Спец.Ссылка <> &Ссылка
	И НЕ Спец.ПометкаУдаления
```

`zaprosy/specifikacii_dlya_proverki.sql`:

```
ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка,
	Спец.Наименование КАК Наименование
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
ГДЕ
	НЕ Спец.ЭтоГруппа
	И НЕ Спец.ПометкаУдаления
	И Спец.ВидСпецификации В (ЗНАЧЕНИЕ(Перечисление.А_ВидыСпецификацийДомов.Дом), ЗНАЧЕНИЕ(Перечисление.А_ВидыСпецификацийДомов.Узел))
	И Спец.Статус <> ЗНАЧЕНИЕ(Перечисление.А_СтатусыСпецификацийДомов.Закрыта)
```

`zaprosy/spisok_formy.sql` (динамический список формы списка, колонка «Еталон для СС»):

```
ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка,
	Спец.ПометкаУдаления КАК ПометкаУдаления,
	Спец.Код КАК Код,
	Спец.Наименование КАК Наименование,
	Спец.ВидСпецификации КАК ВидСпецификации,
	Спец.Статус КАК Статус,
	Спец.СтруктураСебестоимости КАК СтруктураСебестоимости,
	Спец.ДатаИзменения КАК ДатаИзменения,
	Эталоны.ТипБудинкуСС КАК ЭталонДляСС
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
		ЛЕВОЕ СОЕДИНЕНИЕ (ВЫБРАТЬ
			ЭтСС.ЭталоннаяСпецификация КАК Спецификация,
			МАКСИМУМ(ЭтСС.Ссылка) КАК ТипБудинкуСС
		ИЗ
			Справочник.СтруктураСебестоимости КАК ЭтСС
		ГДЕ
			ЭтСС.ЭталоннаяСпецификация <> ЗНАЧЕНИЕ(Справочник.А_СпецификацииДомов.ПустаяСсылка)
		СГРУППИРОВАТЬ ПО
			ЭтСС.ЭталоннаяСпецификация) КАК Эталоны
		ПО Спец.Ссылка = Эталоны.Спецификация
```

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_zaprosy.py`
Expected: `OK uroven_dereva, шапок=2, строк=2`, `OK roditeli, строк=1` (дважды), `OK doma_v_variantah, строк=1`, `OK drugie_zakupochnye, строк=1`, `OK specifikacii_dlya_proverki, строк=N`, `OK spisok_formy, строк=N`, `OK test_e1_zaprosy`. Ошибка текста запроса — исправить `.sql` и повторить (BSL ещё не написан).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/zaprosy/uroven_dereva.sql _Rarzrabotki/Python/test/spec_domov/zaprosy/roditeli.sql \
  _Rarzrabotki/Python/test/spec_domov/zaprosy/doma_v_variantah.sql _Rarzrabotki/Python/test/spec_domov/zaprosy/drugie_zakupochnye.sql \
  _Rarzrabotki/Python/test/spec_domov/zaprosy/specifikacii_dlya_proverki.sql _Rarzrabotki/Python/test/spec_domov/zaprosy/spisok_formy.sql \
  _Rarzrabotki/Python/test/spec_domov/test_e1_zaprosy.py
git commit -m "test(spec_domov): Rule #-1 — запросы этапа 1 к А_СпецификацииДомов проверены Python COM

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 19: Фикстуры и приёмочный тест этапа 1 (падает до 1B)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/fixtures.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py`

- [ ] **Step 1: Написать фикстуры и падающий тест**

`_Rarzrabotki/Python/test/spec_domov/fixtures.py`:

```python
# -*- coding: utf-8 -*-
"""Фикстуры приёмки П1. Тестовые элементы А_СпецификацииДомов создаются ТОЛЬКО внутри tx(B) и откатываются —
эталоны и закупочные в рабочую базу не пишутся. Имена начинаются с «__ТЕСТ_П1 ».
Коды единиц: 2938 компл, 2009 шт, 0101 м, 0123 м2, 5 рул (E3 §4.1)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import err, messages, q  # noqa: E402

PFX = "__ТЕСТ_П1 "


def enum(B, name, value):
    return getattr(getattr(B.Перечисления, name), value)


def unit(B, code):
    ref = B.Справочники.КлассификаторЕдиницИзмерения.НайтиПоКоду(code)
    assert not ref.Пустая(), f"единица {code}"
    return ref


def card(B, code):
    ref = B.Справочники.Номенклатура.НайтиПоКоду(code)
    assert not ref.Пустая(), f"карточка {code}"
    return ref


def ss(B, code):
    ref = B.Справочники.СтруктураСебестоимости.НайтиПоКоду(code)
    assert not ref.Пустая(), f"СС {code}"
    return ref


def ss_row(B, ss_code, n):
    rows = q(B, """ВЫБРАТЬ
	Компл.Этап КАК Этап, Компл.НоменклатураСС КАК НоменклатураСС, Компл.Количество КАК Количество,
	Компл.ЕдиницаСС КАК ЕдиницаСС, Компл.ОбщееНазвание КАК ОбщееНазвание, Компл.Номенклатура КАК Номенклатура,
	Компл.Сумма КАК Сумма, Компл.НомерСтроки КАК НомерСтроки
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код
	И Компл.НомерСтроки = &Н""", Код=ss_code, Н=n)
    assert len(rows) == 1, (ss_code, n)
    return rows[0]


def ss_first_row_with_card(B, ss_code):
    rows = q(B, """ВЫБРАТЬ ПЕРВЫЕ 1
	Компл.НомерСтроки КАК Н
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код
	И Компл.Количество > 0
	И Компл.Номенклатура <> ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка)
УПОРЯДОЧИТЬ ПО
	Компл.НомерСтроки""", Код=ss_code)
    return ss_row(B, ss_code, int(rows[0]["Н"]))


def house(B, ss_code, n=1):
    rows = q(B, """ВЫБРАТЬ Подр.Ссылка КАК Ссылка ИЗ Справочник.Подразделения КАК Подр
ГДЕ Подр.Родитель = &Родитель И НЕ Подр.ПометкаУдаления УПОРЯДОЧИТЬ ПО Подр.Код""", Родитель=ss(B, ss_code).ПодразделениеОбъекта)
    return rows[n - 1]["Ссылка"]


def new_spec(B, kind, name, ss_ref=None, base_qty=1, unit_ref=None, st="Черновик"):
    obj = B.Справочники.А_СпецификацииДомов.СоздатьЭлемент()
    obj.Наименование = PFX + name
    obj.ВидСпецификации = enum(B, "А_ВидыСпецификацийДомов", kind)
    obj.Статус = enum(B, "А_СтатусыСпецификацийДомов", st)
    obj.КоличествоВыпуска = base_qty
    if ss_ref is not None:
        obj.СтруктураСебестоимости = ss_ref
    if unit_ref is not None:
        obj.ЕдиницаВыпуска = unit_ref
    return obj


def add_row(B, obj, kind="Деталь", card_ref=None, qty=1, unit_ref=None, node=None, stage=None, pos="", checked=False,
            ne_post=False, comment="", k=1, ksrc="ЕдиницыСовпадают", on=None, price=0, ready=False):
    r = obj.Состав.Добавить()
    r.ВидСтроки = enum(B, "А_ВидыСтрокСпецификацииДомов", kind)
    if card_ref is not None:
        r.Номенклатура = card_ref
        r.ОбщееНазвание = card_ref.ОбщееНазваниеНоменклатуры
    if on is not None:
        r.ОбщееНазвание = on
    if node is not None:
        r.Узел = node
    if stage is not None:
        r.ЭтапСС = stage
    r.ПозицияСС = pos
    r.Количество = qty
    if unit_ref is not None:
        r.ЕдиницаСтроки = unit_ref
    elif card_ref is not None:
        r.ЕдиницаСтроки = card_ref.БазоваяЕдиницаИзмерения
    r.Коэффициент = k
    r.ИсточникКоэффициента = enum(B, "А_ИсточникиКоэффициентаПересчетаСС", ksrc)
    r.Проверено = checked
    r.НеПоставляется = ne_post
    r.Комментарий = comment
    r.ЦенаНормы = price
    r.ПокупаетсяГотовым = ready
    return r


def write(obj, silent=False):
    if silent:
        obj.ОбменДанными.Загрузка = True
    obj.Записать()
    return obj.Ссылка


def expect_write_error(B, obj, *needles):
    """Запись должна упасть, в ошибке и сообщениях — все needles. После падения транзакция ошибочна: кейс завершать."""
    messages(B)
    try:
        obj.Записать()
    except Exception as e:
        text = err(e) + "\n" + "\n".join(messages(B))
        for n in needles:
            assert n in text, f"нет «{n}» в:\n{text}"
        return text
    raise AssertionError("запись прошла, а должна была упасть: " + ", ".join(needles))


def kn(B, number):
    rows = q(B, """ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка ИЗ Документ.КомплектацияНоменклатуры КАК Док
ГДЕ Док.Номер = &Номер И Док.Дата >= ДАТАВРЕМЯ(2026, 1, 1)""", Номер=number)
    assert len(rows) == 1, number
    return rows[0]["Ссылка"]


def srv(B):
    return B.А_СпецификацииДомовСервер


def row_index_by_ss(B, obj, ss_code, n):
    """Индекс (с 0) строки «Состав», у которой позиция = n-я строка СС (по ключу этап + текст)."""
    row = ss_row(B, ss_code, n)
    m = B.А_СтруктураСебестоимостиСервер
    key = m.КлючПозиции(row["Этап"], row["НоменклатураСС"])
    for i in range(obj.Состав.Количество()):
        r = obj.Состав.Получить(i)
        if r.ПозицияСС and m.КлючПозиции(r.ЭтапСС, r.ПозицияСС) == key:
            return i
    raise AssertionError(f"нет строки с позицией {ss_code}:{n}")


def col(vt, name):
    return [getattr(vt.Получить(i), name) for i in range(vt.Количество())]


def arr_list(a):
    return [a.Получить(i) for i in range(a.Количество())]
```

`_Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py`:

```python
# -*- coding: utf-8 -*-
"""Приёмка этапа 1 (§10.1) — Python COM, каждый негативный кейс в своей транзакции (откат).
Кейсы: k0 запись в транзакции откатывается; k1 цикл не записывается; k2 цикл тихой записью ловится «Перевірити цикли» и при «Діє»;
k3 9 уровней узлов пишутся и разворачиваются; k4 10-й уровень не записывается; k5 10-й уровень тихой записью — исключение
разузлования; k6 Б1 — эталон на СС 003: 214/214 карточек, регистрация к обмену не тронута; k7 эталон только из «Діє»;
k8 Т9 — деталь без карточки; k9 закупочная без проверенных строк не в «Діє»; k10 обмен не тронут за всю приёмку."""
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import OUT, call, changes_rows, connect, count_rows, q, same, tx  # noqa: E402
from fixtures import (PFX, add_row, arr_list, card, expect_write_error, house, new_spec, srv, ss,  # noqa: E402
                      ss_first_row_with_card, ss_row, unit, write)

RES = {}


def k0_tx(B):
    with tx(B):
        write(new_spec(B, "Узел", "откат", unit_ref=unit(B, "2938")))
        assert len(q(B, "ВЫБРАТЬ Т.Ссылка КАК С ИЗ Справочник.А_СпецификацииДомов КАК Т ГДЕ Т.Наименование = &Н", Н=PFX + "откат")) == 1
    assert len(q(B, "ВЫБРАТЬ Т.Ссылка КАК С ИЗ Справочник.А_СпецификацииДомов КАК Т ГДЕ Т.Наименование = &Н", Н=PFX + "откат")) == 0


def k1_cycle_write(B):
    with tx(B):
        u = unit(B, "2938")
        septik = new_spec(B, "Узел", "Септик", unit_ref=u)
        s_ref = write(septik)
        zm = new_spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        add_row(B, zm, "Узел", node=s_ref, unit_ref=u)
        zm_ref = write(zm)
        add_row(B, septik, "Узел", node=zm_ref, unit_ref=u)
        expect_write_error(B, septik, "входить сам у себе", f"{PFX}Септик → {PFX}Зовнішні мережі → {PFX}Септик")


def k2_cycle_dije(B):
    with tx(B):
        u = unit(B, "2938")
        septik = new_spec(B, "Узел", "Септик", unit_ref=u)
        s_ref = write(septik)
        zm = new_spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        add_row(B, zm, "Узел", node=s_ref, unit_ref=u, checked=True)
        zm_ref = write(zm)
        add_row(B, septik, "Узел", node=zm_ref, unit_ref=u, checked=True)
        write(septik, silent=True)                       # «два пользователя»: тихая запись обходит ПередЗаписью
        errs = [str(x) for x in arr_list(srv(B).ПроверитьЦиклы())]
        assert any("входить сам у себе" in e for e in errs), errs
        h = new_spec(B, "Дом", "дім", ss_ref=ss(B, "000000009"), st="Действует")
        add_row(B, h, "Узел", node=zm_ref, unit_ref=u, checked=True)
        expect_write_error(B, h, "входить сам у себе")


def chain9(B):
    """Дом → У1 → … → У9 (узлы снизу вверх); возвращает (дом, объекты узлов по номеру)."""
    u = unit(B, "2938")
    objs, prev = {}, None
    for i in range(9, 0, -1):
        n = new_spec(B, "Узел", f"У{i}", unit_ref=u)
        if prev is not None:
            add_row(B, n, "Узел", node=prev, unit_ref=u)
        prev = write(n)
        objs[i] = n
    h = new_spec(B, "Дом", "глибина 9", ss_ref=ss(B, "000000009"))
    add_row(B, h, "Узел", node=prev, unit_ref=u)
    return write(h), objs


def k3_depth9(B):
    with tx(B):
        h, _ = chain9(B)
        r = srv(B).РазвернутьДоКарточек(h)
        levels = [r.Экземпляры.Получить(i).Уровень for i in range(r.Экземпляры.Количество())]
        assert len(levels) == 9 and max(levels) == 9, levels


def k4_depth10_write(B):
    with tx(B):
        h, objs = chain9(B)
        u10 = write(new_spec(B, "Узел", "У10", unit_ref=unit(B, "2938")))
        add_row(B, objs[9], "Узел", node=u10, unit_ref=unit(B, "2938"))
        expect_write_error(B, objs[9], "Глибина вкладеності")


def k5_depth10_unfold(B):
    with tx(B):
        h, objs = chain9(B)
        u10 = write(new_spec(B, "Узел", "У10", unit_ref=unit(B, "2938")))
        add_row(B, objs[9], "Узел", node=u10, unit_ref=unit(B, "2938"))
        write(objs[9], silent=True)
        res, e = call(srv(B).РазвернутьДоКарточек, h)
        assert res is None and "Глибина вкладеності" in e, e


def etalon_ss003(B, st):
    s3 = ss(B, "000000003")
    row = ss_first_row_with_card(B, "000000003")
    h = new_spec(B, "Дом", "еталон СС 003", ss_ref=s3, st=st)
    add_row(B, h, "Деталь", card_ref=row["Номенклатура"], qty=row["Количество"], stage=row["Этап"],
            pos=row["НоменклатураСС"], checked=True)
    return s3, write(h)


def k6_b1(B):
    cards_q = """ВЫБРАТЬ
	КОЛИЧЕСТВО(Компл.НомерСтроки) КАК Строк,
	СУММА(ВЫБОР КОГДА Компл.Номенклатура = ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка) ТОГДА 0 ИНАЧЕ 1 КОНЕЦ) КАК СКарточкой
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код"""
    with tx(B):
        s3 = ss(B, "000000003")
        before_chg, before_nom = changes_rows(B, s3), count_rows(B, "Справочник.Номенклатура.Изменения")
        _, h = etalon_ss003(B, "Действует")
        assert srv(B).СделатьЭталономСС(h)
        c = q(B, cards_q, Код="000000003")[0]
        assert (int(c["Строк"]), int(c["СКарточкой"])) == (214, 214), c
        assert changes_rows(B, s3) == before_chg, (before_chg, changes_rows(B, s3))
        assert count_rows(B, "Справочник.Номенклатура.Изменения") == before_nom
        e = q(B, "ВЫБРАТЬ СС.ЭталоннаяСпецификация КАК Э ИЗ Справочник.СтруктураСебестоимости КАК СС ГДЕ СС.Код = &Код",
              Код="000000003")[0]["Э"]
        assert same(B, e, h)
        RES["b1"] = {"cards": [214, 214], "changes": before_chg}


def k7_etalon_only_dije(B):
    with tx(B):
        _, h = etalon_ss003(B, "Черновик")
        res, e = call(srv(B).СделатьЭталономСС, h)
        assert res is None and "переведіть специфікацію в «Діє»" in e, e


def k8_t9(B):
    s9 = ss(B, "000000009")
    row = ss_row(B, "000000009", 2)
    common = dict(stage=row["Этап"], pos=row["НоменклатураСС"], qty=row["Количество"])
    with tx(B):
        h = new_spec(B, "Дом", "чернетка без картки", ss_ref=s9)
        add_row(B, h, ksrc="НеНайден", **common)
        write(h)
    with tx(B):
        h = new_spec(B, "Дом", "діє без картки", ss_ref=s9, st="Действует")
        add_row(B, h, checked=True, ksrc="НеНайден", **common)
        expect_write_error(B, h, "деталь без картки")
    with tx(B):
        h = new_spec(B, "Дом", "нп без причини", ss_ref=s9)
        add_row(B, h, ne_post=True, ksrc="НеНайден", **common)
        expect_write_error(B, h, "потрібна причина")
    with tx(B):
        h = new_spec(B, "Дом", "нп з причиною", ss_ref=s9, st="Действует")
        add_row(B, h, ne_post=True, comment="тест: не постачається", ksrc="НеНайден", **common)
        write(h)


def k9_zakupochna(B):
    """Закупочная: без проверенных строк не в «Діє»; дом факта другого типа (15 м в закупочной 30 м) — отказ (§4.5);
    проверенная строка и дом своего типа — записывается."""
    s9 = ss(B, "000000009")
    row = ss_row(B, "000000009", 42)
    for checked, home_ss, needle in ((False, "000000009", "не перевірено"), (True, "000000008", "не будинок цього типу"),
                                     (True, "000000009", None)):
        with tx(B):
            et = new_spec(B, "Дом", "еталон", ss_ref=s9, st="Действует")
            add_row(B, et, card_ref=card(B, "00-00006711"), qty=1, stage=row["Этап"], pos=row["НоменклатураСС"], checked=True)
            et_ref = write(et)
            z = new_spec(B, "Закупочная", "закупівельна", ss_ref=s9, st="Действует")
            z.Основание = et_ref
            z.ДомаФакта.Добавить().Подразделение = house(B, home_ss, 1)
            zr = z.ЗакупочныйСостав.Добавить()
            zr.Номенклатура = card(B, "00-00006711")
            zr.Количество = 1
            zr.Проверено = checked
            if needle is None:
                write(z)
            else:
                expect_write_error(B, z, needle)


EXCHANGE_Q = """ВЫБРАТЬ
	Изм.Ссылка.Код КАК Код,
	ПРЕДСТАВЛЕНИЕ(Изм.Узел) КАК Узел,
	ЕСТЬNULL(Изм.НомерСообщения, -1) КАК НомерСообщения
ИЗ
	Справочник.СтруктураСебестоимости.Изменения КАК Изм
УПОРЯДОЧИТЬ ПО
	Код"""
START = {}


def exchange_snapshot(B):
    return [{"code": r["Код"], "node": r["Узел"], "msg": int(r["НомерСообщения"])} for r in q(B, EXCHANGE_Q)]


def k10_exchange_untouched(B):
    """Регистрация СС к обмену в конце прогона = в его начале (START заполняет main). Инвариант прогона, а не снимок
    этапа 0: между днями таблицу меняют обмен и правки пользователей (память feedback_invariantnye_testy_ne_baseline).
    Расхождение — сперва проверить, не правил ли кто-то СС во время прогона (код и узел — в тексте ошибки)."""
    now = exchange_snapshot(B)
    assert now == START["exchange"], ([x for x in now if x not in START["exchange"]],
                                      [x for x in START["exchange"] if x not in now])


CASES = [k0_tx, k1_cycle_write, k2_cycle_dije, k3_depth9, k4_depth10_write, k5_depth10_unfold, k6_b1,
         k7_etalon_only_dije, k8_t9, k9_zakupochna, k10_exchange_untouched]


def main():
    B = connect()
    START["exchange"] = exchange_snapshot(B)
    fails = {}
    for fn in CASES:
        try:
            fn(B)
            RES[fn.__name__] = "OK"
            print("OK  ", fn.__name__)
        except Exception as e:                         # noqa: BLE001 — тест собирает все провалы
            fails[fn.__name__] = repr(e)[:500]
            print("FAIL", fn.__name__, repr(e)[:500])
        assert not B.TransactionActive()
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e1_priemka_{datetime.datetime.now():%Y%m%d_%H%M}.json")
    json.dump({"ok": RES, "fail": fails}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("->", out, "ПРОВАЛОВ:", len(fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py`
Expected: FAIL — `k0_tx` проходит (справочник есть), остальные — `AttributeError` (модуля `А_СпецификацииДомовСервер` нет) или «запись прошла, а должна была упасть» (нет `ПередЗаписью`); `ПРОВАЛОВ: 9`.

- [ ] **Step 3: Реализация** — задачи 20–27. **Step 4: PASS** — Task 28.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/fixtures.py _Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py
git commit -m "test(spec_domov): приёмка этапа 1 — цикл, глубина Т6, Б1 на СС 003, Т9, закупочная (падает до 1B)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 20: Сборщик модулей `build_bsl.py`, общее и единицы (часть 1 общего модуля)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/build_bsl.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_10_obschee.bsl`
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_15_edinicy.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e1_funkcii.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e1_funkcii.py`:

```python
# -*- coding: utf-8 -*-
"""Функции этапа 1 общего модуля через COM: хеш, единицы (Р23), разузлование (§6, Р26) на тестовом дереве в транзакции."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect, same, tx  # noqa: E402
from fixtures import add_row, card, new_spec, srv, ss, unit, write  # noqa: E402

SEPTIK = [("00-00010526", 1, 2834.672495), ("00000017651", 2, 2834.683076), ("00-00006019", 1, 661.278277),
          ("00-00012879", 1, 2834.683076)]


def t_hash_units(B):
    s = srv(B)
    h = s.ХешMD5("а/б")
    assert len(h) == 32 and h == s.ХешMD5("а/б") and h != s.ХешMD5("а/в") and h == h.lower()
    for text, key in (("пач", "пак"), ("м.п.", "м"), ("пог.м", "м"), ("балон", "шт"), ("паков", "паков"), (" т", "т"), ("шт.", "шт")):
        assert s.КлючЕдиницыСпецификации(text) == key, (text, s.КлючЕдиницыСпецификации(text))
    assert s.ЕдиницыСовпадают("пог.м", unit(B, "0101")) and s.ЕдиницыСовпадают("пач", "пак")
    assert not s.ЕдиницыСовпадают("т", "кг") and not s.ЕдиницыСовпадают("", "")
    print("OK хеш и единицы")


def t_unfold(B):
    with tx(B):
        u = unit(B, "2938")
        sep = new_spec(B, "Узел", "Септик", unit_ref=u)
        for code, qty, price in SEPTIK:
            add_row(B, sep, card_ref=card(B, code), qty=qty, price=price)
        sep_ref = write(sep)
        gen = new_spec(B, "Узел", "Генератор", unit_ref=u)
        gen.НоменклатураВузла = card(B, "00-00023184")
        gen_ref = write(gen)
        zm = new_spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        add_row(B, zm, "Узел", node=sep_ref, unit_ref=u)
        add_row(B, zm, card_ref=card(B, "00-00022147"), qty=18.75)
        zm_ref = write(zm)
        h = new_spec(B, "Дом", "розгортання", ss_ref=ss(B, "000000009"))
        add_row(B, h, card_ref=card(B, "00-00009980"), qty=92, unit_ref=unit(B, "0101"), k=3, ksrc="Вручную")
        add_row(B, h, "Узел", node=zm_ref, unit_ref=u)
        add_row(B, h, "Узел", node=gen_ref, unit_ref=u, ready=True)
        add_row(B, h, pos="не постачається", ne_post=True, comment="тест", ksrc="НеНайден")
        h_ref = write(h)
        r = srv(B).РазвернутьДоКарточек(h_ref)
        lv = [r.Листья.Получить(i) for i in range(r.Листья.Количество())]
        ek = [r.Экземпляры.Получить(i) for i in range(r.Экземпляры.Количество())]
        assert len(lv) == 7 and len(ek) == 2, (len(lv), len(ek))
        assert len({x.КлючЛиста for x in lv}) == 7 and all(len(x.КлючЛиста) == 32 for x in lv)
        plint = [x for x in lv if same(B, x.Номенклатура, card(B, "00-00009980"))][0]
        assert round(plint.Количество, 4) == 30.6667 and plint.Уровень == 0 and plint.КлючЭкземпляра == ""
        gl = [x for x in lv if same(B, x.Номенклатура, card(B, "00-00023184"))]
        assert len(gl) == 1, "«купується готовим» — лист с карточкой узла"
        sep_l = [x for x in lv if x.Уровень == 2]
        assert sorted(round(x.Количество, 6) for x in sep_l) == [1, 1, 1, 2]
        assert round(sum(x.СуммаНормы for x in sep_l), 2) == 12000.00
        codes = f"{zm_ref.Код.strip()}/{sep_ref.Код.strip()}"
        assert all(x.ПутьУзла == codes for x in sep_l), [x.ПутьУзла for x in sep_l]
        e_sep = [e for e in ek if same(B, e.Узел, sep_ref)][0]
        assert e_sep.Уровень == 2 and round(e_sep.СуммаНормы, 2) == 12000.00
        assert all(x.КлючЭкземпляра == e_sep.КлючЭкземпляра for x in sep_l)
        r2 = srv(B).РазвернутьДоКарточек(h_ref, True)
        assert r2.Листья.Количество() == 8, "строка «Не постачається» — только с флагом"
        ya = [r2.ЯвныеПозиции.Получить(i).ПозицияСС for i in range(r2.ЯвныеПозиции.Количество())]
        assert ya == ["не постачається"], ya
    print("OK разузлование")


def t_ploshchadnoy(B):
    with tx(B):
        m2 = unit(B, "0123")
        pid = new_spec(B, "Узел", "Підлога", base_qty=18, unit_ref=m2)
        add_row(B, pid, card_ref=card(B, "00-00009980"), qty=36)
        pid_ref = write(pid)
        h = new_spec(B, "Дом", "площадний", ss_ref=ss(B, "000000009"))
        add_row(B, h, "Узел", node=pid_ref, qty=18, unit_ref=m2)
        r = srv(B).РазвернутьДоКарточек(write(h))
        assert round(r.Листья.Получить(0).Количество, 6) == 36, "Р26: состав на базовую площадь 18 м²"
    print("OK площадной узел")


if __name__ == "__main__":
    B = connect()
    t_hash_units(B)
    t_unfold(B)
    t_ploshchadnoy(B)
    print("OK test_e1_funkcii")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_funkcii.py`
Expected: FAIL — `AttributeError: …А_СпецификацииДомовСервер` (модуля в базе нет).

- [ ] **Step 3: Реализовать сборщик и две части модуля**

`_Rarzrabotki/Python/test/spec_domov/build_bsl.py`:

```python
# -*- coding: utf-8 -*-
"""Сборка модулей П1 из исходников spec_domov/bsl/*.bsl в дамп WORKTREE (UTF-8 BOM + CRLF).
Строка «<отступ>//@ЗАПРОС <имя>» заменяется многострочным литералом BSL из zaprosy/<имя>.sql (кавычки удваиваются,
строки после первой начинаются с «|»): в BSL попадает тот же текст, что исполнил Python COM тест (Rule #-1).
Партии накопительные: 1 → 2a → 2b → 2c. Запуск: python build_bsl.py --partiya 1"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, HERE, write_text, zapros  # noqa: E402

ORDER = ["1", "2a", "2b", "2c"]
SERVER_PARTS = [("1", "srv_10_obschee.bsl"), ("1", "srv_15_edinicy.bsl"), ("1", "srv_20_derevo.bsl"),
                ("1", "srv_30_cikly.bsl"), ("1", "srv_32_proverki.bsl"), ("1", "srv_35_etalon.bsl"),
                ("2a", "srv_40_edinicy2.bsl"), ("2a", "srv_50_fakt.bsl"), ("2a", "srv_55_zapolnenie.bsl"),
                ("2b", "srv_60_uzly.bsl"), ("2b", "srv_70_nulevye.bsl"), ("2b", "srv_80_ceny.bsl"),
                ("2b", "srv_90_sravnenie.bsl"), ("2c", "srv_A0_perepriv.bsl"), ("2c", "srv_B0_kopiya.bsl")]
HEADER = ("////////////////////////////////////////////////////////////////////////////////\n"
          "// А_СпецификацииДомовСервер — многоуровневая спецификация домов и узлов (П1 «Специфікація і еталон»).\n"
          "// Спецификация: docs/superpowers/specs/2026-10-03-mnogourovnevaya-specifikaciya-domov-design.md (§4–§6).\n"
          "// Сервер + внешнее соединение, не привилегированный. Собран build_bsl.py из _Rarzrabotki/Python/test/spec_domov/bsl;\n"
          "// тексты запросов — spec_domov/zaprosy/*.sql (проверены Python COM). Править исходники, а не этот файл.\n"
          "////////////////////////////////////////////////////////////////////////////////\n\n")
MODULE = "CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"
OBJECT_MODULE = "Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl"
RE_MARK = re.compile(r"^(?P<ind>[ \t]*)//@ЗАПРОС (?P<name>\w+)[ \t]*$", re.M)


def literal(name, ind):
    lines = zapros(name).replace('"', '""').split("\n")
    out = [f'{ind}"{lines[0]}'] + [f"{ind}|{ln}" for ln in lines[1:]]
    out[-1] += '"'
    return "\n".join(out)


def expand(text):
    text = RE_MARK.sub(lambda m: literal(m.group("name"), m.group("ind")), text)
    assert "//@ЗАПРОС" not in text
    return text


def read_src(name):
    return open(os.path.join(HERE, "bsl", name), encoding="utf-8-sig").read().replace("\r\n", "\n")


def build(partiya):
    upto = ORDER.index(partiya)
    parts = [read_src(f) for p, f in SERVER_PARTS if ORDER.index(p) <= upto]
    text = HEADER + "\n".join(expand(p).rstrip("\n") + "\n" for p in parts)
    write_text(os.path.join(BB_WT, MODULE), text)
    print(f"  {MODULE}: частей {len(parts)}, строк {text.count(chr(10))}")
    # Модуль объекта грузится только партией 1B: в партиях 2A–2C его не трогать, иначе worktree-копия разойдётся
    # с каноном платформы в основном каталоге (dump_check --canon, Task 28) и to_main.py --verify e1b (Task 45) упадёт.
    if partiya == "1" and os.path.exists(os.path.join(HERE, "bsl", "obj_module.bsl")):
        write_text(os.path.join(BB_WT, OBJECT_MODULE), expand(read_src("obj_module.bsl")))
        print(f"  {OBJECT_MODULE}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--partiya", choices=ORDER, required=True)
    build(ap.parse_args().partiya)
```

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_10_obschee.bsl`:

```bsl
#Область Общее

// MD5 строки — 32 шестнадцатеричных символа в нижнем регистре ([M5] §4.4: КлючЛиста, КлючЭкземпляра).
Функция ХешMD5(Знач Текст) Экспорт

	Хеширование = Новый ХешированиеДанных(ХешФункция.MD5);
	Хеширование.Добавить(Текст);
	Возврат НРег(ПолучитьHexСтрокуИзДвоичныхДанных(Хеширование.ХешСумма));

КонецФункции

// Пустой или повторный КлючСтроки получает новый УИД (образец — А_ТехкартаКМД.ЗаповнитиКлючіРядків).
// Копия всей спецификации сохраняет ключи (§4.3), копия строки в форме даёт дубль — он получает новый ключ.
Процедура ЗаполнитьКлючиСтрок(ТабличнаяЧасть) Экспорт

	Занятые = Новый Соответствие;
	Для Каждого СтрокаТЧ Из ТабличнаяЧасть Цикл
		КлючТекстом = Строка(СтрокаТЧ.КлючСтроки);
		Если Не ЗначениеЗаполнено(СтрокаТЧ.КлючСтроки) Или Занятые.Получить(КлючТекстом) <> Неопределено Тогда
			СтрокаТЧ.КлючСтроки = Новый УникальныйИдентификатор;
			КлючТекстом = Строка(СтрокаТЧ.КлючСтроки);
		КонецЕсли;
		Занятые.Вставить(КлючТекстом, Истина);
	КонецЦикла;

КонецПроцедуры

// Позиции СС по ключу А_СтруктураСебестоимостиСервер.КлючПозиции (этап + текст; ОН и единица в ключ не входят).
// Возврат: Соответствие — ключ → Структура(НомерСтроки, Этап, НоменклатураСС, ОбщееНазвание, Количество, ЕдиницаСС,
// КоличествоОригинал, ЕдиницаОригинал, Сумма, Единица). Дубль ключа (две нулевые «Профіль для підлоги» 15 м) — первая строка.
Функция ПозицииСтруктурыСебестоимости(СсылкаСС) Экспорт

	Позиции = Новый Соответствие;
	Если Не ЗначениеЗаполнено(СсылкаСС) Тогда
		Возврат Позиции;
	КонецЕсли;
	Для Каждого Компл Из СсылкаСС.Комплектующие Цикл
		Ключ = А_СтруктураСебестоимостиСервер.КлючПозиции(Компл.Этап, Компл.НоменклатураСС);
		Если Позиции.Получить(Ключ) <> Неопределено Тогда
			Продолжить;
		КонецЕсли;
		Позиции.Вставить(Ключ, Новый Структура(
			"НомерСтроки, Этап, НоменклатураСС, ОбщееНазвание, Количество, ЕдиницаСС, КоличествоОригинал, ЕдиницаОригинал, Сумма, Единица",
			Компл.НомерСтроки, Компл.Этап, Компл.НоменклатураСС, Компл.ОбщееНазвание, Компл.Количество, Компл.ЕдиницаСС,
			Компл.КоличествоСС_Оригинал, Компл.ЕдиницаСС_Оригинал, Компл.Сумма, Компл.Единица));
	КонецЦикла;
	Возврат Позиции;

КонецФункции

Функция ЭтоЦифра(Символ)

	Возврат Символ <> "" И СтрНайти("0123456789", Символ) > 0;

КонецФункции

Функция ЭтоБуква(Символ)

	Возврат Символ <> "" И ВРег(Символ) <> НРег(Символ);

КонецФункции

#КонецОбласти
```

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_15_edinicy.bsl`:

```bsl
#Область Единицы

// Ключ единицы спецификации (Р23): ключ СС (А_СтруктураСебестоимостиСервер.КлючЕдиницы) плюс свои синонимы
// «пог.м» = «м», «пач» = «пак», «балон» = «шт». Функция СС не меняется. «паков» не сводится (Д3).
Функция КлючЕдиницыСпецификации(Знач ТекстЕдиницы) Экспорт

	Ключ = А_СтруктураСебестоимостиСервер.КлючЕдиницы(ТекстЕдиницы);
	Если Ключ = "пог.м" Тогда
		Возврат "м";
	ИначеЕсли Ключ = "пач" Тогда
		Возврат "пак";
	ИначеЕсли Ключ = "балон" Тогда
		Возврат "шт";
	КонецЕсли;
	Возврат Ключ;

КонецФункции

// Единицы сравниваются только ключами по наименованию, не ссылками: «т», «шт», «м2», «м3», «пак» живут на двух элементах
// классификатора (E3 §4.1). Каждый параметр — строка или СправочникСсылка.КлассификаторЕдиницИзмерения.
Функция ЕдиницыСовпадают(ЕдиницаА, ЕдиницаБ) Экспорт

	КлючА = КлючЕдиницыСпецификации(НаименованиеЕдиницы(ЕдиницаА));
	КлючБ = КлючЕдиницыСпецификации(НаименованиеЕдиницы(ЕдиницаБ));
	Возврат Не ПустаяСтрока(КлючА) И КлючА = КлючБ;

КонецФункции

// k строки по единицам (§4.3, [M5] §1.2 п. 3): у «Деталь» с карточкой совпали единицы — 1 и «Одиниці збігаються»,
// иначе 1 и «Не знайдено» (k вводит бухгалтер); «Вузол» и «Група» — всегда 1. Строка — ТЧ, ТЗ или данные формы.
Процедура УстановитьКоэффициентПоЕдиницам(СтрокаСостава) Экспорт

	Источники = Перечисления.А_ИсточникиКоэффициентаПересчетаСС;
	СтрокаСостава.Коэффициент = 1;
	Если СтрокаСостава.ВидСтроки <> Перечисления.А_ВидыСтрокСпецификацииДомов.Деталь Тогда
		СтрокаСостава.ИсточникКоэффициента = Источники.ЕдиницыСовпадают;
	ИначеЕсли Не ЗначениеЗаполнено(СтрокаСостава.Номенклатура) Тогда
		СтрокаСостава.ИсточникКоэффициента = Источники.НеНайден;
	ИначеЕсли ЕдиницыСовпадают(ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СтрокаСостава.Номенклатура, "БазоваяЕдиницаИзмерения"),
		СтрокаСостава.ЕдиницаСтроки) Тогда
		СтрокаСостава.ИсточникКоэффициента = Источники.ЕдиницыСовпадают;
	Иначе
		СтрокаСостава.ИсточникКоэффициента = Источники.НеНайден;
	КонецЕсли;

КонецПроцедуры

Функция НаименованиеЕдиницы(Единица)

	Если ТипЗнч(Единица) = Тип("Строка") Тогда
		Возврат Единица;
	ИначеЕсли Не ЗначениеЗаполнено(Единица) Тогда
		Возврат "";
	КонецЕсли;
	Возврат ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Единица, "Наименование");

КонецФункции

#КонецОбласти
```

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/build_bsl.py --partiya 1`
Expected: FAIL `FileNotFoundError: …srv_20_derevo.bsl` — это нормально до Task 21. Временная проверка частей: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_10_obschee.bsl _Rarzrabotki/Python/test/spec_domov/bsl/srv_15_edinicy.bsl` (исходники пишутся инструментом Write без BOM — сообщения «нет BOM»/«не CRLF» для `bsl/*.bsl` ожидаемы и игнорируются; остальных ошибок быть не должно).

- [ ] **Step 4: PASS** — `test_e1_funkcii.py::t_hash_units` после загрузки 1B (Task 28).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/build_bsl.py _Rarzrabotki/Python/test/spec_domov/test_e1_funkcii.py \
  _Rarzrabotki/Python/test/spec_domov/bsl/srv_10_obschee.bsl _Rarzrabotki/Python/test/spec_domov/bsl/srv_15_edinicy.bsl
git commit -m "feat(spec_domov): сборщик модулей с вставкой проверенных запросов; общее и единицы Р23

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 21: Дерево и разузлование (§6, [M5] §4.2–4.4, Т6)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_20_derevo.bsl`
- Test: `test_e1_funkcii.py::t_unfold`, `::t_ploshchadnoy` (Task 20)

- [ ] **Step 1: Тест уже написан** (Task 20). **Step 2: FAIL** — тот же `AttributeError`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_20_derevo.bsl`:

```bsl
#Область ДеревоИРазузлование

// Загрузка дерева спецификаций по уровням: запрос uroven_dereva (Rule #-1, Task 18) по массиву ссылок, пока появляются
// новые узлы. Корень — СправочникСсылка или СправочникОбъект.А_СпецификацииДомов (несохранённые строки формы).
// Возврат: Структура(Кэш — Соответствие ключ → данные; КлючКорня — ссылка, у нового объекта — строка "КОРЕНЬ").
// Данные: Структура(Ссылка, Код, Наименование, ВидСпецификации, Статус, КоличествоВыпуска, ЕдиницаВыпуска,
// НоменклатураВузла, СтруктураСебестоимости, ДатаИзменения, Строки — Массив строк «Состав» или строк результата запроса).
Функция ЗагрузитьДеревоСпецификаций(Корень) Экспорт

	Кэш = Новый Соответствие;
	Очередь = Новый Массив;
	Если ТипЗнч(Корень) = Тип("СправочникОбъект.А_СпецификацииДомов") Тогда
		КлючКорня = ?(Корень.Ссылка.Пустая(), "КОРЕНЬ", Корень.Ссылка);
		СтрокиКорня = Новый Массив;
		Для Каждого СтрокаТЧ Из Корень.Состав Цикл
			СтрокиКорня.Добавить(СтрокаТЧ);
		КонецЦикла;
		Кэш.Вставить(КлючКорня, ДанныеСпецификации(Корень, СтрокиКорня));
		ДобавитьУзлыВОчередь(СтрокиКорня, Кэш, Очередь);
	Иначе
		КлючКорня = Корень;
		Очередь.Добавить(Корень);
	КонецЕсли;

	Уровней = 0;
	Пока Очередь.Количество() > 0 И Уровней < 50 Цикл
		Уровней = Уровней + 1;
		Запрос = Новый Запрос(ТекстЗапросаУровняДерева());
		Запрос.УстановитьПараметр("СпецификацииУровня", Очередь);
		Пакет = Запрос.ВыполнитьПакет();
		Шапки = Пакет[0].Выгрузить();
		СтрокиСостава = Пакет[1].Выгрузить();
		Очередь = Новый Массив;
		Для Каждого Шапка Из Шапки Цикл
			Строки = Новый Массив;
			Для Каждого СтрокаСостава Из СтрокиСостава.НайтиСтроки(Новый Структура("Спецификация", Шапка.Ссылка)) Цикл
				Строки.Добавить(СтрокаСостава);
			КонецЦикла;
			Кэш.Вставить(Шапка.Ссылка, ДанныеСпецификации(Шапка, Строки));
		КонецЦикла;
		Для Каждого Шапка Из Шапки Цикл
			ДобавитьУзлыВОчередь(Кэш.Получить(Шапка.Ссылка).Строки, Кэш, Очередь);
		КонецЦикла;
	КонецЦикла;

	Возврат Новый Структура("Кэш, КлючКорня", Кэш, КлючКорня);

КонецФункции

// Разузлование дома, варианта или узла до карточек (§6 п. 3–4, [M5] §4.2–4.4). Ничего не пишет.
// КолЕдСтроки = q × M / КоличествоВыпуска(S); узел (не «купується готовим») — вглубь с M' = КолЕдСтроки;
// лист: Количество = КолЕдСтроки / k, СуммаНормы = Окр(Количество × ЦенаНормы, 2); округлений внутри нет (Р8).
// «Не постачається» — только при ВключатьНеПоставляется (ціни норми и И4, Р24). Дом — уровень 0; Уровень >= 10 — исключение (Т6).
// Корень — ссылка или объект. Возврат: Структура(Листья, Экземпляры, ЯвныеПозиции) — колонки как ТЧ РК
// НормаДома / ЭкземплярыУзлов (§4.7) плюс служебные.
Функция РазвернутьДоКарточек(Корень, ВключатьНеПоставляется = Ложь) Экспорт

	Дерево = ЗагрузитьДеревоСпецификаций(Корень);
	Результат = Новый Структура("Листья, Экземпляры, ЯвныеПозиции",
		НоваяТаблицаЛистьев(), НоваяТаблицаЭкземпляров(), НоваяТаблицаЯвныхПозиций());
	Путь = Новый Массив;
	Путь.Добавить(Дерево.КлючКорня);
	Обойти(Дерево.Кэш, Дерево.КлючКорня, 1, "", "", Справочники.ЭтапыРабот.ПустаяСсылка(), "", 0, Путь, Результат,
		ВключатьНеПоставляется);
	ПосчитатьСуммыЭкземпляров(Результат);
	Возврат Результат;

КонецФункции

Функция НоваяТаблицаЛистьев() Экспорт

	Таблица = Новый ТаблицаЗначений;
	Колонки = Таблица.Колонки;
	Колонки.Добавить("КлючЛиста", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(32)));
	Колонки.Добавить("КлючСтроки", Новый ОписаниеТипов("УникальныйИдентификатор"));
	Колонки.Добавить("КлючЭкземпляра", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(32)));
	Колонки.Добавить("Спецификация", Новый ОписаниеТипов("СправочникСсылка.А_СпецификацииДомов"));
	Колонки.Добавить("ВидСтроки", Новый ОписаниеТипов("ПеречислениеСсылка.А_ВидыСтрокСпецификацииДомов"));
	Колонки.Добавить("Номенклатура", Новый ОписаниеТипов("СправочникСсылка.Номенклатура"));
	Колонки.Добавить("ОбщееНазвание", Новый ОписаниеТипов("СправочникСсылка.ОбщиеНазванияНоменклатуры"));
	Колонки.Добавить("ЭтапСС", Новый ОписаниеТипов("СправочникСсылка.ЭтапыРабот"));
	Колонки.Добавить("ПозицияСС", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(200)));
	Колонки.Добавить("ЕдиницаСтроки", Новый ОписаниеТипов("СправочникСсылка.КлассификаторЕдиницИзмерения"));
	Колонки.Добавить("КоличествоЕдСтроки", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("Количество", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("ЦенаНормы", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("СуммаНормы", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("ПутьУзла", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(200)));
	Колонки.Добавить("Уровень", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("НеПоставляется", Новый ОписаниеТипов("Булево"));
	Колонки.Добавить("Коэффициент", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("ИсточникКоэффициента", Новый ОписаниеТипов("ПеречислениеСсылка.А_ИсточникиКоэффициентаПересчетаСС"));
	Колонки.Добавить("ИсточникКарточки", Новый ОписаниеТипов("ПеречислениеСсылка.А_ИсточникиКарточкиСпецификации"));
	Колонки.Добавить("Комментарий", Новый ОписаниеТипов("Строка"));
	Возврат Таблица;

КонецФункции

Функция НоваяТаблицаЭкземпляров() Экспорт

	Таблица = Новый ТаблицаЗначений;
	Колонки = Таблица.Колонки;
	Колонки.Добавить("КлючЭкземпляра", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(32)));
	Колонки.Добавить("КлючЭкземпляраРодителя", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(32)));
	Колонки.Добавить("Узел", Новый ОписаниеТипов("СправочникСсылка.А_СпецификацииДомов"));
	Колонки.Добавить("НоменклатураВузла", Новый ОписаниеТипов("СправочникСсылка.Номенклатура"));
	Колонки.Добавить("ПокупаетсяГотовым", Новый ОписаниеТипов("Булево"));
	Колонки.Добавить("ПутьУзла", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(200)));
	Колонки.Добавить("Уровень", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("Количество", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("СуммаНормы", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("Спецификация", Новый ОписаниеТипов("СправочникСсылка.А_СпецификацииДомов"));
	Колонки.Добавить("КлючСтроки", Новый ОписаниеТипов("УникальныйИдентификатор"));
	Возврат Таблица;

КонецФункции

Функция НоваяТаблицаЯвныхПозиций() Экспорт

	Таблица = Новый ТаблицаЗначений;
	Колонки = Таблица.Колонки;
	Колонки.Добавить("Спецификация", Новый ОписаниеТипов("СправочникСсылка.А_СпецификацииДомов"));
	Колонки.Добавить("КлючСтроки", Новый ОписаниеТипов("УникальныйИдентификатор"));
	Колонки.Добавить("ЭтапСС", Новый ОписаниеТипов("СправочникСсылка.ЭтапыРабот"));
	Колонки.Добавить("ПозицияСС", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(200)));
	Колонки.Добавить("КоличествоЕдСтроки", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("ЕдиницаСтроки", Новый ОписаниеТипов("СправочникСсылка.КлассификаторЕдиницИзмерения"));
	Колонки.Добавить("ЭтоУзел", Новый ОписаниеТипов("Булево"));
	Колонки.Добавить("НеПоставляется", Новый ОписаниеТипов("Булево"));
	Колонки.Добавить("Комментарий", Новый ОписаниеТипов("Строка"));
	Возврат Таблица;

КонецФункции

Процедура Обойти(Кэш, КлючСпец, Множитель, ЦепочкаКлючей, КлючЭкз, ЭтапРод, ПозицияРод, Уровень, Путь, Результат, ВключатьНП)

	Если Уровень >= 10 Тогда
		ВызватьИсключение СтрШаблон(НСтр("ru = 'Глубина вложенности узлов больше 9 (дом — уровень 0): %1'; uk = 'Глибина вкладеності вузлів більша за 9 (будинок — рівень 0): %1'"),
			ПредставлениеПути(Путь, Кэш));
	КонецЕсли;
	Данные = Кэш.Получить(КлючСпец);
	Если Данные = Неопределено Тогда
		ВызватьИсключение СтрШаблон(НСтр("ru = 'Спецификация не найдена: %1'; uk = 'Специфікацію не знайдено: %1'"), Строка(КлючСпец));
	КонецЕсли;
	Доля = Множитель / ?(Данные.КоличествоВыпуска = 0, 1, Данные.КоличествоВыпуска);
	СсылкаСпец = ?(ТипЗнч(КлючСпец) = Тип("Строка"), Справочники.А_СпецификацииДомов.ПустаяСсылка(), КлючСпец);
	ВидУзел = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;

	Для Каждого Стр Из Данные.Строки Цикл
		Если Стр.НеПоставляется И Не ВключатьНП Тогда
			Продолжить;
		КонецЕсли;
		КолЕдСтр = Стр.Количество * Доля;
		Цепочка = ЦепочкаКлючей + Строка(Стр.КлючСтроки);
		Ключ = ХешMD5(Цепочка);
		Этап = ?(ЗначениеЗаполнено(Стр.ЭтапСС), Стр.ЭтапСС, ЭтапРод);
		Позиция = ?(ЗначениеЗаполнено(Стр.ПозицияСС), Стр.ПозицияСС, ПозицияРод);
		ЭтоУзел = (Стр.ВидСтроки = ВидУзел);
		Если ЗначениеЗаполнено(Стр.ПозицияСС) Тогда
			Явная = Результат.ЯвныеПозиции.Добавить();
			Явная.Спецификация = СсылкаСпец;
			Явная.КлючСтроки = Стр.КлючСтроки;
			Явная.ЭтапСС = Этап;
			Явная.ПозицияСС = Стр.ПозицияСС;
			Явная.КоличествоЕдСтроки = КолЕдСтр;
			Явная.ЕдиницаСтроки = Стр.ЕдиницаСтроки;
			Явная.ЭтоУзел = ЭтоУзел;
			Явная.НеПоставляется = Стр.НеПоставляется;
			Явная.Комментарий = Стр.Комментарий;
		КонецЕсли;
		Если ЭтоУзел И Не Стр.ПокупаетсяГотовым Тогда
			Если Путь.Найти(Стр.Узел) <> Неопределено Тогда
				Путь.Добавить(Стр.Узел);
				ВызватьИсключение СтрШаблон(НСтр("ru = 'Узел входит сам в себя: %1'; uk = 'Вузол входить сам у себе: %1'"),
					ПредставлениеПути(Путь, Кэш));
			КонецЕсли;
			ДанныеУзла = Кэш.Получить(Стр.Узел);
			Экз = Результат.Экземпляры.Добавить();
			Экз.КлючЭкземпляра = Ключ;
			Экз.КлючЭкземпляраРодителя = КлючЭкз;
			Экз.Узел = Стр.Узел;
			Экз.НоменклатураВузла = ?(ДанныеУзла = Неопределено, Справочники.Номенклатура.ПустаяСсылка(), ДанныеУзла.НоменклатураВузла);
			Экз.ПутьУзла = ПутьКодами(Путь, Кэш, Стр.Узел);
			Экз.Уровень = Уровень + 1;
			Экз.Количество = КолЕдСтр;
			Экз.Спецификация = СсылкаСпец;
			Экз.КлючСтроки = Стр.КлючСтроки;
			Путь.Добавить(Стр.Узел);
			Обойти(Кэш, Стр.Узел, КолЕдСтр, Цепочка + "/", Ключ, Этап, Позиция, Уровень + 1, Путь, Результат, ВключатьНП);
			Путь.Удалить(Путь.ВГраница());
		Иначе
			Лист = Результат.Листья.Добавить();
			Лист.КлючЛиста = Ключ;
			Лист.КлючСтроки = Стр.КлючСтроки;
			Лист.КлючЭкземпляра = КлючЭкз;
			Лист.Спецификация = СсылкаСпец;
			Лист.ВидСтроки = Стр.ВидСтроки;
			Если ЭтоУзел Тогда
				ДанныеУзла = Кэш.Получить(Стр.Узел);
				Лист.Номенклатура = ?(ДанныеУзла = Неопределено, Справочники.Номенклатура.ПустаяСсылка(), ДанныеУзла.НоменклатураВузла);
			Иначе
				Лист.Номенклатура = Стр.Номенклатура;
			КонецЕсли;
			Лист.ОбщееНазвание = Стр.ОбщееНазвание;
			Лист.ЭтапСС = Этап;
			Лист.ПозицияСС = Позиция;
			Лист.ЕдиницаСтроки = Стр.ЕдиницаСтроки;
			Лист.КоличествоЕдСтроки = КолЕдСтр;
			Лист.Коэффициент = ?(Стр.Коэффициент = 0, 1, Стр.Коэффициент);
			Лист.Количество = КолЕдСтр / Лист.Коэффициент;
			Лист.ЦенаНормы = Стр.ЦенаНормы;
			Лист.СуммаНормы = Окр(Лист.Количество * Лист.ЦенаНормы, 2);
			Лист.ПутьУзла = ПутьКодами(Путь, Кэш);
			Лист.Уровень = Уровень;
			Лист.НеПоставляется = Стр.НеПоставляется;
			Лист.ИсточникКоэффициента = Стр.ИсточникКоэффициента;
			Лист.ИсточникКарточки = Стр.ИсточникКарточки;
			Лист.Комментарий = Стр.Комментарий;
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

// V_U экземпляра = Σ СуммаНормы листьев под ним на всех уровнях ([M5] §2.9.2).
Процедура ПосчитатьСуммыЭкземпляров(Результат)

	Родители = Новый Соответствие;
	ПоКлючу = Новый Соответствие;
	Для Каждого Экз Из Результат.Экземпляры Цикл
		Родители.Вставить(Экз.КлючЭкземпляра, Экз.КлючЭкземпляраРодителя);
		ПоКлючу.Вставить(Экз.КлючЭкземпляра, Экз);
		Экз.СуммаНормы = 0;
	КонецЦикла;
	Для Каждого Лист Из Результат.Листья Цикл
		КлючЭкз = Лист.КлючЭкземпляра;
		Пока ЗначениеЗаполнено(КлючЭкз) Цикл
			Экз = ПоКлючу.Получить(КлючЭкз);
			Если Экз = Неопределено Тогда
				Прервать;
			КонецЕсли;
			Экз.СуммаНормы = Экз.СуммаНормы + Лист.СуммаНормы;
			КлючЭкз = Родители.Получить(КлючЭкз);
		КонецЦикла;
	КонецЦикла;

КонецПроцедуры

Процедура ДобавитьУзлыВОчередь(Строки, Кэш, Очередь)

	ВидУзел = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	Для Каждого СтрокаСостава Из Строки Цикл
		Если СтрокаСостава.ВидСтроки = ВидУзел И ЗначениеЗаполнено(СтрокаСостава.Узел)
			И Кэш.Получить(СтрокаСостава.Узел) = Неопределено И Очередь.Найти(СтрокаСостава.Узел) = Неопределено Тогда
			Очередь.Добавить(СтрокаСостава.Узел);
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

Функция ДанныеСпецификации(Источник, Строки)

	Возврат Новый Структура(
		"Ссылка, Код, Наименование, ВидСпецификации, Статус, КоличествоВыпуска, ЕдиницаВыпуска, НоменклатураВузла, СтруктураСебестоимости, ДатаИзменения, Строки",
		Источник.Ссылка, Источник.Код, Источник.Наименование, Источник.ВидСпецификации, Источник.Статус,
		Источник.КоличествоВыпуска, Источник.ЕдиницаВыпуска, Источник.НоменклатураВузла,
		Источник.СтруктураСебестоимости, Источник.ДатаИзменения, Строки);

КонецФункции

// Путь из КОДОВ узлов (без корня): «000000012/000000015» ([M5] §2.9.1 — переименование узла путь не портит).
Функция ПутьКодами(Путь, Кэш, Последний = Неопределено)

	Коды = Новый Массив;
	Для Индекс = 1 По Путь.ВГраница() Цикл
		Данные = Кэш.Получить(Путь[Индекс]);
		Коды.Добавить(?(Данные = Неопределено, "", СокрЛП(Данные.Код)));
	КонецЦикла;
	Если Последний <> Неопределено Тогда
		Данные = Кэш.Получить(Последний);
		Коды.Добавить(?(Данные = Неопределено, "", СокрЛП(Данные.Код)));
	КонецЕсли;
	Возврат Лев(СтрСоединить(Коды, "/"), 200);

КонецФункции

Функция ПредставлениеПути(Путь, Кэш)

	Части = Новый Массив;
	Для Каждого КлючСпец Из Путь Цикл
		Данные = Кэш.Получить(КлючСпец);
		Части.Добавить(?(Данные = Неопределено, Строка(КлючСпец), Данные.Наименование));
	КонецЦикла;
	Возврат СтрСоединить(Части, " → ");

КонецФункции

Функция ТекстЗапросаУровняДерева()

	Возврат
	//@ЗАПРОС uroven_dereva
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS** — Task 28 (`t_unfold`, `t_ploshchadnoy`). Статическая проверка исходника: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_20_derevo.bsl` — кроме «нет BOM/не CRLF/неразвёрнутый //@ЗАПРОС» (исходник) ошибок нет.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_20_derevo.bsl
git commit -m "feat(spec_domov): разузлование до карточек — дерево по уровням, экземпляры узлов, явные позиции, Т6

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 22: Входимость, уровень в домах и поиск циклов (§4.6, Т6, E3 §5)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_30_cikly.bsl`
- Test: `test_e1_priemka.py::k1…k5`

- [ ] **Step 1: Тест уже написан** (Task 19). **Step 2: FAIL** — `AttributeError` (модуля в базе нет).

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_30_cikly.bsl`:

```bsl
#Область ВходимостьИЦиклы

// «Входимість» узла ([M5] §3.4): все спецификации-предки по ссылкам «Состав.Узел», вверх по уровням (запрос roditeli).
// Возврат: ТаблицаЗначений(Спецификация, ВидСпецификации, СтруктураСебестоимости, Статус, Уровень, ЧерезУзел);
// Уровень — расстояние вверх (1 — строка этой спецификации прямо ссылается на узел).
Функция Входимость(Узел) Экспорт

	Результат = Новый ТаблицаЗначений;
	Результат.Колонки.Добавить("Спецификация");
	Результат.Колонки.Добавить("ВидСпецификации");
	Результат.Колонки.Добавить("СтруктураСебестоимости");
	Результат.Колонки.Добавить("Статус");
	Результат.Колонки.Добавить("Уровень", Новый ОписаниеТипов("Число"));
	Результат.Колонки.Добавить("ЧерезУзел");
	Пройдены = Новый Соответствие;
	Пройдены.Вставить(Узел, Истина);
	Текущие = Новый Массив;
	Текущие.Добавить(Узел);
	Уровень = 0;
	Пока Текущие.Количество() > 0 И Уровень < 50 Цикл
		Уровень = Уровень + 1;
		Запрос = Новый Запрос(ТекстЗапросаРодителей());
		Запрос.УстановитьПараметр("Узлы", Текущие);
		Выборка = Запрос.Выполнить().Выбрать();
		Текущие = Новый Массив;
		Пока Выборка.Следующий() Цикл
			Если Пройдены.Получить(Выборка.Спецификация) <> Неопределено Тогда
				Продолжить;
			КонецЕсли;
			Пройдены.Вставить(Выборка.Спецификация, Истина);
			НоваяСтрока = Результат.Добавить();
			ЗаполнитьЗначенияСвойств(НоваяСтрока, Выборка);
			НоваяСтрока.Уровень = Уровень;
			Текущие.Добавить(Выборка.Спецификация);
		КонецЦикла;
	КонецЦикла;
	Возврат Результат;

КонецФункции

// Наибольший уровень, на котором спецификация стоит в домах (Т6): дом — 0; узел — max(уровень родителя + 1);
// узел, который пока не входит ни в один дом, — 1 (выше дома он стоять не может). Нужен проверке глубины при записи:
// обход только вниз не ловит узел, записанный ВНУТРЬ уже глубокого дерева (E3 §5).
Функция УровеньВДомах(Спецификация) Экспорт

	Если Не ЗначениеЗаполнено(Спецификация) Тогда
		Возврат 1;
	КонецЕсли;
	Возврат УровеньСпецификации(Спецификация, Новый Соответствие, Новый Массив);

КонецФункции

Функция УровеньСпецификации(Спецификация, Кэш, Стек)

	Известный = Кэш.Получить(Спецификация);
	Если Известный <> Неопределено Тогда
		Возврат Известный;
	КонецЕсли;
	Если ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Спецификация, "ВидСпецификации") = Перечисления.А_ВидыСпецификацийДомов.Дом Тогда
		Кэш.Вставить(Спецификация, 0);
		Возврат 0;
	КонецЕсли;
	Если Стек.Найти(Спецификация) <> Неопределено Или Стек.Количество() > 20 Тогда
		Возврат 1;                                                      // цикл вверх — его найдёт обход вниз
	КонецЕсли;
	Стек.Добавить(Спецификация);
	Узлы = Новый Массив;
	Узлы.Добавить(Спецификация);
	Запрос = Новый Запрос(ТекстЗапросаРодителей());
	Запрос.УстановитьПараметр("Узлы", Узлы);
	Выборка = Запрос.Выполнить().Выбрать();
	Уровень = 1;
	Пока Выборка.Следующий() Цикл
		Уровень = Макс(Уровень, УровеньСпецификации(Выборка.Спецификация, Кэш, Стек) + 1);
	КонецЦикла;
	Стек.Удалить(Стек.ВГраница());
	Кэш.Вставить(Спецификация, Уровень);
	Возврат Уровень;

КонецФункции

// Цикл и глубина (§4.6 слои 1–2, Т6). Корень — ссылка или объект (несохранённые строки). Обход вниз по ВСЕМ строкам-узлам
// (включая «купується готовим»); уровень корня — УровеньВДомах. Возврат: Неопределено или Структура(Вид — "Цикл"/"Глибина", Текст).
Функция НайтиЦикл(Корень) Экспорт

	Дерево = ЗагрузитьДеревоСпецификаций(Корень);
	ДанныеКорня = Дерево.Кэш.Получить(Дерево.КлючКорня);
	Если ДанныеКорня = Неопределено Тогда
		Возврат Неопределено;
	КонецЕсли;
	Если ДанныеКорня.ВидСпецификации = Перечисления.А_ВидыСпецификацийДомов.Дом Тогда
		УровеньКорня = 0;
	ИначеЕсли ТипЗнч(Дерево.КлючКорня) = Тип("Строка") Тогда
		УровеньКорня = 1;
	Иначе
		УровеньКорня = УровеньВДомах(Дерево.КлючКорня);
	КонецЕсли;
	Путь = Новый Массив;
	Путь.Добавить(Дерево.КлючКорня);
	Возврат ОбойтиДляПроверки(Дерево.Кэш, Дерево.КлючКорня, УровеньКорня, Путь);

КонецФункции

Функция ОбойтиДляПроверки(Кэш, КлючСпец, Уровень, Путь)

	Данные = Кэш.Получить(КлючСпец);
	Если Данные = Неопределено Тогда
		Возврат Неопределено;
	КонецЕсли;
	ВидУзел = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	Для Каждого Стр Из Данные.Строки Цикл
		Если Стр.ВидСтроки <> ВидУзел Или Не ЗначениеЗаполнено(Стр.Узел) Тогда
			Продолжить;
		КонецЕсли;
		Путь.Добавить(Стр.Узел);
		Если Путь.Найти(Стр.Узел) < Путь.ВГраница() Тогда
			Возврат Новый Структура("Вид, Текст", "Цикл",
				СтрШаблон(НСтр("ru = 'Узел входит сам в себя: %1'; uk = 'Вузол входить сам у себе: %1'"), ПредставлениеПути(Путь, Кэш)));
		КонецЕсли;
		Если Уровень + 1 >= 10 Тогда
			Возврат Новый Структура("Вид, Текст", "Глибина",
				СтрШаблон(НСтр("ru = 'Глубина вложенности узлов больше 9 (дом — уровень 0): %1'; uk = 'Глибина вкладеності вузлів більша за 9 (будинок — рівень 0): %1'"),
					ПредставлениеПути(Путь, Кэш)));
		КонецЕсли;
		Найдено = ОбойтиДляПроверки(Кэш, Стр.Узел, Уровень + 1, Путь);
		Если Найдено <> Неопределено Тогда
			Возврат Найдено;
		КонецЕсли;
		Путь.Удалить(Путь.ВГраница());
	КонецЦикла;
	Возврат Неопределено;

КонецФункции

// «Перевірити цикли» (слой 2, команда формы списка): все действующие и черновые дома и узлы по ТЕКУЩИМ данным.
// Возврат: Массив строк «Наименование: текст ошибки».
Функция ПроверитьЦиклы() Экспорт

	Ошибки = Новый Массив;
	Запрос = Новый Запрос(ТекстЗапросаСпецификацийДляПроверки());
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Найдено = НайтиЦикл(Выборка.Ссылка);
		Если Найдено <> Неопределено Тогда
			Ошибки.Добавить(Выборка.Наименование + ": " + Найдено.Текст);
		КонецЕсли;
	КонецЦикла;
	Возврат Ошибки;

КонецФункции

Функция ТекстЗапросаРодителей()

	Возврат
	//@ЗАПРОС roditeli
	;

КонецФункции

Функция ТекстЗапросаСпецификацийДляПроверки()

	Возврат
	//@ЗАПРОС specifikacii_dlya_proverki
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_30_cikly.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 28 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_30_cikly.bsl
git commit -m "feat(spec_domov): входимость узла, уровень в домах и поиск цикла/глубины (Т6) для записи и «Перевірити цикли»

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 23: Проверки при записи и при «Діє», `ПередЗаписью`, `ПриКопировании` (§4.6, Т9)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_32_proverki.bsl`
- Test: `test_e1_priemka.py::k1, k2, k4, k8, k9`

- [ ] **Step 1: Тест уже написан** (Task 19). **Step 2: FAIL** — нет модуля.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_32_proverki.bsl`:

```bsl
#Область ПроверкиСпецификации

// Из модуля объекта: ПередЗаписью → сюда (тонкий модуль, §4.1). Ставит ДатаИзменения, ключи строк, умолчания
// (Статус «Чернетка», КоличествоВыпуска 1, вид строки «Деталь», k по единицам) и выполняет проверки §4.6:
// «при записи» всегда, «при переводе в Діє» — только при переходе. Ошибки — сообщениями с Отказ = Истина.
Процедура ПередЗаписьюСпецификации(СпецОбъект, Отказ) Экспорт

	Если СпецОбъект.ЭтоГруппа Тогда
		Возврат;
	КонецЕсли;
	СпецОбъект.ДатаИзменения = ТекущаяДатаСеанса();
	Если Не ЗначениеЗаполнено(СпецОбъект.Статус) Тогда
		СпецОбъект.Статус = Перечисления.А_СтатусыСпецификацийДомов.Черновик;
	КонецЕсли;
	Если СпецОбъект.КоличествоВыпуска = 0 Тогда
		СпецОбъект.КоличествоВыпуска = 1;
	КонецЕсли;
	Если СпецОбъект.Статус = Перечисления.А_СтатусыСпецификацийДомов.Действует И Не ЗначениеЗаполнено(СпецОбъект.ДатаНачала) Тогда
		СпецОбъект.ДатаНачала = НачалоДня(ТекущаяДатаСеанса());
	КонецЕсли;
	ЗаполнитьКлючиСтрок(СпецОбъект.Состав);
	ЗаполнитьКлючиСтрок(СпецОбъект.ЗакупочныйСостав);
	Для Каждого Стр Из СпецОбъект.Состав Цикл
		Если Не ЗначениеЗаполнено(Стр.ВидСтроки) Тогда
			Стр.ВидСтроки = Перечисления.А_ВидыСтрокСпецификацииДомов.Деталь;
		КонецЕсли;
		Если Стр.Коэффициент = 0 Тогда
			УстановитьКоэффициентПоЕдиницам(Стр);
		КонецЕсли;
	КонецЦикла;
	Ошибки = ОшибкиПриЗаписи(СпецОбъект);
	Если ЭтоПереходВДействует(СпецОбъект) Тогда
		Для Каждого ТекстОшибки Из ОшибкиПереходаВДействует(СпецОбъект) Цикл
			Ошибки.Добавить(ТекстОшибки);
		КонецЦикла;
	КонецЕсли;
	Для Каждого ТекстОшибки Из Ошибки Цикл
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(ТекстОшибки, СпецОбъект, , , Отказ);
	КонецЦикла;

КонецПроцедуры

// Копия элемента (F9) — новая версия или вариант вручную: «Чернетка», Основание = исходная, СоставОснования = её состав
// (ключи строк при копировании сохраняются, §4.3–4.4).
Процедура ПриКопированииСпецификации(СпецОбъект, ОбъектКопирования) Экспорт

	Если ОбъектКопирования.ЭтоГруппа Тогда
		Возврат;
	КонецЕсли;
	СпецОбъект.Статус = Перечисления.А_СтатусыСпецификацийДомов.Черновик;
	СпецОбъект.ДатаНачала = Дата(1, 1, 1);
	СпецОбъект.Основание = ОбъектКопирования.Ссылка;
	СпецОбъект.СоставОснования.Загрузить(ОбъектКопирования.Состав.Выгрузить());

КонецПроцедуры

// Проверки «при записи» (§4.6, Т9). Работают и при программной записи и в COM-тестах (Записать() не вызывает
// ОбработкаПроверкиЗаполнения, E2 §6.3 п. 2). Возврат: Массив строк.
Функция ОшибкиПриЗаписи(СпецОбъект) Экспорт

	Ошибки = Новый Массив;
	Если СпецОбъект.ЭтоГруппа Тогда
		Возврат Ошибки;
	КонецЕсли;
	Виды = Перечисления.А_ВидыСпецификацийДомов;
	Вид = СпецОбъект.ВидСпецификации;
	Если Не ЗначениеЗаполнено(Вид) Тогда
		Ошибки.Добавить(НСтр("ru = 'Не заполнен вид спецификации.'; uk = 'Не заповнено вид специфікації.'"));
		Возврат Ошибки;
	КонецЕсли;
	Если (Вид = Виды.Дом Или Вид = Виды.Закупочная) И Не ЗначениеЗаполнено(СпецОбъект.СтруктураСебестоимости) Тогда
		Ошибки.Добавить(НСтр("ru = 'Не указан тип дома (СС).'; uk = 'Не вказано тип будинку (СС).'"));
	КонецЕсли;
	Если СпецОбъект.КоличествоВыпуска <= 0 Тогда
		Ошибки.Добавить(НСтр("ru = 'Количество выпуска должно быть больше 0.'; uk = 'Кількість випуску має бути більшою за 0.'"));
	ИначеЕсли Вид <> Виды.Узел И СпецОбъект.КоличествоВыпуска <> 1 Тогда
		Ошибки.Добавить(НСтр("ru = 'У дома и закупочной количество выпуска = 1.'; uk = 'У будинку й закупівельної кількість випуску = 1.'"));
	КонецЕсли;
	Если Вид = Виды.Узел И Не ЗначениеЗаполнено(СпецОбъект.ЕдиницаВыпуска) Тогда
		Ошибки.Добавить(НСтр("ru = 'У узла обязательна единица выпуска.'; uk = 'У вузла обовʼязкова одиниця випуску.'"));
	КонецЕсли;
	Если Вид = Виды.Закупочная Тогда
		ПроверитьЗакупочнуюПриЗаписи(СпецОбъект, Ошибки);
		Возврат Ошибки;
	КонецЕсли;
	ПроверитьСоставПриЗаписи(СпецОбъект, Ошибки);
	Найдено = НайтиЦикл(СпецОбъект);
	Если Найдено <> Неопределено Тогда
		Ошибки.Добавить(Найдено.Текст);
	КонецЕсли;
	Если Вид = Виды.Дом И СпецОбъект.Дома.Количество() > 0 Тогда
		ПроверитьДомаВариантов(СпецОбъект, Ошибки);
	КонецЕсли;
	Возврат Ошибки;

КонецФункции

// Проверки «при переводе в Діє» (§4.6). Возврат: Массив строк.
Функция ОшибкиПереходаВДействует(СпецОбъект) Экспорт

	Ошибки = Новый Массив;
	Если СпецОбъект.ВидСпецификации = Перечисления.А_ВидыСпецификацийДомов.Закупочная Тогда
		ПроверитьЗакупочнуюПередДействует(СпецОбъект, Ошибки);
	Иначе
		ПроверитьСоставПередДействует(СпецОбъект, Ошибки);
	КонецЕсли;
	Возврат Ошибки;

КонецФункции

Процедура ПроверитьСоставПриЗаписи(СпецОбъект, Ошибки)

	ВидыСтрок = Перечисления.А_ВидыСтрокСпецификацииДомов;
	Черновик = (СпецОбъект.Статус = Перечисления.А_СтатусыСпецификацийДомов.Черновик);
	ВариантДома = ЗначениеЗаполнено(СпецОбъект.Основание) И СпецОбъект.Дома.Количество() > 0;
	Для Каждого Стр Из СпецОбъект.Состав Цикл
		Префикс = СтрШаблон(НСтр("ru = 'Строка %1: '; uk = 'Рядок %1: '"), Стр.НомерСтроки);
		Если Стр.Количество <= 0 Тогда
			Ошибки.Добавить(Префикс + НСтр("ru = 'количество должно быть больше 0.'; uk = 'кількість має бути більшою за 0.'"));
		КонецЕсли;
		Если Стр.Коэффициент <= 0 Тогда
			Ошибки.Добавить(Префикс + НСтр("ru = 'коэффициент должен быть больше 0.'; uk = 'коефіцієнт має бути більшим за 0.'"));
		КонецЕсли;
		Если Стр.НеПоставляется И ПустаяСтрока(Стр.Комментарий) Тогда
			Ошибки.Добавить(Префикс + НСтр("ru = '«Не поставляется» — нужна причина в примечании.'; uk = '«Не постачається» — потрібна причина в примітці.'"));
		КонецЕсли;
		Если Стр.ВидСтроки = ВидыСтрок.Деталь Тогда
			Если Не ЗначениеЗаполнено(Стр.Номенклатура) И Не Стр.НеПоставляется И Не Черновик Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'деталь без карточки — только с «Не поставляется» и причиной (в «Черновике» допускается).'; uk = 'деталь без картки — лише з «Не постачається» і причиною (у «Чернетці» допускається).'"));
			КонецЕсли;
		ИначеЕсли Стр.ВидСтроки = ВидыСтрок.Узел Тогда
			ПроверитьСтрокуУзла(СпецОбъект, Стр, Префикс, Ошибки);
		ИначеЕсли Стр.ВидСтроки = ВидыСтрок.ГруппаПоОбщемуНазванию Тогда
			Если Не ЗначениеЗаполнено(Стр.ОбщееНазвание) Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'у группы обязательно общее название.'; uk = 'у групи обовʼязкова загальна назва.'"));
			КонецЕсли;
			Если ЗначениеЗаполнено(Стр.Номенклатура) Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'у группы карточки нет — очистите карточку.'; uk = 'у групи картки немає — очистіть картку.'"));
			КонецЕсли;
			Если Стр.Коэффициент <> 1 Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'у группы k = 1.'; uk = 'у групи k = 1.'"));
			КонецЕсли;
		КонецЕсли;
		Если ВариантДома И Стр.ВидСтроки <> ВидыСтрок.Узел И Не ЗначениеЗаполнено(Стр.ПозицияСС) И ПустаяСтрока(Стр.Комментарий) Тогда
			Ошибки.Добавить(Префикс + НСтр("ru = 'строка без позиции СС в варианте — «Добавлено в доме», примечание обязательно (Р27).'; uk = 'рядок без позиції СС у варіанті — «Додано в будинку», примітка обовʼязкова (Р27).'"));
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

Процедура ПроверитьСтрокуУзла(СпецОбъект, Стр, Префикс, Ошибки)

	Если Не ЗначениеЗаполнено(Стр.Узел) Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = 'не указан узел.'; uk = 'не вказано вузол.'"));
		Возврат;
	КонецЕсли;
	Если Стр.Узел = СпецОбъект.Ссылка Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = 'узел не может входить сам в себя.'; uk = 'вузол не може входити сам у себе.'"));
		Возврат;
	КонецЕсли;
	Реквизиты = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(Стр.Узел, "ВидСпецификации, Статус, ЕдиницаВыпуска, НоменклатураВузла");
	Если Реквизиты.ВидСпецификации <> Перечисления.А_ВидыСпецификацийДомов.Узел Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = 'ссылаться можно только на спецификацию вида «Узел».'; uk = 'посилатися можна лише на специфікацію виду «Вузол».'"));
	КонецЕсли;
	Если Реквизиты.Статус = Перечисления.А_СтатусыСпецификацийДомов.Закрыта Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = 'узел закрыт.'; uk = 'вузол закрито.'"));
	КонецЕсли;
	Если Не ЕдиницыСовпадают(Стр.ЕдиницаСтроки, Реквизиты.ЕдиницаВыпуска) Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = 'единица строки должна совпадать с единицей выпуска узла.'; uk = 'одиниця рядка має збігатися з одиницею випуску вузла.'"));
	КонецЕсли;
	Если Стр.Коэффициент <> 1 Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = 'у строки-узла k = 1.'; uk = 'у рядка-вузла k = 1.'"));
	КонецЕсли;
	Если Стр.ПокупаетсяГотовым И Не ЗначениеЗаполнено(Реквизиты.НоменклатураВузла) Тогда
		Ошибки.Добавить(Префикс + НСтр("ru = '«Покупается готовым» — у узла должна быть карточка узла.'; uk = '«Купується готовим» — у вузла має бути картка вузла.'"));
	КонецЕсли;

КонецПроцедуры

Процедура ПроверитьДомаВариантов(СпецОбъект, Ошибки)

	Если Не ЗначениеЗаполнено(СпецОбъект.СтруктураСебестоимости) Тогда
		Возврат;
	КонецЕсли;
	ПодразделениеОбъекта = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СпецОбъект.СтруктураСебестоимости, "ПодразделениеОбъекта");

	Дома = Новый Массив;
	Для Каждого СтрокаДома Из СпецОбъект.Дома Цикл
		Если Не ЗначениеЗаполнено(СтрокаДома.Подразделение) Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дома, строка %1: не указан дом.'; uk = 'Будинки, рядок %1: не вказано будинок.'"), СтрокаДома.НомерСтроки));
		ИначеЕсли ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СтрокаДома.Подразделение, "Родитель") <> ПодразделениеОбъекта Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дом %1 — не дом этого типа (СС).'; uk = 'Будинок %1 — не будинок цього типу (СС).'"), СтрокаДома.Подразделение));
		ИначеЕсли Дома.Найти(СтрокаДома.Подразделение) <> Неопределено Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дом %1 указан дважды.'; uk = 'Будинок %1 вказано двічі.'"), СтрокаДома.Подразделение));
		Иначе
			Дома.Добавить(СтрокаДома.Подразделение);
		КонецЕсли;
	КонецЦикла;
	Если СпецОбъект.Статус <> Перечисления.А_СтатусыСпецификацийДомов.Действует Или Дома.Количество() = 0 Тогда
		Возврат;
	КонецЕсли;
	Запрос = Новый Запрос(ТекстЗапросаДомовВВариантах());
	Запрос.УстановитьПараметр("Дома", Дома);
	Запрос.УстановитьПараметр("Ссылка", СпецОбъект.Ссылка);
	Запрос.УстановитьПараметр("СсылкаСС", СпецОбъект.СтруктураСебестоимости);
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дом %1 уже в действующем варианте «%2» этой СС.'; uk = 'Будинок %1 вже в діючому варіанті «%2» цієї СС.'"),
			Выборка.Подразделение, Выборка.Спецификация));
	КонецЦикла;

КонецПроцедуры

Процедура ПроверитьЗакупочнуюПриЗаписи(СпецОбъект, Ошибки)

	Карточки = Новый Соответствие;
	Для Каждого СтрокаЗакупки Из СпецОбъект.ЗакупочныйСостав Цикл
		Если Не ЗначениеЗаполнено(СтрокаЗакупки.Номенклатура) Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Закупочный состав, строка %1: не указана карточка.'; uk = 'Закупівельний склад, рядок %1: не вказано картку.'"),
				СтрокаЗакупки.НомерСтроки));
		ИначеЕсли Карточки.Получить(СтрокаЗакупки.Номенклатура) <> Неопределено Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Закупочный состав: карточка %1 повторяется (строки %2 и %3).'; uk = 'Закупівельний склад: картка %1 повторюється (рядки %2 і %3).'"),
				СтрокаЗакупки.Номенклатура, Карточки.Получить(СтрокаЗакупки.Номенклатура), СтрокаЗакупки.НомерСтроки));
		Иначе
			Карточки.Вставить(СтрокаЗакупки.Номенклатура, СтрокаЗакупки.НомерСтроки);
		КонецЕсли;
	КонецЦикла;
	// §4.5: дом факта — дочерний к СС.ПодразделениеОбъекта, не повторяется.
	ПодразделениеОбъекта = Неопределено;
	Если ЗначениеЗаполнено(СпецОбъект.СтруктураСебестоимости) Тогда
		ПодразделениеОбъекта = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СпецОбъект.СтруктураСебестоимости, "ПодразделениеОбъекта");
	КонецЕсли;
	Дома = Новый Соответствие;
	Для Каждого СтрокаДома Из СпецОбъект.ДомаФакта Цикл
		Если Не ЗначениеЗаполнено(СтрокаДома.Подразделение) Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дома факта, строка %1: не указан дом.'; uk = 'Будинки факту, рядок %1: не вказано будинок.'"),
				СтрокаДома.НомерСтроки));
			Продолжить;
		КонецЕсли;
		Если ПодразделениеОбъекта <> Неопределено
			И ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СтрокаДома.Подразделение, "Родитель") <> ПодразделениеОбъекта Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дома факта: %1 — не дом этого типа (СС).'; uk = 'Будинки факту: %1 — не будинок цього типу (СС).'"),
				СтрокаДома.Подразделение));
		КонецЕсли;
		Если Дома.Получить(СтрокаДома.Подразделение) <> Неопределено Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Дома факта: дом %1 указан дважды.'; uk = 'Будинки факту: будинок %1 вказано двічі.'"), СтрокаДома.Подразделение));
		КонецЕсли;
		Дома.Вставить(СтрокаДома.Подразделение, Истина);
	КонецЦикла;

КонецПроцедуры

Процедура ПроверитьСоставПередДействует(СпецОбъект, Ошибки)

	ВидыСтрок = Перечисления.А_ВидыСтрокСпецификацииДомов;
	Если СпецОбъект.Состав.Количество() = 0 Тогда
		Ошибки.Добавить(НСтр("ru = 'Состав пуст.'; uk = 'Склад порожній.'"));
		Возврат;
	КонецЕсли;
	ЭтоДом = (СпецОбъект.ВидСпецификации = Перечисления.А_ВидыСпецификацийДомов.Дом);
	ВариантДома = ЗначениеЗаполнено(СпецОбъект.Основание) И СпецОбъект.Дома.Количество() > 0;
	ЕстьСС = ЗначениеЗаполнено(СпецОбъект.СтруктураСебестоимости);
	ПозицииСС = ПозицииСтруктурыСебестоимости(СпецОбъект.СтруктураСебестоимости);
	НеНайден = Перечисления.А_ИсточникиКоэффициентаПересчетаСС.НеНайден;
	Для Каждого Стр Из СпецОбъект.Состав Цикл
		Префикс = СтрШаблон(НСтр("ru = 'Строка %1: '; uk = 'Рядок %1: '"), Стр.НомерСтроки);
		Если Не Стр.НеПоставляется Тогда
			Если (Стр.ВидСтроки = ВидыСтрок.Деталь Или Стр.ВидСтроки = ВидыСтрок.Узел) И Не Стр.Проверено Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'не отмечено «Проверено».'; uk = 'не позначено «Перевірено».'"));
			КонецЕсли;
			Если Стр.ВидСтроки = ВидыСтрок.Деталь И Не ЗначениеЗаполнено(Стр.Номенклатура) Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'деталь без карточки — только с «Не поставляется» и причиной.'; uk = 'деталь без картки — лише з «Не постачається» і причиною.'"));
			ИначеЕсли Стр.ВидСтроки = ВидыСтрок.Деталь И Стр.ИсточникКоэффициента = НеНайден Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'коэффициент «Не найден» — укажите 1 ед. карточки = k ед. строки.'; uk = 'коефіцієнт «Не знайдено» — вкажіть 1 од. картки = k од. рядка.'"));
			КонецЕсли;
		КонецЕсли;
		Если ЗначениеЗаполнено(Стр.ПозицияСС) И ЕстьСС Тогда
			Позиция = ПозицииСС.Получить(А_СтруктураСебестоимостиСервер.КлючПозиции(Стр.ЭтапСС, Стр.ПозицияСС));
			Если Позиция = Неопределено Тогда
				Ошибки.Добавить(Префикс + СтрШаблон(НСтр("ru = 'позиции «%1» нет в СС — перепривяжите позиции СС.'; uk = 'позиції «%1» немає в СС — перепривʼяжіть позиції СС.'"), Стр.ПозицияСС));
			ИначеЕсли Позиция.Количество = 0 И Стр.ЦенаНормы = 0 И Не Стр.НеПоставляется Тогда
				Ошибки.Добавить(Префикс + НСтр("ru = 'позиция СС нулевая — нужна цена нормы, введённая вручную (Р22).'; uk = 'позиція СС нульова — потрібна ціна норми, введена вручну (Р22).'"));
			КонецЕсли;
		ИначеЕсли ЭтоДом И Стр.ВидСтроки <> ВидыСтрок.Узел И Не ЗначениеЗаполнено(Стр.ПозицияСС)
			И Не (ВариантДома И Не ПустаяСтрока(Стр.Комментарий)) Тогда
			Ошибки.Добавить(Префикс + НСтр("ru = 'строка без позиции СС (в варианте — «Добавлено в доме» с примечанием).'; uk = 'рядок без позиції СС (у варіанті — «Додано в будинку» з приміткою).'"));
		КонецЕсли;
	КонецЦикла;
	Дерево = ЗагрузитьДеревоСпецификаций(СпецОбъект);
	Для Каждого ЭлементКэша Из Дерево.Кэш Цикл
		Если ЭлементКэша.Ключ <> Дерево.КлючКорня
			И ЭлементКэша.Значение.Статус <> Перечисления.А_СтатусыСпецификацийДомов.Действует Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Узел «%1» не в статусе «Действует».'; uk = 'Вузол «%1» не в статусі «Діє».'"),
				ЭлементКэша.Значение.Наименование));
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

Процедура ПроверитьЗакупочнуюПередДействует(СпецОбъект, Ошибки)

	Если Не ЗначениеЗаполнено(СпецОбъект.Основание)
		Или ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СпецОбъект.Основание, "Статус") <> Перечисления.А_СтатусыСпецификацийДомов.Действует Тогда
		Ошибки.Добавить(НСтр("ru = 'Основание (эталон типа) должно быть в статусе «Действует».'; uk = 'Основа (еталон типу) має бути в статусі «Діє».'"));
	КонецЕсли;
	Если СпецОбъект.ДомаФакта.Количество() = 0 Тогда
		Ошибки.Добавить(НСтр("ru = 'Не заполнены «Дома факта».'; uk = 'Не заповнено «Будинки факту».'"));
	КонецЕсли;
	Если СпецОбъект.ЗакупочныйСостав.Количество() = 0 Тогда
		Ошибки.Добавить(НСтр("ru = 'Закупочный состав пуст.'; uk = 'Закупівельний склад порожній.'"));
	КонецЕсли;
	Для Каждого СтрокаЗакупки Из СпецОбъект.ЗакупочныйСостав Цикл
		Если Не СтрокаЗакупки.Проверено Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Закупочный состав, строка %1: не проверено.'; uk = 'Закупівельний склад, рядок %1: не перевірено.'"), СтрокаЗакупки.НомерСтроки));
		КонецЕсли;
		Если СтрокаЗакупки.Количество = 0 Тогда
			Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Закупочный состав, строка %1: количество 0 — удалите строку или укажите количество.'; uk = 'Закупівельний склад, рядок %1: кількість 0 — видаліть рядок або вкажіть кількість.'"), СтрокаЗакупки.НомерСтроки));
		КонецЕсли;
	КонецЦикла;
	Запрос = Новый Запрос(ТекстЗапросаДругихЗакупочных());
	Запрос.УстановитьПараметр("Ссылка", СпецОбъект.Ссылка);
	Запрос.УстановитьПараметр("СсылкаСС", СпецОбъект.СтруктураСебестоимости);
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Ошибки.Добавить(СтрШаблон(НСтр("ru = 'Уже действует закупочная «%1» этой СС — сначала закройте её (одна действующая на СС).'; uk = 'Вже діє закупівельна «%1» цієї СС — спочатку закрийте її (одна діюча на СС).'"), Выборка.Ссылка));
	КонецЦикла;

КонецПроцедуры

Функция ЭтоПереходВДействует(СпецОбъект)

	Действует = Перечисления.А_СтатусыСпецификацийДомов.Действует;
	Если СпецОбъект.Статус <> Действует Тогда
		Возврат Ложь;
	ИначеЕсли СпецОбъект.ЭтоНовый() Тогда
		Возврат Истина;
	КонецЕсли;
	Возврат ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СпецОбъект.Ссылка, "Статус") <> Действует;

КонецФункции

Функция ТекстЗапросаДомовВВариантах()

	Возврат
	//@ЗАПРОС doma_v_variantah
	;

КонецФункции

Функция ТекстЗапросаДругихЗакупочных()

	Возврат
	//@ЗАПРОС drugie_zakupochnye
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_32_proverki.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 28 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_32_proverki.bsl
git commit -m "feat(spec_domov): проверки §4.6 при записи и при «Діє», ПередЗаписью/ПриКопировании (Т9)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 24: «Зробити еталоном для СС» без ПередЗаписью СС и без регистрации обмена (Р3, Б1)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_35_etalon.bsl`
- Test: `test_e1_priemka.py::k6_b1`, `::k7_etalon_only_dije`

- [ ] **Step 1: Тест уже написан** (Task 19). **Step 2: FAIL** — нет модуля.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_35_etalon.bsl`:

```bsl
#Область ЭталонСС

// «Зробити еталоном для СС» (Р3, §4.7): СС.ЭталоннаяСпецификация = Спецификация, НЕ выполняя ПередЗаписью СС
// (строка 34 модуля СС стирает Комплектующие.Номенклатура, Б1) и НЕ регистрируя СС к обмену с ERP: флаг
// ОтключитьМеханизмРегистрацииОбъектов (ОбменДаннымиСобытия, стр. 897); одна Загрузка регистрацию НЕ отключает
// (память buhbud_zapis_bez_registracii_obmena). Условия: вид «Будинок», статус «Діє», ТЧ «Будинки» пуста.
// Ошибка — исключение с текстом. Возврат: Истина.
Функция СделатьЭталономСС(Спецификация) Экспорт

	Реквизиты = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(Спецификация, "ВидСпецификации, Статус, СтруктураСебестоимости");
	Если Реквизиты.ВидСпецификации <> Перечисления.А_ВидыСпецификацийДомов.Дом Тогда
		ВызватьИсключение НСтр("ru = 'Эталоном может быть только спецификация вида «Дом».'; uk = 'Еталоном може бути лише специфікація виду «Будинок».'");
	КонецЕсли;
	Если Реквизиты.Статус <> Перечисления.А_СтатусыСпецификацийДомов.Действует Тогда
		ВызватьИсключение НСтр("ru = 'Сначала переведите спецификацию в «Действует».'; uk = 'Спочатку переведіть специфікацію в «Діє».'");
	КонецЕсли;
	Если Спецификация.Дома.Количество() > 0 Тогда
		ВызватьИсключение НСтр("ru = 'Это вариант (заполнены «Дома») — эталоном может быть только спецификация без домов.'; uk = 'Це варіант (заповнено «Будинки») — еталоном може бути лише специфікація без будинків.'");
	КонецЕсли;
	СсылкаСС = Реквизиты.СтруктураСебестоимости;
	Если Не ЗначениеЗаполнено(СсылкаСС) Тогда
		ВызватьИсключение НСтр("ru = 'Не указан тип дома (СС).'; uk = 'Не вказано тип будинку (СС).'");
	КонецЕсли;

	НачатьТранзакцию();
	Попытка
		Блокировка = Новый БлокировкаДанных;
		ЭлементБлокировки = Блокировка.Добавить("Справочник.СтруктураСебестоимости");
		ЭлементБлокировки.УстановитьЗначение("Ссылка", СсылкаСС);
		Блокировка.Заблокировать();
		ОбъектСС = СсылкаСС.ПолучитьОбъект();
		ОбъектСС.ЭталоннаяСпецификация = Спецификация;
		ОбъектСС.ОбменДанными.Загрузка = Истина;
		ОбъектСС.ДополнительныеСвойства.Вставить("ОтключитьМеханизмРегистрацииОбъектов", Истина);
		ОбъектСС.Записать();
		ЗафиксироватьТранзакцию();
	Исключение
		ОтменитьТранзакцию();
		ВызватьИсключение;
	КонецПопытки;
	Возврат Истина;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_35_etalon.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 28 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_35_etalon.bsl
git commit -m "feat(spec_domov): «Зробити еталоном для СС» — запись СС без ПередЗаписью и без регистрации обмена (Р3, Б1)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 25: Общий модуль (метаданные) и модуль объекта справочника

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/meta/e1_module.json`
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/obj_module.bsl`
- Create (генерируются): `_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер.xml`, `…/А_СпецификацииДомовСервер/Ext/Module.bsl`, `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl`

- [ ] **Step 1: Тест** — проверка флагов модуля в `test_e1_files.py` дописывается разделом `module`:

Добавить в `test_e1_files.py` перед блоком `if __name__ == "__main__":`

```python
def check_module():
    t = load("CommonModules/А_СпецификацииДомовСервер.xml")
    clean_new("CommonModules/А_СпецификацииДомовСервер.xml", t)
    for tag, val in (("Server", "true"), ("ExternalConnection", "true"), ("ServerCall", "false"), ("Privileged", "false"),
                     ("ClientManagedApplication", "false"), ("Global", "false")):
        assert f"<{tag}>{val}</{tag}>" in t, (tag, val)
    m = open(os.path.join(BB_WT, "CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"), encoding="utf-8-sig").read()
    for fn in ("Функция ХешMD5(", "Функция РазвернутьДоКарточек(", "Функция НайтиЦикл(", "Процедура ПередЗаписьюСпецификации(",
               "Функция СделатьЭталономСС(", "Функция Входимость(", "Функция ПроверитьЦиклы("):
        assert fn in m, fn
    o = open(os.path.join(BB_WT, "Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl"), encoding="utf-8-sig").read()
    assert "А_СпецификацииДомовСервер.ПередЗаписьюСпецификации(ЭтотОбъект, Отказ)" in o
    print("module OK")
```

и в словарь `parts` — `"module": check_module`.

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only module`
Expected: FAIL — `нет файла CommonModules/А_СпецификацииДомовСервер.xml`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/meta/e1_module.json`:

```json
{"type": "CommonModule", "name": "А_СпецификацииДомовСервер", "synonym": "Спецификации домов (сервер)",
 "server": true, "externalConnection": true}
```

`_Rarzrabotki/Python/test/spec_domov/bsl/obj_module.bsl`:

```bsl
#Область ОбработчикиСобытий

Процедура ПередЗаписью(Отказ)

	Если ОбменДанными.Загрузка Тогда
		Возврат;
	КонецЕсли;
	А_СпецификацииДомовСервер.ПередЗаписьюСпецификации(ЭтотОбъект, Отказ);

КонецПроцедуры

Процедура ПриКопировании(ОбъектКопирования)

	А_СпецификацииДомовСервер.ПриКопированииСпецификации(ЭтотОбъект, ОбъектКопирования);

КонецПроцедуры

#КонецОбласти
```

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/gen_meta.py meta/e1_module.json
python -X utf8 _Rarzrabotki/Python/test/spec_domov/build_bsl.py --partiya 1
python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py "CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl" "Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl" "CommonModules/А_СпецификацииДомовСервер.xml"
```
Expected: `создан CommonModules/А_СпецификацииДомовСервер.xml`; `Module.bsl: частей 6, строк ≈ 700`; `ObjectModule.bsl`; `ошибок: 0` (собранные файлы — с BOM и CRLF, запросы развёрнуты).

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only module`
Expected: `module OK`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/meta/e1_module.json _Rarzrabotki/Python/test/spec_domov/bsl/obj_module.bsl _Rarzrabotki/Python/test/spec_domov/test_e1_files.py \
  "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер.xml" "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl" \
  "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl"
git commit -m "feat(BASEBuh): общий модуль А_СпецификацииДомовСервер (Сервер + ВнешнееСоединение) и тонкий модуль объекта

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 26: Генератор форм `gen_forms.py` и формы этапа 1

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/gen_forms.py`
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_1.bsl`, `_Rarzrabotki/Python/test/spec_domov/bsl/forma_spiska_1.bsl`
- Create (генерируются): `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl`; `…/Forms/ФормаСписка.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl`; `…/Forms/ФормаГруппы.xml` + `/Ext/Form.xml`
- Modify: `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml` (`DefaultObjectForm`, `DefaultListForm`, `DefaultFolderForm`, `<Form>…</Form>`)
- Test: `test_e1_files.py --only forms`

`form-compile` для BASEBuh непригоден (E1 §7.3: версия 2.17, только ru, `Autofill=false`), поэтому формы строит свой генератор по образцу `PT/mk9_formy_dovidnykiv.py` с порядком свойств, выверенным по формам BASEBuh. В BuhBud нет модуля `СобытияФорм` (E1 §3.4) — БСП-события форм не вызываются; командная панель формы автозаполняемая («единое окно»).

- [ ] **Step 1: Написать падающий тест**

Добавить в `test_e1_files.py` перед `if __name__ == "__main__":`

```python
def check_forms():
    cat = load("Catalogs/А_СпецификацииДомов.xml")
    for tag, form in (("DefaultObjectForm", "ФормаЭлемента"), ("DefaultListForm", "ФормаСписка"),
                      ("DefaultFolderForm", "ФормаГруппы")):
        assert f"<{tag}>Catalog.А_СпецификацииДомов.Form.{form}</{tag}>" in cat, tag
        assert cat.count(f"<Form>{form}</Form>") == 1, form
        md = load(f"Catalogs/А_СпецификацииДомов/Forms/{form}.xml")
        clean_new(form, md)
        fx = load(f"Catalogs/А_СпецификацииДомов/Forms/{form}/Ext/Form.xml")
        clean_new(form + "/Form.xml", fx)
        assert 'version="2.13"' in fx and "<Action xsi:type" not in fx
    el = load("Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента/Ext/Form.xml")
    for name in ("Состав", "Дома", "ДеревоУзлов", "РазвернутыйСостав", "СравнениеСС", "ЗакупочныйСостав", "ДомаФакта", "ФактПоДомам"):
        assert f'<Table name="{name}"' in el, name
    for cmd in ("СделатьЭталономСС", "ПоказатьВходимость", "ОбновитьДерево", "РазвернутьДоДеталей"):
        assert f'<Command name="{cmd}"' in el and f"<Action>{cmd}</Action>" in el, cmd
    sp = load("Catalogs/А_СпецификацииДомов/Forms/ФормаСписка/Ext/Form.xml")
    assert "<ManualQuery>true</ManualQuery>" in sp and "Список.ЭталонДляСС" in sp and '<Command name="ПроверитьЦиклы"' in sp
    for mod in ("ФормаЭлемента", "ФормаСписка"):
        m = open(os.path.join(BB_WT, f"Catalogs/А_СпецификацииДомов/Forms/{mod}/Ext/Form/Module.bsl"), "rb").read()
        assert m.startswith(b"\xef\xbb\xbf") and b"\r\n" in m, mod
    print("forms OK")
```

и `"forms": check_forms` в словарь `parts`.

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only forms`
Expected: FAIL — `AssertionError: DefaultObjectForm`.

- [ ] **Step 3: Реализовать генератор и модули форм этапа 1**

`_Rarzrabotki/Python/test/spec_domov/gen_forms.py`:

```python
# -*- coding: utf-8 -*-
"""Формы справочника А_СпецификацииДомов (П1) в дамп WORKTREE.
--etap 1: ФормаЭлемента (страницы «Склад», «Будинки», «Дерево», «Розгорнутий склад», «Порівняння», «Закупівля»,
          «Факт по будинках»; команды этапа 1), ФормаСписка («Еталон для СС», «Перевірити цикли»), ФормаГруппы.
--etap 2: то же с командами этапа 2 + ФормаПодбораКарточки (Р35) + ФормаПерепривязки (§5 шаг 7).
Модули форм — из spec_domov/bsl/forma_*.bsl (накопительно). UUID обёрток форм — uuid5: перегенерация их не меняет
(память raschet_komplektacij_dokument). Порядок свойств полей и таблиц — канон BASEBuh (E1 §3).
Запуск: python gen_forms.py --etap 1|2"""
import argparse
import os
import re
import sys
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_bsl import read_src  # noqa: E402
from common import BB_WT, read_text, write_text, zapros  # noqa: E402

CAT = "А_СпецификацииДомов"
NS = uuid.UUID("3f6a1c52-7d0e-4b8a-9c21-5e4d2b7a9f10")
XNS = {"app": "http://v8.1c.ru/8.2/managed-application/core", "cfg": "http://v8.1c.ru/8.1/data/enterprise/current-config",
      "cmi": "http://v8.1c.ru/8.2/managed-application/cmi", "dcscor": "http://v8.1c.ru/8.1/data-composition-system/core",
      "dcsset": "http://v8.1c.ru/8.1/data-composition-system/settings", "ent": "http://v8.1c.ru/8.1/data/enterprise",
      "lf": "http://v8.1c.ru/8.2/managed-application/logform", "style": "http://v8.1c.ru/8.1/data/ui/style",
      "sys": "http://v8.1c.ru/8.1/data/ui/fonts/system", "v8": "http://v8.1c.ru/8.1/data/core", "v8ui": "http://v8.1c.ru/8.1/data/ui",
      "web": "http://v8.1c.ru/8.1/data/ui/colors/web", "win": "http://v8.1c.ru/8.1/data/ui/colors/windows",
      "xen": "http://v8.1c.ru/8.3/xcf/enums", "xpr": "http://v8.1c.ru/8.3/xcf/predef", "xr": "http://v8.1c.ru/8.3/xcf/readable",
      "xs": "http://www.w3.org/2001/XMLSchema", "xsi": "http://www.w3.org/2001/XMLSchema-instance"}


def xmlns(*prefixes):
    """Объявления в каноническом порядке выгрузки платформы — по алфавиту префиксов (образец А_ТехкартаКМД/Forms)."""
    return " ".join(f'xmlns:{p}="{XNS[p]}"' for p in sorted(prefixes))


FORM_HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n<Form xmlns="http://v8.1c.ru/8.3/xcf/logform" '
             + xmlns("app", "cfg", "dcscor", "dcsset", "ent", "lf", "style", "sys", "v8", "v8ui", "web", "win", "xr", "xs", "xsi")
             + ' version="2.13">\n')
MD_HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" '
           + xmlns("app", "cfg", "cmi", "ent", "lf", "style", "sys", "v8", "v8ui", "web", "win", "xen", "xpr", "xr", "xs", "xsi")
           + ' version="2.13">\n')


def uid(key):
    return str(uuid.uuid5(NS, key))


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def ml(tag, ru, uk, t, attrs=""):
    return (f"{t}<{tag}{attrs}>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{esc(ru)}</v8:content>\n"
            f"{t}\t</v8:item>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>uk</v8:lang>\n{t}\t\t<v8:content>{esc(uk)}</v8:content>\n"
            f"{t}\t</v8:item>\n{t}</{tag}>\n")


class Ids:
    def __init__(self):
        self.n = 1

    def __call__(self):
        self.n += 1
        return self.n - 1


def events_xml(events, t):
    if not events:
        return ""
    return f"{t}<Events>\n" + "".join(f'{t}\t<Event name="{e}">{h}</Event>\n' for e, h in events) + f"{t}</Events>\n"


def field(ids, name, path, t, kind="InputField", title=None, ro=False, visible=True, in_table=False, width=None,
          height=None, stretch=False, multiline=False, footer=None, checkbox_type=False, title_loc=None, events=None):
    """Порядок (по 518 формам BASEBuh): DataPath, Visible, ReadOnly, Title, TitleLocation, EditMode, CheckBoxType,
    FooterDataPath, Width, Height, HorizontalStretch, MultiLine, ContextMenu, ExtendedTooltip, Events."""
    s = f'{t}<{kind} name="{name}" id="{ids()}">\n{t}\t<DataPath>{path}</DataPath>\n'
    if not visible:
        s += f"{t}\t<Visible>false</Visible>\n"
    if ro:
        s += f"{t}\t<ReadOnly>true</ReadOnly>\n"
    if title:
        s += ml("Title", title[0], title[1], t + "\t")
    if title_loc:
        s += f"{t}\t<TitleLocation>{title_loc}</TitleLocation>\n"
    if in_table:
        s += f"{t}\t<EditMode>EnterOnInput</EditMode>\n"
    if checkbox_type:
        s += f"{t}\t<CheckBoxType>Auto</CheckBoxType>\n"
    if footer:
        s += f"{t}\t<FooterDataPath>{footer}</FooterDataPath>\n"
    if width:
        s += f"{t}\t<Width>{width}</Width>\n"
    if height:
        s += f"{t}\t<Height>{height}</Height>\n"
    if stretch:
        s += f"{t}\t<HorizontalStretch>true</HorizontalStretch>\n"
    if multiline:
        s += f"{t}\t<MultiLine>true</MultiLine>\n"
    s += f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
    return s + events_xml(events, t + "\t") + f"{t}</{kind}>\n"


def group(ids, name, children, t, orient="Horizontal"):
    s = (f'{t}<UsualGroup name="{name}" id="{ids()}">\n{t}\t<Group>{orient}</Group>\n{t}\t<Behavior>Usual</Behavior>\n'
         f"{t}\t<Representation>None</Representation>\n{t}\t<ShowTitle>false</ShowTitle>\n"
         f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}\t<ChildItems>\n')
    s += "".join(c(ids, t + "\t\t") for c in children)
    return s + f"{t}\t</ChildItems>\n{t}</UsualGroup>\n"


def button(ids, name, command, t):
    return (f'{t}<Button name="{name}" id="{ids()}">\n{t}\t<Type>CommandBarButton</Type>\n'
            f"{t}\t<CommandName>Form.Command.{command}</CommandName>\n"
            f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}</Button>\n')


def additions(ids, name, t):
    s = ""
    for tag, suffix, typ in (("SearchStringAddition", "СтрокаПоиска", "SearchStringRepresentation"),
                             ("ViewStatusAddition", "СостояниеПросмотра", "ViewStatusRepresentation"),
                             ("SearchControlAddition", "УправлениеПоиском", "SearchControl")):
        n = name + suffix
        s += (f'{t}<{tag} name="{n}" id="{ids()}">\n{t}\t<AdditionSource>\n{t}\t\t<Item>{name}</Item>\n'
              f"{t}\t\t<Type>{typ}</Type>\n{t}\t</AdditionSource>\n"
              f'{t}\t<ContextMenu name="{n}КонтекстноеМеню" id="{ids()}"/>\n'
              f'{t}\t<ExtendedTooltip name="{n}РасширеннаяПодсказка" id="{ids()}"/>\n{t}</{tag}>\n')
    return s


def table(ids, name, path, columns, t, buttons=(), events=None, ro=False, tree=False, change_rows=True, height=None):
    """Порядок: Representation, ReadOnly, ChangeRowSet, ChangeRowOrder, Height, AutoInsertNewRow, InitialTreeView,
    EnableStartDrag, EnableDrag, DataPath, RowFilter, ContextMenu, AutoCommandBar, ExtendedTooltip, дополнения, Events, ChildItems."""
    s = f'{t}<Table name="{name}" id="{ids()}">\n{t}\t<Representation>{"Tree" if tree else "List"}</Representation>\n'
    if ro:
        s += f"{t}\t<ReadOnly>true</ReadOnly>\n"
    if not change_rows:
        s += f"{t}\t<ChangeRowSet>false</ChangeRowSet>\n{t}\t<ChangeRowOrder>false</ChangeRowOrder>\n"
    if height:
        s += f"{t}\t<Height>{height}</Height>\n"
    if change_rows:
        s += f"{t}\t<AutoInsertNewRow>true</AutoInsertNewRow>\n"
    if tree:
        s += f"{t}\t<InitialTreeView>ExpandAllLevels</InitialTreeView>\n"
    s += f"{t}\t<EnableStartDrag>true</EnableStartDrag>\n{t}\t<EnableDrag>true</EnableDrag>\n{t}\t<DataPath>{path}</DataPath>\n"
    if not tree:
        s += f'{t}\t<RowFilter xsi:nil="true"/>\n'
    s += f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
    if buttons:
        s += f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ids()}">\n{t}\t\t<ChildItems>\n'
        s += "".join(button(ids, f"{name}{cmd}", cmd, t + "\t\t\t") for cmd in buttons)
        s += f"{t}\t\t</ChildItems>\n{t}\t</AutoCommandBar>\n"
    else:
        s += f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ids()}"/>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
    s += additions(ids, name, t + "\t") + events_xml(events, t + "\t")
    s += f"{t}\t<ChildItems>\n" + "".join(c(ids, t + "\t\t") for c in columns) + f"{t}\t</ChildItems>\n"
    return s + f"{t}</Table>\n"


def dyn_table(ids, name, columns, t, user_group):
    """Канон ОбщиеНазванияНоменклатуры.ФормаСписка (E1 §3.6): панель таблицы скрыта, команды — в панели формы."""
    s = (f'{t}<Table name="{name}" id="{ids()}">\n{t}\t<CommandBarLocation>None</CommandBarLocation>\n'
         f"{t}\t<DefaultItem>true</DefaultItem>\n{t}\t<UseAlternationRowColor>true</UseAlternationRowColor>\n"
         f"{t}\t<InitialTreeView>ExpandTopLevel</InitialTreeView>\n{t}\t<EnableStartDrag>true</EnableStartDrag>\n"
         f"{t}\t<EnableDrag>true</EnableDrag>\n{t}\t<DataPath>{name}</DataPath>\n"
         f"{t}\t<RowPictureDataPath>{name}.DefaultPicture</RowPictureDataPath>\n"
         f"{t}\t<AutoRefresh>false</AutoRefresh>\n{t}\t<AutoRefreshPeriod>60</AutoRefreshPeriod>\n{t}\t<Period>\n"
         f'{t}\t\t<v8:variant xsi:type="v8:StandardPeriodVariant">Custom</v8:variant>\n'
         f"{t}\t\t<v8:startDate>0001-01-01T00:00:00</v8:startDate>\n{t}\t\t<v8:endDate>0001-01-01T00:00:00</v8:endDate>\n"
         f"{t}\t</Period>\n{t}\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>\n{t}\t<RestoreCurrentRow>false</RestoreCurrentRow>\n"
         f'{t}\t<TopLevelParent xsi:nil="true"/>\n{t}\t<ShowRoot>true</ShowRoot>\n{t}\t<AllowRootChoice>false</AllowRootChoice>\n'
         f"{t}\t<UpdateOnDataChange>Auto</UpdateOnDataChange>\n{t}\t<UserSettingsGroup>{user_group}</UserSettingsGroup>\n"
         f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
         f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ids()}">\n{t}\t\t<Autofill>false</Autofill>\n{t}\t</AutoCommandBar>\n'
         f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n')
    s += additions(ids, name, t + "\t")
    s += f"{t}\t<ChildItems>\n" + "".join(c(ids, t + "\t\t") for c in columns) + f"{t}\t</ChildItems>\n"
    return s + f"{t}</Table>\n"


def user_settings_group(ids, name, t):
    return (f'{t}<UsualGroup name="{name}" id="{ids()}">\n' + ml("Title", "Группа пользовательских настроек",
            "Група користувацьких налаштувань", t + "\t") +
            f"{t}\t<VerticalStretch>false</VerticalStretch>\n{t}\t<Group>Vertical</Group>\n{t}\t<Behavior>Usual</Behavior>\n"
            f'{t}\t<ShowTitle>false</ShowTitle>\n{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
            f"{t}</UsualGroup>\n")


def page(ids, name, ru, uk, children, t):
    s = f'{t}<Page name="{name}" id="{ids()}">\n' + ml("Title", ru, uk, t + "\t")
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}\t<ChildItems>\n'
    s += "".join(c(ids, t + "\t\t") for c in children)
    return s + f"{t}\t</ChildItems>\n{t}</Page>\n"


def pages(ids, name, children, t):
    s = (f'{t}<Pages name="{name}" id="{ids()}">\n{t}\t<PagesRepresentation>TabsOnTop</PagesRepresentation>\n'
         f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}\t<ChildItems>\n')
    s += "".join(c(ids, t + "\t\t") for c in children)
    return s + f"{t}\t</ChildItems>\n{t}</Pages>\n"


def type_xml(kind, t):
    if kind.startswith("cfg:"):
        return f"{t}<v8:Type>{kind}</v8:Type>\n"
    if kind == "bool":
        return f"{t}<v8:Type>xs:boolean</v8:Type>\n"
    if kind == "uuid":
        return f"{t}<v8:Type>v8:UUID</v8:Type>\n"
    if kind == "vt":
        return f"{t}<v8:Type>v8:ValueTable</v8:Type>\n"
    if kind == "vtree":
        return f"{t}<v8:Type>v8:ValueTree</v8:Type>\n"
    if kind.startswith("str:"):
        return (f"{t}<v8:Type>xs:string</v8:Type>\n{t}<v8:StringQualifiers>\n{t}\t<v8:Length>{kind[4:]}</v8:Length>\n"
                f"{t}\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{t}</v8:StringQualifiers>\n")
    if kind.startswith("num:"):
        d, f = kind[4:].split(",")
        return (f"{t}<v8:Type>xs:decimal</v8:Type>\n{t}<v8:NumberQualifiers>\n{t}\t<v8:Digits>{d}</v8:Digits>\n"
                f"{t}\t<v8:FractionDigits>{f}</v8:FractionDigits>\n{t}\t<v8:AllowedSign>Any</v8:AllowedSign>\n{t}</v8:NumberQualifiers>\n")
    raise ValueError(kind)


def attr(aid, name, kind, t, title=None, columns=None, main=False, saved=False):
    s = f'{t}<Attribute name="{name}" id="{aid}">\n'
    if title:
        s += ml("Title", title[0], title[1], t + "\t")
    s += f"{t}\t<Type>\n" + type_xml(kind, t + "\t\t") + f"{t}\t</Type>\n"
    if main:
        s += f"{t}\t<MainAttribute>true</MainAttribute>\n"
    if saved:
        s += f"{t}\t<SavedData>true</SavedData>\n"
    if columns:
        s += f"{t}\t<Columns>\n"
        for cid, (cname, ckind, cru, cuk) in enumerate(columns, 1):
            s += f'{t}\t\t<Column name="{cname}" id="{cid}">\n' + ml("Title", cru, cuk, t + "\t\t\t")
            s += f"{t}\t\t\t<Type>\n" + type_xml(ckind, t + "\t\t\t\t") + f"{t}\t\t\t</Type>\n{t}\t\t</Column>\n"
        s += f"{t}\t</Columns>\n"
    return s + f"{t}</Attribute>\n"


def dyn_attr(aid, query, order_field, t):
    q_xml = query.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return (f'{t}<Attribute name="Список" id="{aid}">\n{t}\t<Type>\n{t}\t\t<v8:Type>cfg:DynamicList</v8:Type>\n{t}\t</Type>\n'
            f"{t}\t<MainAttribute>true</MainAttribute>\n{t}\t<Settings xsi:type=\"DynamicList\">\n"
            f"{t}\t\t<ManualQuery>true</ManualQuery>\n{t}\t\t<DynamicDataRead>true</DynamicDataRead>\n"
            f"{t}\t\t<QueryText>{q_xml}</QueryText>\n{t}\t\t<MainTable>Catalog.{CAT}</MainTable>\n{t}\t\t<ListSettings>\n"
            f"{t}\t\t\t<dcsset:filter>\n{t}\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n"
            f"{t}\t\t\t\t<dcsset:userSettingID>dfcece9d-5077-440b-b6b3-45a5cb4538eb</dcsset:userSettingID>\n{t}\t\t\t</dcsset:filter>\n"
            f"{t}\t\t\t<dcsset:order>\n{t}\t\t\t\t<dcsset:item xsi:type=\"dcsset:OrderItemField\">\n"
            f"{t}\t\t\t\t\t<dcsset:field>{order_field}</dcsset:field>\n{t}\t\t\t\t\t<dcsset:orderType>Asc</dcsset:orderType>\n"
            f"{t}\t\t\t\t</dcsset:item>\n{t}\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n"
            f"{t}\t\t\t\t<dcsset:userSettingID>88619765-ccb3-46c6-ac52-38e9c992ebd4</dcsset:userSettingID>\n{t}\t\t\t</dcsset:order>\n"
            f"{t}\t\t\t<dcsset:conditionalAppearance>\n{t}\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n"
            f"{t}\t\t\t\t<dcsset:userSettingID>b75fecce-942b-4aed-abc9-e6a02e460fb3</dcsset:userSettingID>\n"
            f"{t}\t\t\t</dcsset:conditionalAppearance>\n{t}\t\t\t<dcsset:itemsViewMode>Normal</dcsset:itemsViewMode>\n"
            f"{t}\t\t\t<dcsset:itemsUserSettingID>911b6018-f537-43e8-a417-da56b22f9aec</dcsset:itemsUserSettingID>\n"
            f"{t}\t\t</ListSettings>\n{t}\t</Settings>\n{t}</Attribute>\n")


def command(cid, name, ru, uk, t, modifies=False):
    s = f'{t}<Command name="{name}" id="{cid}">\n' + ml("Title", ru, uk, t + "\t") + ml("ToolTip", ru, uk, t + "\t")
    s += f"{t}\t<Action>{name}</Action>\n"
    if modifies:
        s += f"{t}\t<ModifiesSavedData>true</ModifiesSavedData>\n"
    return s + f"{t}</Command>\n"


def parameter(name, kind, t):
    return f'{t}<Parameter name="{name}">\n{t}\t<Type>\n' + type_xml(kind, t + "\t\t") + f"{t}\t</Type>\n{t}</Parameter>\n"


def form_xml(body, attrs, cmds="", params="", events=(), use_for=None, bar=""):
    s = FORM_HEAD
    if use_for:
        s += f"\t<UseForFoldersAndItems>{use_for}</UseForFoldersAndItems>\n"
    if bar:
        s += '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n\t\t<ChildItems>\n' + bar + "\t\t</ChildItems>\n\t</AutoCommandBar>\n"
    else:
        s += '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1"/>\n'
    if events:
        s += "\t<Events>\n" + "".join(f'\t\t<Event name="{e}">{h}</Event>\n' for e, h in events) + "\t</Events>\n"
    s += "\t<ChildItems>\n" + body + "\t</ChildItems>\n\t<Attributes>\n" + attrs + "\t</Attributes>\n"
    if cmds:
        s += "\t<Commands>\n" + cmds + "\t</Commands>\n"
    if params:
        s += "\t<Parameters>\n" + params + "\t</Parameters>\n"
    return s + "</Form>\n"


def form_md(form, ru, uk):
    return (MD_HEAD + f'\t<Form uuid="{uid(CAT + "." + form)}">\n\t\t<Properties>\n\t\t\t<Name>{form}</Name>\n'
            + ml("Synonym", ru, uk, "\t\t\t") +
            "\t\t\t<Comment/>\n\t\t\t<FormType>Managed</FormType>\n\t\t\t<IncludeHelpInContents>false</IncludeHelpInContents>\n"
            "\t\t\t<UsePurposes>\n\t\t\t\t<v8:Value xsi:type=\"app:ApplicationUsePurpose\">PlatformApplication</v8:Value>\n"
            "\t\t\t\t<v8:Value xsi:type=\"app:ApplicationUsePurpose\">MobilePlatformApplication</v8:Value>\n"
            "\t\t\t</UsePurposes>\n\t\t</Properties>\n\t</Form>\n</MetaDataObject>\n")


def register_form(form, default_tag=None):
    path = os.path.join(BB_WT, f"Catalogs/{CAT}.xml")
    s, _, _ = read_text(path)
    s2 = s
    if default_tag:
        ref = f"Catalog.{CAT}.Form.{form}"
        s2 = re.sub(rf"<{default_tag}/>|<{default_tag}>[^<]*</{default_tag}>", f"<{default_tag}>{ref}</{default_tag}>", s2, count=1)
        assert f"<{default_tag}>{ref}</{default_tag}>" in s2, default_tag
    if f"<Form>{form}</Form>" not in s2:
        head, sep, tail = s2.rpartition("\t\t</ChildObjects>\n\t</Catalog>")
        assert sep, "конец ChildObjects справочника"
        s2 = head + f"\t\t\t<Form>{form}</Form>\n" + sep + tail
    if s2 != s:
        write_text(path, s2)


def write_form(form, ru, uk, xml, module_parts, default_tag=None):
    base = f"Catalogs/{CAT}/Forms/{form}"
    write_text(os.path.join(BB_WT, base + ".xml"), form_md(form, ru, uk))
    write_text(os.path.join(BB_WT, base + "/Ext/Form.xml"), xml)
    if module_parts:
        write_text(os.path.join(BB_WT, base + "/Ext/Form/Module.bsl"), "\n".join(read_src(p).rstrip("\n") + "\n" for p in module_parts))
    register_form(form, default_tag)
    print("  форма", form)


def F(name, path, **kw):
    return lambda i, t: field(i, name, path, t, **kw)


def C(name, path, **kw):
    """Колонка таблицы формы."""
    return lambda i, t: field(i, name, path, t, in_table=True, **kw)


def G(name, children, **kw):
    return lambda i, t: group(i, name, children, t, **kw)


def T(name, path, columns, **kw):
    return lambda i, t: table(i, name, path, columns, t, **kw)


def P(name, ru, uk, children):
    return lambda i, t: page(i, name, ru, uk, children, t)


REF_UNIT = "cfg:CatalogRef.КлассификаторЕдиницИзмерения"
TREE_COLS = [("Представление", "str:250", "Строка", "Рядок"), ("ПозицияСС", "str:200", "Позиция СС", "Позиція СС"),
             ("Количество", "num:15,6", "Количество на дом", "Кількість на будинок"),
             ("ЕдиницаСтроки", REF_UNIT, "Единица строки", "Одиниця рядка"),
             ("КоличествоКарточки", "num:15,6", "В ед. карточки", "В од. картки"),
             ("СуммаНормы", "num:15,2", "Сумма нормы, с НДС", "Сума норми, з ПДВ")]
RAZV_COLS = [("Номенклатура", "cfg:CatalogRef.Номенклатура", "Карточка", "Картка"),
             ("ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "Общее название", "Загальна назва"),
             ("ПозицияСС", "str:200", "Позиция СС", "Позиція СС"),
             ("КоличествоЕдСтроки", "num:15,6", "Количество в ед. строки", "Кількість в од. рядка"),
             ("ЕдиницаСтроки", REF_UNIT, "Единица строки", "Одиниця рядка"),
             ("Количество", "num:15,6", "Норма в ед. карточки", "Норма в од. картки"),
             ("ЦенаНормы", "num:16,6", "Цена нормы, с НДС", "Ціна норми, з ПДВ"),
             ("СуммаНормы", "num:15,2", "Сумма нормы, с НДС", "Сума норми, з ПДВ"),
             ("ПутьУзла", "str:200", "Путь узлов", "Шлях вузлів"), ("НеПоставляется", "bool", "Не поставляется", "Не постачається")]
SR_COLS = [("НомерСтроки", "num:5,0", "№ СС", "№ СС"), ("ПозицияСС", "str:200", "Позиция СС", "Позиція СС"),
           ("КоличествоСС", "num:15,6", "Количество СС", "Кількість СС"), ("ЕдиницаСС", "str:20", "Ед. СС", "Од. СС"),
           ("КоличествоСпец", "num:15,6", "Количество спецификации", "Кількість специфікації"),
           ("СуммаСС", "num:15,2", "Сумма СС", "Сума СС"), ("СуммаНормы", "num:15,2", "Сумма нормы", "Сума норми"),
           ("Статус", "str:100", "Статус", "Статус")]
CMD1 = [("СделатьЭталономСС", "Сделать эталоном для СС", "Зробити еталоном для СС", False),
        ("ПоказатьВходимость", "Входимость", "Входимість", False),
        ("ОбновитьДерево", "Обновить дерево", "Оновити дерево", False),
        ("РазвернутьДоДеталей", "Развернуть до деталей", "Розгорнути до деталей", False)]
CMD2_SOSTAV = [("ЗаполнитьИзСС", "Заполнить из СС", "Заповнити з СС", True),
               ("ОтметитьЗеленыеПроверенными", "Отметить зелёные проверенными", "Позначити зелені перевіреними", True),
               ("ПодобратьКарточку", "Подобрать карточку", "Підібрати картку", True),
               ("ВзятьИзОригиналаExcel", "Взять из оригинала Excel", "Взяти з оригіналу Excel", True),
               ("ОбъединитьВУзел", "Объединить в узел", "Обʼєднати у вузол", True),
               ("РасформироватьУзел", "Расформировать узел", "Розформувати вузол", True),
               ("СоставИзДокументаКомплектации", "Состав из документа комплектации", "Склад з документа комплектації", True)]
CMD2_FORM = [("РассчитатьЦеныНормы", "Рассчитать цены нормы", "Розрахувати ціни норми", True),
             ("СравнитьСоСС", "Сравнить с СС", "Порівняти з СС", False),
             ("НулевыеПозицииСФактом", "Нулевые позиции с фактом", "Нульові позиції з фактом", False),
             ("СкопироватьИзДругойСпецификации", "Скопировать карточки и узлы из другой спецификации",
              "Скопіювати картки й дерево вузлів з іншої специфікації", True)]


def forma_elementa(etap):
    ids = Ids()
    sostav_cols = [
        C("СоставНомерСтроки", "Объект.Состав.LineNumber", kind="LabelField", width=4),
        C("СоставПроверено", "Объект.Состав.Проверено", kind="CheckBoxField", checkbox_type=True),
        C("СоставВидСтроки", "Объект.Состав.ВидСтроки"),
        C("СоставЭтапСС", "Объект.Состав.ЭтапСС"),
        C("СоставПозицияСС", "Объект.Состав.ПозицияСС", width=25),
        C("СоставНоменклатура", "Объект.Состав.Номенклатура", width=25, events=[("OnChange", "СоставНоменклатураПриИзменении")]),
        C("СоставИсточникКарточки", "Объект.Состав.ИсточникКарточки"),
        C("СоставУзел", "Объект.Состав.Узел", events=[("OnChange", "СоставУзелПриИзменении")]),
        C("СоставОбщееНазвание", "Объект.Состав.ОбщееНазвание"),
        C("СоставКоличество", "Объект.Состав.Количество", events=[("OnChange", "СоставКоличествоПриИзменении")]),
        C("СоставЕдиницаСтроки", "Объект.Состав.ЕдиницаСтроки", events=[("OnChange", "СоставЕдиницаСтрокиПриИзменении")]),
        C("СоставКоэффициент", "Объект.Состав.Коэффициент", events=[("OnChange", "СоставКоэффициентПриИзменении")]),
        C("СоставИсточникКоэффициента", "Объект.Состав.ИсточникКоэффициента"),
        C("СоставЦенаНормы", "Объект.Состав.ЦенаНормы"),
        C("СоставНеПоставляется", "Объект.Состав.НеПоставляется", kind="CheckBoxField", checkbox_type=True),
        C("СоставПокупаетсяГотовым", "Объект.Состав.ПокупаетсяГотовым", kind="CheckBoxField", checkbox_type=True),
        C("СоставКомментарий", "Объект.Состав.Комментарий", width=20),
        C("СоставКлючСтроки", "Объект.Состав.КлючСтроки", visible=False),
    ]
    zak_cols = [C(f"ЗакупочныйСостав{n}", f"Объект.ЗакупочныйСостав.{n}") for n in (
        "Номенклатура", "ГруппаЗакупки", "ОбщееНазвание", "Узел", "ЭтапСС", "ПозицияСС", "НормаЭталона", "СреднееФакта",
        "ДомовСФактом", "КоэффициентВариации", "Количество", "ЦенаБезНДС", "ЦенаСНДС", "ИсточникЦены", "ДокументЦены",
        "ЦенаНормы", "Комментарий")]
    zak_cols += [C(f"ЗакупочныйСостав{n}", f"Объект.ЗакупочныйСостав.{n}", kind="CheckBoxField", checkbox_type=True)
                 for n in ("ИзмененоВручную", "ПотребуетРешения", "Проверено")]
    fakt_cols = [C(f"ФактПоДомам{n}", f"Объект.ФактПоДомам.{n}") for n in (
        "Подразделение", "Номенклатура", "Списано", "СписаноСумма", "Лежит", "ЛежитСумма")]
    fakt_cols.append(C("ФактПоДомамИзСборки", "Объект.ФактПоДомам.ИзСборки", kind="CheckBoxField", checkbox_type=True))
    head = G("ГруппаШапка", [
        G("ГруппаШапкаЛево", [
            F("Наименование", "Объект.Description", stretch=True),
            G("ГруппаВидСтатус", [
                F("ВидСпецификации", "Объект.ВидСпецификации", events=[("OnChange", "ВидСпецификацииПриИзменении")]),
                F("Статус", "Объект.Статус"),
                F("ДатаНачала", "Объект.ДатаНачала")]),
            F("СтруктураСебестоимости", "Объект.СтруктураСебестоимости", stretch=True),
            F("Основание", "Объект.Основание", ro=True, stretch=True)], orient="Vertical"),
        G("ГруппаШапкаПраво", [
            G("ГруппаКодРодитель", [F("Код", "Объект.Code"), F("Родитель", "Объект.Parent", title=("Папка", "Папка"))]),
            G("ГруппаВыпуск", [F("КоличествоВыпуска", "Объект.КоличествоВыпуска"), F("ЕдиницаВыпуска", "Объект.ЕдиницаВыпуска")]),
            F("НоменклатураВузла", "Объект.НоменклатураВузла", stretch=True),
            G("ГруппаДаты", [F("ДатаИзменения", "Объект.ДатаИзменения", ro=True), F("ДатаЦенНормы", "Объект.ДатаЦенНормы", ro=True),
                             F("ДатаФакта", "Объект.ДатаФакта", ro=True)])], orient="Vertical")])
    pg = [
        P("СтраницаСостав", "Состав", "Склад", [
            G("ГруппаОтбор", [
                F("ТолькоПроблемные", "ТолькоПроблемные", kind="CheckBoxField", title_loc="Right", checkbox_type=True,
                  events=[("OnChange", "ТолькоПроблемныеПриИзменении")]),
                F("ИтогСравнения", "ИтогСравнения", kind="LabelField", stretch=True)]),
            T("Состав", "Объект.Состав", sostav_cols, buttons=[c[0] for c in CMD2_SOSTAV] if etap == 2 else [])]),
        P("СтраницаДома", "Дома", "Будинки", [T("Дома", "Объект.Дома", [
            C("ДомаНомерСтроки", "Объект.Дома.LineNumber", kind="LabelField", width=4),
            C("ДомаПодразделение", "Объект.Дома.Подразделение"), C("ДомаКомментарий", "Объект.Дома.Комментарий")])]),
        P("СтраницаДерево", "Дерево", "Дерево", [T("ДеревоУзлов", "ДеревоУзлов",
            [C(f"ДеревоУзлов{n}", f"ДеревоУзлов.{n}") for n, *_ in TREE_COLS], tree=True, ro=True, change_rows=False,
            buttons=["ОбновитьДерево"])]),
        P("СтраницаРазвернутыйСостав", "Развёрнутый состав", "Розгорнутий склад", [T("РазвернутыйСостав", "РазвернутыйСостав",
            [C(f"РазвернутыйСостав{n}", f"РазвернутыйСостав.{n}") for n, *_ in RAZV_COLS], ro=True, change_rows=False,
            buttons=["РазвернутьДоДеталей"])]),
        P("СтраницаСравнение", "Сравнение с СС", "Порівняння з СС", [T("СравнениеСС", "СравнениеСС",
            [C(f"СравнениеСС{n}", f"СравнениеСС.{n}") for n, *_ in SR_COLS], ro=True, change_rows=False,
            buttons=["СравнитьСоСС"] if etap == 2 else [])]),
        P("СтраницаЗакупка", "Закупка", "Закупівля", [
            T("ДомаФакта", "Объект.ДомаФакта", [C("ДомаФактаПодразделение", "Объект.ДомаФакта.Подразделение"),
                                                C("ДомаФактаФактСумма", "Объект.ДомаФакта.ФактСумма"),
                                                C("ДомаФактаКомментарий", "Объект.ДомаФакта.Комментарий")], height=4),
            T("ЗакупочныйСостав", "Объект.ЗакупочныйСостав", zak_cols)]),
        P("СтраницаФактПоДомам", "Факт по домам", "Факт по будинках", [T("ФактПоДомам", "Объект.ФактПоДомам", fakt_cols)]),
    ]
    body = head(ids, "\t\t") + pages(ids, "Страницы", pg, "\t\t")
    body += field(ids, "Комментарий", "Объект.Комментарий", "\t\t", height=2, stretch=True, multiline=True)
    cmds = CMD1 + (CMD2_SOSTAV + CMD2_FORM if etap == 2 else [])
    bar_cmds = ["СделатьЭталономСС", "ПоказатьВходимость"] + ([c[0] for c in CMD2_FORM if c[0] != "СравнитьСоСС"] if etap == 2 else [])
    bar = "".join(button(ids, f"Форма{c}", c, "\t\t\t") for c in bar_cmds)
    attrs = ('\t\t<Attribute name="Объект" id="1">\n\t\t\t<Type>\n\t\t\t\t<v8:Type>cfg:CatalogObject.А_СпецификацииДомов</v8:Type>\n'
             "\t\t\t</Type>\n\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<SavedData>true</SavedData>\n\t\t</Attribute>\n")
    attrs += attr(2, "ДеревоУзлов", "vtree", "\t\t", title=("Дерево узлов", "Дерево вузлів"), columns=TREE_COLS)
    attrs += attr(3, "РазвернутыйСостав", "vt", "\t\t", title=("Развёрнутый состав", "Розгорнутий склад"), columns=RAZV_COLS)
    attrs += attr(4, "СравнениеСС", "vt", "\t\t", title=("Сравнение с СС", "Порівняння з СС"), columns=SR_COLS)
    attrs += attr(5, "ТолькоПроблемные", "bool", "\t\t", title=("Только проблемные (не проверенные)", "Лише проблемні (не перевірені)"))
    attrs += attr(6, "ИтогСравнения", "str:300", "\t\t", title=("Σ нормы / Σ СС", "Σ норми / Σ СС"))
    cmd_xml = "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1))
    xml = form_xml(body, attrs, cmd_xml, events=[("OnCreateAtServer", "ПриСозданииНаСервере")], use_for="Items", bar=bar)
    parts = ["forma_elementa_1.bsl"] + (["forma_elementa_2.bsl"] if etap == 2 else [])
    write_form("ФормаЭлемента", "Форма элемента", "Форма елемента", xml, parts, "DefaultObjectForm")


def forma_spiska(etap):
    ids = Ids()
    cols = [F("Наименование", "Список.Наименование", kind="LabelField"), F("Код", "Список.Код", kind="LabelField"),
            F("ВидСпецификации", "Список.ВидСпецификации", kind="LabelField"),
            F("Статус", "Список.Статус", kind="LabelField"),
            F("СтруктураСебестоимости", "Список.СтруктураСебестоимости", kind="LabelField"),
            F("ЭталонДляСС", "Список.ЭталонДляСС", kind="LabelField", title=("Эталон для СС", "Еталон для СС")),
            F("ДатаИзменения", "Список.ДатаИзменения", kind="LabelField")]
    ug = "СписокКомпоновщикНастроекПользовательскиеНастройки"
    body = user_settings_group(ids, ug, "\t\t") + dyn_table(ids, "Список", cols, "\t\t", ug)
    cmds = [("ПроверитьЦиклы", "Проверить циклы", "Перевірити цикли", False)]
    if etap == 2:
        cmds.append(("ПерепривязатьПозицииСС", "Перепривязать позиции СС…", "Перепривʼязати позиції СС…", False))
    bar = "".join(button(ids, f"Форма{c[0]}", c[0], "\t\t\t") for c in cmds)
    xml = form_xml(body, dyn_attr(1, zapros("spisok_formy"), "Наименование", "\t\t"),
                   "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1)), bar=bar)
    parts = ["forma_spiska_1.bsl"] + (["forma_spiska_2.bsl"] if etap == 2 else [])
    write_form("ФормаСписка", "Форма списка", "Форма списку", xml, parts, "DefaultListForm")


def forma_gruppy():
    ids = Ids()
    body = group(ids, "ГруппаНаименованиеКод", [F("Наименование", "Объект.Description", stretch=True),
                                                F("Код", "Объект.Code")], "\t\t")
    body += field(ids, "Родитель", "Объект.Parent", "\t\t", title=("Входит в папку", "Входить у папку"))
    attrs = ('\t\t<Attribute name="Объект" id="1">\n\t\t\t<Type>\n\t\t\t\t<v8:Type>cfg:CatalogObject.А_СпецификацииДомов</v8:Type>\n'
             "\t\t\t</Type>\n\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<SavedData>true</SavedData>\n\t\t</Attribute>\n")
    write_form("ФормаГруппы", "Форма группы", "Форма групи", form_xml(body, attrs, use_for="Folders"), [], "DefaultFolderForm")


PODBOR_COLS = [("Номенклатура", "cfg:CatalogRef.Номенклатура", "Карточка", "Картка"), ("Код", "str:20", "Код", "Код"),
               ("Единица", REF_UNIT, "Единица", "Одиниця"),
               ("ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "Общее название", "Загальна назва"),
               ("СписаноКоличество", "num:15,3", "Списано на дома типа", "Списано на будинки типу"),
               ("СписаноСумма", "num:15,2", "Списано, без НДС", "Списано, без ПДВ"),
               ("ДомовСФактом", "num:3,0", "Домов", "Будинків"), ("ФактПоДомам", "str:200", "По домам", "По будинках"),
               ("СкладПроекта", "num:15,3", "Склад проекта", "Склад проєкту"),
               ("ЦенаПоследнейЗакупки", "num:15,2", "Цена посл. закупки, с НДС", "Ціна ост. закупівлі, з ПДВ"),
               ("РазмерСовпал", "bool", "Размер из названия совпал", "Розмір з назви збігся")]
PEREPRIV_COLS = [("Применить", "bool", "Применить", "Застосувати"),
                 ("Спецификация", "cfg:CatalogRef.А_СпецификацииДомов", "Спецификация", "Специфікація"),
                 ("КлючСтроки", "uuid", "Ключ строки", "Ключ рядка"),
                 ("СтарыйЭтап", "cfg:CatalogRef.ЭтапыРабот", "Этап (был)", "Етап (був)"),
                 ("СтараяПозиция", "str:200", "Позиция (была)", "Позиція (була)"),
                 ("НовыйЭтап", "cfg:CatalogRef.ЭтапыРабот", "Этап (новый)", "Етап (новий)"),
                 ("НоваяПозиция", "str:200", "Позиция (новая)", "Позиція (нова)"),
                 ("НомерСтрокиСС", "num:5,0", "№ СС", "№ СС"), ("Цвет", "str:10", "Предложение", "Пропозиція"),
                 ("Причина", "str:250", "Причина", "Причина"), ("ДругаяСС", "bool", "Узел другой СС", "Вузол іншої СС")]


def forma_podbora():
    ids = Ids()
    body = group(ids, "ГруппаКонтекст", [F("ОбщееНазвание", "ОбщееНазвание", ro=True), F("ПозицияСС", "ПозицияСС", ro=True, stretch=True)], "\t\t")
    body += group(ids, "ГруппаОтбор", [
        F("ТолькоОН", "ТолькоОН", kind="CheckBoxField", title_loc="Right", checkbox_type=True, events=[("OnChange", "ТолькоОНПриИзменении")]),
        F("СтрокаПоиска", "СтрокаПоиска", stretch=True, events=[("OnChange", "СтрокаПоискаПриИзменении")])], "\t\t")
    body += table(ids, "Кандидаты", "Кандидаты", [C(f"Кандидаты{n}", f"Кандидаты.{n}", **({"kind": "CheckBoxField", "checkbox_type": True} if k == "bool" else {}))
                                                  for n, k, *_ in PODBOR_COLS], "\t\t", ro=True, change_rows=False,
                  events=[("Selection", "КандидатыВыбор")])
    attrs = (attr(1, "Кандидаты", "vt", "\t\t", title=("Кандидаты", "Кандидати"), columns=PODBOR_COLS)
             + attr(2, "ТолькоОН", "bool", "\t\t", title=("Только общее название позиции", "Лише загальна назва позиції"))
             + attr(3, "СтрокаПоиска", "str:100", "\t\t", title=("Поиск в названии", "Пошук у назві"))
             + attr(4, "ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t", title=("Тип дома (СС)", "Тип будинку (СС)"))
             + attr(5, "ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "\t\t", title=("Общее название", "Загальна назва"))
             + attr(6, "ПозицияСС", "str:200", "\t\t", title=("Позиция СС", "Позиція СС")))
    cmds = [("ВыбратьКарточку", "Выбрать", "Вибрати", False), ("ОбновитьКандидатов", "Обновить", "Оновити", False)]
    bar = "".join(button(ids, f"Форма{c[0]}", c[0], "\t\t\t") for c in cmds)
    params = (parameter("ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t")
              + parameter("ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "\t\t") + parameter("ПозицияСС", "str:200", "\t\t"))
    xml = form_xml(body, attrs, "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1)),
                   params, events=[("OnCreateAtServer", "ПриСозданииНаСервере")], bar=bar)
    write_form("ФормаПодбораКарточки", "Подбор карточки", "Підбір картки", xml, ["forma_podbora.bsl"])


def forma_perepriv():
    ids = Ids()
    body = field(ids, "ТипБудинкуСС", "ТипБудинкуСС", "\t\t", events=[("OnChange", "ТипБудинкуССПриИзменении")])
    cols = []
    for n, k, *_ in PEREPRIV_COLS:
        kw = {"kind": "CheckBoxField", "checkbox_type": True} if k == "bool" else {}
        if n != "Применить":
            kw["ro"] = True
        if n == "КлючСтроки":
            kw["visible"] = False
        cols.append(C(f"Предложения{n}", f"Предложения.{n}", **kw))
    body += table(ids, "Предложения", "Предложения", cols, "\t\t", change_rows=False)
    attrs = (attr(1, "ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t", title=("Тип дома (СС)", "Тип будинку (СС)"))
             + attr(2, "Предложения", "vt", "\t\t", title=("Предложения", "Пропозиції"), columns=PEREPRIV_COLS))
    cmds = [("НайтиПредложения", "Найти сироты и предложения", "Знайти сироти й пропозиції", False),
            ("ПрименитьОтмеченные", "Применить отмеченные", "Застосувати позначені", False)]
    bar = "".join(button(ids, f"Форма{c[0]}", c[0], "\t\t\t") for c in cmds)
    xml = form_xml(body, attrs, "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1)),
                   parameter("ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t"),
                   events=[("OnCreateAtServer", "ПриСозданииНаСервере")], bar=bar)
    write_form("ФормаПерепривязки", "Перепривязка позиций СС", "Перепривʼязка позицій СС", xml, ["forma_perepriv.bsl"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--etap", type=int, choices=[1, 2], required=True)
    e = ap.parse_args().etap
    forma_elementa(e)
    forma_spiska(e)
    if e == 1:
        forma_gruppy()             # на этапе 2 не меняется: перегенерация затёрла бы канон платформы (Task 28) в worktree
    else:
        forma_podbora()
        forma_perepriv()
    print("OK")
```

`_Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_1.bsl`:

```bsl
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)

	Если Объект.Ссылка.Пустая() И Не Объект.ЭтоГруппа Тогда
		Если Объект.КоличествоВыпуска = 0 Тогда
			Объект.КоличествоВыпуска = 1;
		КонецЕсли;
		Если Не ЗначениеЗаполнено(Объект.Статус) Тогда
			Объект.Статус = Перечисления.А_СтатусыСпецификацийДомов.Черновик;
		КонецЕсли;
	КонецЕсли;
	УстановитьУсловноеОформление();
	НастроитьВидимость(ЭтотОбъект);

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ВидСпецификацииПриИзменении(Элемент)

	НастроитьВидимость(ЭтотОбъект);

КонецПроцедуры

&НаКлиенте
Процедура ТолькоПроблемныеПриИзменении(Элемент)

	Если ТолькоПроблемные Тогда
		Элементы.Состав.ОтборСтрок = Новый ФиксированнаяСтруктура("Проверено", Ложь);
	Иначе
		Элементы.Состав.ОтборСтрок = Неопределено;
	КонецЕсли;

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыСостав

&НаКлиенте
Процедура СоставНоменклатураПриИзменении(Элемент)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	ТекущиеДанные.Проверено = Ложь;
	ТекущиеДанные.ИсточникКарточки = ПредопределенноеЗначение("Перечисление.А_ИсточникиКарточкиСпецификации.Вручную");
	ОбновитьСтрокуПослеСменыКарточки(ТекущиеДанные.ПолучитьИдентификатор());

КонецПроцедуры

&НаКлиенте
Процедура СоставУзелПриИзменении(Элемент)

	СброситьПроверено();

КонецПроцедуры

&НаКлиенте
Процедура СоставКоличествоПриИзменении(Элемент)

	СброситьПроверено();

КонецПроцедуры

&НаКлиенте
Процедура СоставЕдиницаСтрокиПриИзменении(Элемент)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	ТекущиеДанные.Проверено = Ложь;
	ОбновитьСтрокуПослеСменыКарточки(ТекущиеДанные.ПолучитьИдентификатор());

КонецПроцедуры

&НаКлиенте
Процедура СоставКоэффициентПриИзменении(Элемент)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	ТекущиеДанные.Проверено = Ложь;
	ТекущиеДанные.ИсточникКоэффициента = ПредопределенноеЗначение("Перечисление.А_ИсточникиКоэффициентаПересчетаСС.Вручную");

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура СделатьЭталономСС(Команда)

	Если Модифицированность Или Объект.Ссылка.Пустая() Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Сначала запишите спецификацию.'; uk = 'Спочатку запишіть специфікацію.'"));
		Возврат;
	КонецЕсли;
	ПоказатьПредупреждение(, СделатьЭталономСССервер());

КонецПроцедуры

&НаКлиенте
Процедура ПоказатьВходимость(Команда)

	ПоказатьПредупреждение(, ВходимостьТекстом());

КонецПроцедуры

&НаКлиенте
Процедура ОбновитьДерево(Команда)

	ОбновитьДеревоНаСервере();

КонецПроцедуры

&НаКлиенте
Процедура РазвернутьДоДеталей(Команда)

	РазвернутьДоДеталейНаСервере();

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура СброситьПроверено()

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено Тогда
		ТекущиеДанные.Проверено = Ложь;
	КонецЕсли;

КонецПроцедуры

&НаСервере
Процедура ОбновитьСтрокуПослеСменыКарточки(ИдентификаторСтроки)

	СтрокаСостава = Объект.Состав.НайтиПоИдентификатору(ИдентификаторСтроки);
	Если СтрокаСостава = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Если ЗначениеЗаполнено(СтрокаСостава.Номенклатура) Тогда
		СтрокаСостава.ОбщееНазвание = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СтрокаСостава.Номенклатура, "ОбщееНазваниеНоменклатуры");
	КонецЕсли;
	А_СпецификацииДомовСервер.УстановитьКоэффициентПоЕдиницам(СтрокаСостава);

КонецПроцедуры

&НаСервере
Функция СделатьЭталономСССервер()

	Попытка
		А_СпецификацииДомовСервер.СделатьЭталономСС(Объект.Ссылка);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	Возврат НСтр("ru = 'Спецификация установлена эталоном для СС.'; uk = 'Специфікацію встановлено еталоном для СС.'");

КонецФункции

&НаСервере
Функция ВходимостьТекстом()

	Если Объект.Ссылка.Пустая() Тогда
		Возврат НСтр("ru = 'Спецификация ещё не записана.'; uk = 'Специфікацію ще не записано.'");
	КонецЕсли;
	Входы = А_СпецификацииДомовСервер.Входимость(Объект.Ссылка);
	Если Входы.Количество() = 0 Тогда
		Возврат НСтр("ru = 'Нигде не используется.'; uk = 'Ніде не використовується.'");
	КонецЕсли;
	Части = Новый Массив;
	Для Каждого Вход Из Входы Цикл
		Части.Добавить(СтрШаблон(НСтр("ru = 'уровень %1: %2 (%3; %4)'; uk = 'рівень %1: %2 (%3; %4)'"),
			Вход.Уровень, Вход.Спецификация, Вход.ВидСпецификации, Вход.СтруктураСебестоимости));
	КонецЦикла;
	Возврат СтрСоединить(Части, Символы.ПС);

КонецФункции

&НаСервере
Процедура ОбновитьДеревоНаСервере()

	Попытка
		Разузлование = А_СпецификацииДомовСервер.РазвернутьДоКарточек(РеквизитФормыВЗначение("Объект"));
	Исключение
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(КраткоеПредставлениеОшибки(ИнформацияОбОшибке()));
		Возврат;
	КонецПопытки;
	Дерево = РеквизитФормыВЗначение("ДеревоУзлов");
	Дерево.Строки.Очистить();
	ДобавитьВетку(Дерево.Строки, "", Разузлование);
	ЗначениеВРеквизитФормы(Дерево, "ДеревоУзлов");

КонецПроцедуры

&НаСервере
Процедура ДобавитьВетку(СтрокиДерева, КлючЭкземпляра, Разузлование)

	Для Каждого Экз Из Разузлование.Экземпляры.НайтиСтроки(Новый Структура("КлючЭкземпляраРодителя", КлючЭкземпляра)) Цикл
		Ветка = СтрокиДерева.Добавить();
		Ветка.Представление = СтрШаблон(НСтр("ru = '[Узел] %1'; uk = '[Вузол] %1'"), Экз.Узел);
		Ветка.Количество = Экз.Количество;
		Ветка.СуммаНормы = Экз.СуммаНормы;
		ДобавитьВетку(Ветка.Строки, Экз.КлючЭкземпляра, Разузлование);
	КонецЦикла;
	Для Каждого Лист Из Разузлование.Листья.НайтиСтроки(Новый Структура("КлючЭкземпляра", КлючЭкземпляра)) Цикл
		Ветка = СтрокиДерева.Добавить();
		Ветка.Представление = ?(ЗначениеЗаполнено(Лист.Номенклатура), Строка(Лист.Номенклатура), Строка(Лист.ОбщееНазвание));
		Ветка.ПозицияСС = Лист.ПозицияСС;
		Ветка.Количество = Лист.КоличествоЕдСтроки;
		Ветка.ЕдиницаСтроки = Лист.ЕдиницаСтроки;
		Ветка.КоличествоКарточки = Лист.Количество;
		Ветка.СуммаНормы = Лист.СуммаНормы;
	КонецЦикла;

КонецПроцедуры

&НаСервере
Процедура РазвернутьДоДеталейНаСервере()

	Попытка
		Разузлование = А_СпецификацииДомовСервер.РазвернутьДоКарточек(РеквизитФормыВЗначение("Объект"), Истина);
	Исключение
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(КраткоеПредставлениеОшибки(ИнформацияОбОшибке()));
		Возврат;
	КонецПопытки;
	Листья = Разузлование.Листья;
	Листья.Сортировать("ПозицияСС, ПутьУзла");
	РазвернутыйСостав.Очистить();
	Для Каждого Лист Из Листья Цикл
		ЗаполнитьЗначенияСвойств(РазвернутыйСостав.Добавить(), Лист);
	КонецЦикла;

КонецПроцедуры

&НаКлиентеНаСервереБезКонтекста
Процедура НастроитьВидимость(Форма)

	Вид = Форма.Объект.ВидСпецификации;
	ЭтоЗакупочная = (Вид = ПредопределенноеЗначение("Перечисление.А_ВидыСпецификацийДомов.Закупочная"));
	ЭтоУзел = (Вид = ПредопределенноеЗначение("Перечисление.А_ВидыСпецификацийДомов.Узел"));
	ЭлементыФормы = Форма.Элементы;
	ЭлементыФормы.СтраницаСостав.Видимость = Не ЭтоЗакупочная;
	ЭлементыФормы.СтраницаДома.Видимость = Не ЭтоЗакупочная И Не ЭтоУзел;
	ЭлементыФормы.СтраницаДерево.Видимость = Не ЭтоЗакупочная;
	ЭлементыФормы.СтраницаРазвернутыйСостав.Видимость = Не ЭтоЗакупочная;
	ЭлементыФормы.СтраницаСравнение.Видимость = Не ЭтоЗакупочная;
	ЭлементыФормы.СтраницаЗакупка.Видимость = ЭтоЗакупочная;
	ЭлементыФормы.СтраницаФактПоДомам.Видимость = ЭтоЗакупочная;
	ЭлементыФормы.ГруппаВыпуск.Видимость = Не ЭтоЗакупочная;
	ЭлементыФормы.НоменклатураВузла.Видимость = ЭтоУзел;
	ЭлементыФормы.ДатаФакта.Видимость = ЭтоЗакупочная;

КонецПроцедуры

&НаСервере
Процедура УстановитьУсловноеОформление()

	Желтые = Новый СписокЗначений;
	Желтые.Добавить(Перечисления.А_ИсточникиКарточкиСпецификации.СписаннаяНаДомаТипа);
	Желтые.Добавить(Перечисления.А_ИсточникиКарточкиСпецификации.КопияИзДругойСпецификации);
	ДобавитьОформление("Состав", "Объект.Состав.ИсточникКарточки", ВидСравненияКомпоновкиДанных.ВСписке, Желтые,
		"ЦветФона", Новый Цвет(255, 242, 204));
	ДобавитьОформление("Состав", "Объект.Состав.Проверено", ВидСравненияКомпоновкиДанных.Равно, Истина,
		"ЦветФона", Новый Цвет(226, 239, 218));
	ДобавитьОформление("СоставКоэффициент, СоставИсточникКоэффициента", "Объект.Состав.ИсточникКоэффициента",
		ВидСравненияКомпоновкиДанных.Равно, Перечисления.А_ИсточникиКоэффициентаПересчетаСС.НеНайден, "ЦветТекста", Новый Цвет(192, 0, 0));
	ДобавитьОформление("Состав", "Объект.Состав.НеПоставляется", ВидСравненияКомпоновкиДанных.Равно, Истина,
		"ЦветТекста", Новый Цвет(128, 128, 128));

КонецПроцедуры

&НаСервере
Процедура ДобавитьОформление(ИменаПолей, ПутьКПолю, ВидСравнения, Значение, ИмяПараметра, ЗначениеПараметра)

	ЭлементОформления = УсловноеОформление.Элементы.Добавить();
	Для Каждого ИмяПоля Из СтрРазделить(ИменаПолей, ",") Цикл
		ОформляемоеПоле = ЭлементОформления.Поля.Элементы.Добавить();
		ОформляемоеПоле.Поле = Новый ПолеКомпоновкиДанных(СокрЛП(ИмяПоля));
	КонецЦикла;
	ЭлементОтбора = ЭлементОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение = Новый ПолеКомпоновкиДанных(ПутьКПолю);
	ЭлементОтбора.ВидСравнения = ВидСравнения;
	ЭлементОтбора.ПравоеЗначение = Значение;
	ЭлементОформления.Оформление.УстановитьЗначениеПараметра(ИмяПараметра, ЗначениеПараметра);

КонецПроцедуры

#КонецОбласти
```

`_Rarzrabotki/Python/test/spec_domov/bsl/forma_spiska_1.bsl`:

```bsl
#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ПроверитьЦиклы(Команда)

	Ошибки = ПроверитьЦиклыНаСервере();
	Если Ошибки.Количество() = 0 Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Циклов и превышения глубины нет.'; uk = 'Циклів і перевищення глибини немає.'"));
	Иначе
		ПоказатьПредупреждение(, СтрСоединить(Ошибки, Символы.ПС));
	КонецЕсли;

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервереБезКонтекста
Функция ПроверитьЦиклыНаСервере()

	Возврат А_СпецификацииДомовСервер.ПроверитьЦиклы();

КонецФункции

#КонецОбласти
```

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/gen_forms.py --etap 1`
Expected: `форма ФормаЭлемента`, `форма ФормаСписка`, `форма ФормаГруппы`, `OK`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only forms`
Expected: `forms OK`.
Run (для каждой из трёх форм): `powershell.exe -NoProfile -File C:/Configuration_downloads/BASERP25/.claude/skills/form-validate/scripts/form-validate.ps1 -FormPath "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента"`
Expected: ошибок нет; допустимы предупреждение о версии 2.13 и ложные `~Список.X` (память `form_validate_tilde_spisok_false_positive`).
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py "Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml" "Catalogs/А_СпецификацииДомов/Forms/ФормаСписка.xml" "Catalogs/А_СпецификацииДомов/Forms/ФормаГруппы.xml"`
Expected: `ошибок: 0` (в т. ч. «директива компиляции» перед каждой процедурой модулей форм).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/gen_forms.py _Rarzrabotki/Python/test/spec_domov/test_e1_files.py \
  _Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_1.bsl _Rarzrabotki/Python/test/spec_domov/bsl/forma_spiska_1.bsl \
  "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml" "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms"
git commit -m "feat(BASEBuh): формы А_СпецификацииДомов этапа 1 — элемент (7 страниц), список с «Еталон для СС», папка

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 27: Поле эталона на форме СС и регистрация общего модуля

**Files:**
- Modify: `_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml` (после поля `КоличествоДомов` в `Группа4`)
- Modify: `_Rarzrabotki/BASEBuh/Configuration.xml` (`<CommonModule>А_СпецификацииДомовСервер</CommonModule>` в конце группы `CommonModule`)
- Test: `test_e1_files.py --only ssform`

- [ ] **Step 1: Написать падающий тест**

Добавить в `test_e1_files.py`:

```python
def check_ssform():
    f = load("Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml")
    m = re.search(r'<InputField name="ЭталоннаяСпецификация" id="(\d+)">(.*?)</InputField>', f, re.S)
    assert m and "<DataPath>Объект.ЭталоннаяСпецификация</DataPath>" in m.group(2) and "<ReadOnly>true</ReadOnly>" in m.group(2)
    g4 = f[f.index('<UsualGroup name="Группа4"'):]
    assert g4.index('name="КоличествоДомов"') < g4.index('name="ЭталоннаяСпецификация"') < g4.index("</UsualGroup>")
    ids = re.findall(r' id="(\d+)"', f.split("<Attributes>")[0])
    assert len(ids) == len(set(ids)), "повтор id в форме СС"
    cfg = load("Configuration.xml")
    assert cfg.count("\t\t\t<CommonModule>А_СпецификацииДомовСервер</CommonModule>\n") == 1
    print("ssform OK")
```

и `"ssform": check_ssform` в `parts`.

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only ssform`
Expected: FAIL — `AssertionError` (поля нет).

- [ ] **Step 3: Реализовать** (функции уже есть в `patch_existing.py --etap form` и `register.py --etap 1b`)

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/patch_existing.py --etap form
python -X utf8 _Rarzrabotki/Python/test/spec_domov/register.py --etap 1b
```
Expected: `форма СС: + ЭталоннаяСпецификация (id 386–388, только просмотр)`; `Configuration.xml: +1`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py`
Expected: все разделы `OK` (`enums`, `catalog`, `existing`, `register`, `module`, `forms`, `ssform`).
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/preflight_config.py _Rarzrabotki/BASEBuh/Configuration.xml`
Expected: `новых П1 в файле, но не в базе: ['А_СпецификацииДомовСервер']`, `ПРОВАЛОВ: 0`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/test_e1_files.py _Rarzrabotki/BASEBuh/Configuration.xml \
  "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml"
git commit -m "feat(BASEBuh): поле «Еталонна специфікація» на форме СС (только просмотр, Р3); регистрация общего модуля

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 28: Загрузка 1B и приёмка этапа 1

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/load_lists/e1b.txt` (9 строк из раздела «Списки Partial-загрузки»)
- Create (результат): `_Rarzrabotki/Python/test/spec_domov/out/e1_priemka_<stamp>.json`, `out/e1b_load_<stamp>.txt`

- [ ] **Step 1: Тест** — `test_e1_funkcii.py`, `test_e1_priemka.py` (Tasks 19–20).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py`
Expected: `ПРОВАЛОВ: 9` (модуля в базе нет).

- [ ] **Step 3: Статика, копия в основной каталог, загрузка**

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/build_bsl.py --partiya 1
python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e1b.txt --tests
python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e1b.txt
powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e1b.txt -Tag e1b
```
Expected: `ошибок: 0`; `скопирован …` без `ОТКАЗ`; `OK load e1b`. Если `CheckModules` сообщает ошибку `{ОбщийМодуль.А_СпецификацииДомовСервер.Модуль(строка,колонка)}` — исправить исходник в `bsl/`, повторить `build_bsl.py` → `to_main.py` → `load_stage.ps1` (список можно сузить до файла модуля).

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_funkcii.py`
Expected: `OK хеш и единицы`, `OK разузлование`, `OK площадной узел`, `OK test_e1_funkcii`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py`
Expected: 11 строк `OK  k0_tx … k10_exchange_untouched`, `ПРОВАЛОВ: 0`. Это приёмка этапа 1 §10.1: цикл «Септик → Зовнішні мережі → Септик» не записывается (k1) и ловится при «Діє» и «Перевірити цикли» (k2); 9-й уровень разворачивается, 10-й — ошибка при записи и при разузловании (k3–k5, Т6); Б1 на СС 003 — 214 из 214, регистрация не изменилась (k6); «Деталь» без карточки — только с «Не постачається» и причиной (k8, Т9); закупочная без проверенных строк не в «Діє» (k9).
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/dump_check.py --tag e1b --canon --objects "ОбщийМодуль.А_СпецификацииДомовСервер,Справочник.А_СпецификацииДомов,Справочник.А_СпецификацииДомов.Форма.ФормаЭлемента,Справочник.А_СпецификацииДомов.Форма.ФормаСписка,Справочник.А_СпецификацииДомов.Форма.ФормаГруппы,Справочник.СтруктураСебестоимости.Форма.ФормаЭлемента"`
Expected: `различие по содержимому 0`.
Ручная проверка в клиенте (1 минута): открыть «Склад → Будинки: списання за СС → Специфікації будинків і вузлів», создать элемент — форма открывается, страницы видны по виду; не записывать.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/load_lists/e1b.txt _Rarzrabotki/Python/test/spec_domov/out/e1_priemka_*.json \
  _Rarzrabotki/Python/test/spec_domov/out/e1b_load_*.txt _Rarzrabotki/Python/test/spec_domov/sync_manifest.json \
  "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер.xml" "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl" \
  "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml" "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl" \
  "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms" "_Rarzrabotki/BASEBuh/Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml"
git commit -m "feat(BASEBuh): этап 1 П1 загружен и принят — проверки §4.6, разузлование, эталон для СС без порчи карточек

Приёмка §10.1 этап 1: цикл, глубина Т6, Б1 (СС 003 — 214/214, регистрация обмена без изменений), Т9, закупочная.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Фаза 2 — этап 2: эталон (загрузки 2A, 2B, 2C, 2D)

### Task 29: Rule #-1 — запросы этапа 2 (Python COM до BSL)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/zaprosy/doma_tipa.sql`, `fakt_spisaniy_domov.sql`, `kandidaty_kartochek.sql`, `ceny_posledney_zakupki.sql`, `ostatki_sklada_proekta.sql`, `kesh_edinic.sql`, `dokumenty_komplektacii.sql`, `sostav_dokumenta_komplektacii.sql`, `nulevye_pozicii.sql`, `specifikacii_ss.sql`, `kartochki_dlya_podbora.sql`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2_zaprosy.py`

Тексты, кроме `specifikacii_ss` и `kartochki_dlya_podbora`, уже исполнялись при подготовке плана (E3 t3–t7, пробы plan_p1/P 04.10.2026); тест подтверждает их перед BSL на день реализации.

- [ ] **Step 1: Написать тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2_zaprosy.py`:

```python
# -*- coding: utf-8 -*-
"""Rule #-1 для запросов этапа 2: каждый zaprosy/*.sql исполняется на bas_industrialbud и сверяется с независимыми числами
(свежий снимок этапа 0 — тот же день; E3 §3–§4; [M5] §4.5). Только чтение, кроме двух тестовых элементов в транзакции."""
import datetime
import glob
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, OUT, arr, connect, q, same, tx, vidy_subkonto, zapros  # noqa: E402
from fixtures import card, kn, ss  # noqa: E402

# «На сейчас» без переноса даты через COM: каждая дата из Python уезжает в 1С на −3 ч (память com_date_shift_3_hours),
# а ТекущаяДатаСеанса(), прочитанная и отданная обратно, отрезала бы движения последних часов. Дальняя дата = все движения.
FAR = datetime.datetime(2100, 1, 1)


def fresh_e0():
    files = sorted(glob.glob(os.path.join(OUT, "e0_snimok_*.json")))
    stamp = datetime.datetime.strptime(os.path.basename(files[-1])[10:23], "%Y%m%d_%H%M") if files else None
    if stamp is None or datetime.datetime.now() - stamp > datetime.timedelta(hours=2):
        subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "e0_snimok.py")], check=True)
        files = sorted(glob.glob(os.path.join(OUT, "e0_snimok_*.json")))
    return json.load(open(files[-1], encoding="utf-8"))["control_sums"]


def houses(B, code):
    rows = q(B, zapros("doma_tipa"), ПодразделениеОбъекта=ss(B, code).ПодразделениеОбъекта)
    return [r["Ссылка"] for r in rows]


def main():
    B = connect()
    e0 = fresh_e0()
    h8, h9 = houses(B, "000000008"), houses(B, "000000009")
    assert (len(h8), len(h9)) == (6, 7), (len(h8), len(h9))
    print("OK doma_tipa: 6 и 7 домов")

    for hs, key in ((h8, "written_off_15_uah_net"), (h9, "written_off_30_uah_net")):
        rows = q(B, zapros("fakt_spisaniy_domov"), Дома=arr(B, hs), ВидыСубконто=vidy_subkonto(B),
                 БезОтбораНоменклатуры=True, Номенклатура=arr(B, []))
        total = round(sum(r["Сумма"] for r in rows), 2)
        assert abs(total - e0[key]) < 0.01, (key, total, e0[key])
        print(f"OK fakt_spisaniy_domov: {total} = снимок {key}")
    one = q(B, zapros("fakt_spisaniy_domov"), Дома=arr(B, h9), ВидыСубконто=vidy_subkonto(B),
            БезОтбораНоменклатуры=False, Номенклатура=arr(B, [card(B, "00000017651")]))
    assert all(same(B, r["Номенклатура"], card(B, "00000017651")) for r in one) and one, "отбор по карточке"

    s9 = ss(B, "000000009")
    kand = q(B, zapros("kandidaty_kartochek"), СС=s9, ПодразделениеОбъекта=s9.ПодразделениеОбъекта, ВидыСубконто=vidy_subkonto(B))
    green = [r for r in kand if r["Количество"] != 0 and int(r["КарточекОН"] or 0) == 1]
    assert len(kand) == 170 and len(green) == 17, (len(kand), len(green))
    assert all("Единица" in r for r in kand)
    print("OK kandidaty_kartochek: 170 позиций, 17 зелёных")

    cards = [card(B, c) for c in ("00-00010526", "00000017651", "00-00006019", "00-00012879")]
    ceny = q(B, zapros("ceny_posledney_zakupki"), Номенклатура=arr(B, cards), ПодразделениеЗакупки=s9.ПодразделениеЗакупки,
             НаДату=FAR)
    price = {B.String(r["Номенклатура"].Код): round(r["СуммаБезНДС"] / r["Количество"], 2) for r in ceny}
    assert price == {"00-00010526": 2679.16, "00000017651": 2679.17, "00-00006019": 625.0, "00-00012879": 2679.17}, price
    print("OK ceny_posledney_zakupki: септик как в [M5] §4.5")

    gr = B.NewObject("Граница", FAR, B.ВидГраницы.Включая)
    ost = q(B, zapros("ostatki_sklada_proekta"), ПодразделениеЗакупки=s9.ПодразделениеЗакупки, НаДату=gr,
            ВидыСубконто=vidy_subkonto(B), БезОтбораНоменклатуры=True, Номенклатура=arr(B, []))
    total = round(sum(r["Сумма"] or 0 for r in ost), 2)
    assert abs(total - e0["lies_project_uah_net"]) < 0.01, (total, e0["lies_project_uah_net"])
    print(f"OK ostatki_sklada_proekta: {total}")

    units = q(B, zapros("kesh_edinic"))
    first = {}
    for r in units:
        key = B.А_СтруктураСебестоимостиСервер.КлючЕдиницы(r["Наименование"])
        first.setdefault(key, r["Код"])
    assert first["т"] == "0306" and first["шт"] == "2009" and first["м"] == "0101", (first.get("т"), first.get("шт"), first.get("м"))
    print("OK kesh_edinic: т → 0306, шт → 2009, м → 0101")

    docs = q(B, zapros("dokumenty_komplektacii"), СС=s9, ПодразделениеОбъекта=s9.ПодразделениеОбъекта)
    assert any(int(r["НомерСтроки"]) == 170 and "000174" in B.String(r["ДокументКомплектации"]) for r in docs), "ІБ00-000174"
    sost = q(B, zapros("sostav_dokumenta_komplektacii"), ДокументКомплектации=kn(B, "ІБ00-000174"))
    assert sorted(float(r["Количество"]) for r in sost) == [1, 1, 1, 2], [r["Количество"] for r in sost]
    print("OK dokumenty_komplektacii / sostav: ІБ00-000174 → 1/2/1/1")

    s8 = ss(B, "000000008")
    nul = q(B, zapros("nulevye_pozicii"), СС=s8, ПодразделениеОбъекта=s8.ПодразделениеОбъекта, ВидыСубконто=vidy_subkonto(B))
    pos = {}
    for r in nul:
        pos.setdefault(int(r["НомерСтроки"]), False)
        if r["Номенклатура"] is not None:
            pos[int(r["НомерСтроки"])] = True
    assert len(pos) == 22 and sum(pos.values()) == 18, (len(pos), sum(pos.values()))
    print("OK nulevye_pozicii: 22 нулевые, 18 с фактом")

    pod = q(B, zapros("kartochki_dlya_podbora"), ТолькоОН=False, ОбщееНазвание=B.Справочники.ОбщиеНазванияНоменклатуры.ПустаяСсылка(),
            БезПоиска=False, Шаблон="%Мембрана супердифузійна%")
    assert any(B.String(r["Код"]) == "00-00006711" for r in pod), "мембрана вне ОН позиции"
    print(f"OK kartochki_dlya_podbora: {len(pod)} карточек")

    with tx(B):
        o = B.Справочники.А_СпецификацииДомов.СоздатьЭлемент()
        o.Наименование = "__ТЕСТ_П1 СС 009"
        o.ВидСпецификации = B.Перечисления.А_ВидыСпецификацийДомов.Дом
        o.Статус = B.Перечисления.А_СтатусыСпецификацийДомов.Черновик
        o.СтруктураСебестоимости = s9
        o.КоличествоВыпуска = 1
        o.ОбменДанными.Загрузка = True
        o.Записать()
        sp = q(B, zapros("specifikacii_ss"), СсылкаСС=s9)
        assert any(same(B, r["Ссылка"], o.Ссылка) for r in sp)
    print("OK specifikacii_ss")
    print("OK test_e2_zaprosy")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2_zaprosy.py`
Expected: FAIL — `FileNotFoundError: …zaprosy\doma_tipa.sql` (после свежего снимка этапа 0).

- [ ] **Step 3: Написать тексты запросов**

`zaprosy/doma_tipa.sql`:
```
ВЫБРАТЬ
	Подр.Ссылка КАК Ссылка
ИЗ
	Справочник.Подразделения КАК Подр
ГДЕ
	Подр.Родитель = &ПодразделениеОбъекта
	И НЕ Подр.ПометкаУдаления

УПОРЯДОЧИТЬ ПО
	Подр.Код
```

`zaprosy/fakt_spisaniy_domov.sql` (правило E3 §3.3; ОН — разыменованием во второй таблице):
```
ВЫБРАТЬ
	Скл.Ссылка КАК Склад,
	Скл.Подразделение КАК Дом
ПОМЕСТИТЬ ВТ_СкладыДомов
ИЗ
	Справочник.Склады КАК Скл
ГДЕ
	Скл.Подразделение В (&Дома)

ИНДЕКСИРОВАТЬ ПО
	Склад
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	ПС.Ссылка КАК Счет
ПОМЕСТИТЬ ВТ_СчетаЗапасов
ИЗ
	ПланСчетов.Хозрасчетный КАК ПС
ГДЕ
	ПС.Код В ("20", "22", "28")
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	СклД.Дом КАК Дом,
	ВЫРАЗИТЬ(Обор.Субконто1 КАК Справочник.Номенклатура) КАК Номенклатура,
	СУММА(Обор.КоличествоОборотКт) КАК Количество,
	СУММА(Обор.СуммаОборотКт) КАК Сумма
ПОМЕСТИТЬ ВТ_Списано
ИЗ
	РегистрБухгалтерии.Хозрасчетный.Обороты(
			,
			,
			,
			Счет В ИЕРАРХИИ
				(ВЫБРАТЬ
					СЗ.Счет
				ИЗ
					ВТ_СчетаЗапасов КАК СЗ),
			&ВидыСубконто,
			,
			НЕ КорСчет В ИЕРАРХИИ
					(ВЫБРАТЬ
						СЗ.Счет
					ИЗ
						ВТ_СчетаЗапасов КАК СЗ),
			) КАК Обор
		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_СкладыДомов КАК СклД
		ПО Обор.Субконто2 = СклД.Склад
ГДЕ
	Обор.КоличествоОборотКт <> 0
	И (&БезОтбораНоменклатуры
			ИЛИ ВЫРАЗИТЬ(Обор.Субконто1 КАК Справочник.Номенклатура) В (&Номенклатура))

СГРУППИРОВАТЬ ПО
	СклД.Дом,
	ВЫРАЗИТЬ(Обор.Субконто1 КАК Справочник.Номенклатура)
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	Спис.Дом КАК Дом,
	Спис.Номенклатура КАК Номенклатура,
	Спис.Номенклатура.ОбщееНазваниеНоменклатуры КАК ОбщееНазвание,
	Спис.Количество КАК Количество,
	Спис.Сумма КАК Сумма
ИЗ
	ВТ_Списано КАК Спис
```

`zaprosy/kandidaty_kartochek.sql` — текст E3 §3.4 (`E3_tochki_vstraivaniya.md`, строки 149–341) **дословно**, с одной добавленной строкой в последнем `ВЫБРАТЬ` после `Поз.ЕдиницаСС КАК ЕдиницаСС,`:
```
	Поз.Единица КАК Единица,
```
(`ЗаполнитьИзСС` берёт её, когда текст `ЕдиницаСС` не распознан.)

`zaprosy/ceny_posledney_zakupki.sql` (Т2; без `ИМЕЮЩИЕ` — память `kontrolbalansa_imeyu_filter`):
```
ВЫБРАТЬ
	Т.Номенклатура КАК Номенклатура,
	Т.Ссылка КАК Документ,
	Т.Ссылка.Дата КАК Дата,
	СУММА(Т.Количество) КАК Количество,
	СУММА(ВЫБОР
			КОГДА Т.Ссылка.СуммаВключаетНДС
				ТОГДА Т.Сумма - Т.СуммаНДС
			ИНАЧЕ Т.Сумма
		КОНЕЦ) КАК СуммаБезНДС,
	СУММА(ВЫБОР
			КОГДА Т.Ссылка.СуммаВключаетНДС
				ТОГДА Т.Сумма
			ИНАЧЕ Т.Сумма + Т.СуммаНДС
		КОНЕЦ) КАК СуммаСНДС
ПОМЕСТИТЬ ВТ_СтрокиПТУ
ИЗ
	Документ.ПоступлениеТоваровУслуг.Товары КАК Т
ГДЕ
	Т.Ссылка.Проведен
	И Т.Номенклатура В (&Номенклатура)
	И Т.Ссылка.Склад.Подразделение = &ПодразделениеЗакупки
	И Т.Ссылка.Дата <= &НаДату

СГРУППИРОВАТЬ ПО
	Т.Номенклатура,
	Т.Ссылка,
	Т.Ссылка.Дата
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	Стр.Номенклатура КАК Номенклатура,
	МАКСИМУМ(Стр.Дата) КАК Дата
ПОМЕСТИТЬ ВТ_Последние
ИЗ
	ВТ_СтрокиПТУ КАК Стр
ГДЕ
	Стр.Количество > 0

СГРУППИРОВАТЬ ПО
	Стр.Номенклатура
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	Стр.Номенклатура КАК Номенклатура,
	МАКСИМУМ(Стр.Документ) КАК Документ,
	СУММА(Стр.Количество) КАК Количество,
	СУММА(Стр.СуммаБезНДС) КАК СуммаБезНДС,
	СУММА(Стр.СуммаСНДС) КАК СуммаСНДС
ИЗ
	ВТ_СтрокиПТУ КАК Стр
		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_Последние КАК Посл
		ПО Стр.Номенклатура = Посл.Номенклатура
			И Стр.Дата = Посл.Дата
ГДЕ
	Стр.Количество > 0

СГРУППИРОВАТЬ ПО
	Стр.Номенклатура
```

`zaprosy/ostatki_sklada_proekta.sql`:
```
ВЫБРАТЬ
	Скл.Ссылка КАК Склад
ПОМЕСТИТЬ ВТ_СкладыПроекта
ИЗ
	Справочник.Склады КАК Скл
ГДЕ
	Скл.Подразделение = &ПодразделениеЗакупки
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	ПС.Ссылка КАК Счет
ПОМЕСТИТЬ ВТ_СчетаЗапасов
ИЗ
	ПланСчетов.Хозрасчетный КАК ПС
ГДЕ
	ПС.Код В ("20", "22", "28")
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.Номенклатура) КАК Номенклатура,
	СУММА(Ост.КоличествоОстаток) КАК Количество,
	СУММА(Ост.СуммаОстаток) КАК Сумма
ИЗ
	РегистрБухгалтерии.Хозрасчетный.Остатки(
			&НаДату,
			Счет В ИЕРАРХИИ
				(ВЫБРАТЬ
					СЗ.Счет
				ИЗ
					ВТ_СчетаЗапасов КАК СЗ),
			&ВидыСубконто,
			Субконто2 В
				(ВЫБРАТЬ
					СП.Склад
				ИЗ
					ВТ_СкладыПроекта КАК СП)) КАК Ост
ГДЕ
	(&БезОтбораНоменклатуры
			ИЛИ ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.Номенклатура) В (&Номенклатура))

СГРУППИРОВАТЬ ПО
	ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.Номенклатура)
```

`zaprosy/kesh_edinic.sql`:
```
ВЫБРАТЬ
	Ед.Ссылка КАК Ссылка,
	Ед.Код КАК Код,
	Ед.Наименование КАК Наименование,
	Ед.ПометкаУдаления КАК ПометкаУдаления,
	ЕСТЬNULL(Частоты.Частота, 0) КАК Карточек
ИЗ
	Справочник.КлассификаторЕдиницИзмерения КАК Ед
		ЛЕВОЕ СОЕДИНЕНИЕ (ВЫБРАТЬ
			Ном.БазоваяЕдиницаИзмерения КАК Единица,
			КОЛИЧЕСТВО(*) КАК Частота
		ИЗ
			Справочник.Номенклатура КАК Ном
		ГДЕ
			НЕ Ном.ПометкаУдаления
		СГРУППИРОВАТЬ ПО
			Ном.БазоваяЕдиницаИзмерения) КАК Частоты
		ПО Ед.Ссылка = Частоты.Единица

УПОРЯДОЧИТЬ ПО
	ПометкаУдаления,
	Карточек УБЫВ,
	Код
```

`zaprosy/dokumenty_komplektacii.sql` — текст E3 §3.6 (строки 385–403 заметки) дословно.

`zaprosy/sostav_dokumenta_komplektacii.sql` — текст E3 §3.6 (строки 411–423) дословно.

`zaprosy/nulevye_pozicii.sql` (Р22; текст `ТЕКСТ_НУЛЕВЫЕ` пробы `plan_p1/E3/t7_nulevye_komplekt.py`, 04.10.2026: на 15 м — 22 нулевые позиции, у 18 есть факт, Σ факта 147 559,67; перепечатан сюда, чтобы не зависеть от Temp; параметры `&СС`, `&ПодразделениеОбъекта`, `&ВидыСубконто`):
```
ВЫБРАТЬ
	Скл.Ссылка КАК Склад,
	Скл.Подразделение КАК Дом
ПОМЕСТИТЬ ВТ_СкладыДомов
ИЗ
	Справочник.Склады КАК Скл
ГДЕ
	Скл.Подразделение.Родитель = &ПодразделениеОбъекта
ИНДЕКСИРОВАТЬ ПО
	Склад
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	ПС.Ссылка КАК Счет
ПОМЕСТИТЬ ВТ_СчетаЗапасов
ИЗ
	ПланСчетов.Хозрасчетный КАК ПС
ГДЕ
	ПС.Код В ("20", "22", "28")
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	Компл.НомерСтроки КАК НомерСтроки,
	Компл.Этап КАК Этап,
	Компл.НоменклатураСС КАК НоменклатураСС,
	Компл.ОбщееНазвание КАК ОбщееНазвание
ПОМЕСТИТЬ ВТ_Нулевые
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка = &СС
	И Компл.Количество = 0
ИНДЕКСИРОВАТЬ ПО
	ОбщееНазвание
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	ВЫРАЗИТЬ(Обор.Субконто1 КАК Справочник.Номенклатура) КАК Номенклатура,
	СклД.Дом КАК Дом,
	СУММА(Обор.КоличествоОборотКт) КАК Количество,
	СУММА(Обор.СуммаОборотКт) КАК Сумма
ПОМЕСТИТЬ ВТ_СписаноНаДома
ИЗ
	РегистрБухгалтерии.Хозрасчетный.Обороты(
			,
			,
			,
			Счет В ИЕРАРХИИ
				(ВЫБРАТЬ
					СЗ.Счет
				ИЗ
					ВТ_СчетаЗапасов КАК СЗ),
			&ВидыСубконто,
			,
			НЕ КорСчет В ИЕРАРХИИ
					(ВЫБРАТЬ
						СЗ.Счет
					ИЗ
						ВТ_СчетаЗапасов КАК СЗ),
			) КАК Обор
		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_СкладыДомов КАК СклД
		ПО Обор.Субконто2 = СклД.Склад
ГДЕ
	Обор.КоличествоОборотКт <> 0
СГРУППИРОВАТЬ ПО
	ВЫРАЗИТЬ(Обор.Субконто1 КАК Справочник.Номенклатура),
	СклД.Дом
;

////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	Нул.НомерСтроки КАК НомерСтроки,
	Нул.Этап КАК Этап,
	Нул.НоменклатураСС КАК НоменклатураСС,
	Нул.ОбщееНазвание КАК ОбщееНазвание,
	Спис.Номенклатура КАК Номенклатура,
	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Спис.Дом) КАК ДомовСФактом,
	СУММА(Спис.Количество) КАК Количество,
	СУММА(Спис.Сумма) КАК Сумма
ИЗ
	ВТ_Нулевые КАК Нул
		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_СписаноНаДома КАК Спис
		ПО Нул.ОбщееНазвание = Спис.Номенклатура.ОбщееНазваниеНоменклатуры
СГРУППИРОВАТЬ ПО
	Нул.НомерСтроки,
	Нул.Этап,
	Нул.НоменклатураСС,
	Нул.ОбщееНазвание,
	Спис.Номенклатура
УПОРЯДОЧИТЬ ПО
	НомерСтроки
```

`zaprosy/specifikacii_ss.sql`:
```
ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
ГДЕ
	НЕ Спец.ЭтоГруппа
	И НЕ Спец.ПометкаУдаления
	И Спец.СтруктураСебестоимости = &СсылкаСС
	И Спец.ВидСпецификации В (ЗНАЧЕНИЕ(Перечисление.А_ВидыСпецификацийДомов.Дом), ЗНАЧЕНИЕ(Перечисление.А_ВидыСпецификацийДомов.Узел))
	И Спец.Статус <> ЗНАЧЕНИЕ(Перечисление.А_СтатусыСпецификацийДомов.Закрыта)
```

`zaprosy/kartochki_dlya_podbora.sql`:
```
ВЫБРАТЬ ПЕРВЫЕ 200
	Ном.Ссылка КАК Номенклатура,
	Ном.Код КАК Код,
	Ном.БазоваяЕдиницаИзмерения КАК Единица,
	Ном.ОбщееНазваниеНоменклатуры КАК ОбщееНазвание
ИЗ
	Справочник.Номенклатура КАК Ном
ГДЕ
	НЕ Ном.ПометкаУдаления
	И НЕ Ном.ЭтоГруппа
	И (НЕ &ТолькоОН
			ИЛИ Ном.ОбщееНазваниеНоменклатуры = &ОбщееНазвание)
	И (&БезПоиска
			ИЛИ Ном.Наименование ПОДОБНО &Шаблон)

УПОРЯДОЧИТЬ ПО
	Ном.Наименование
```

Для трёх текстов «дословно из заметок» (`_p1_notes/E3_tochki_vstraivaniya.md` закоммичен в Task 1) — скопировать скриптом и проверить, что строки совпали:

```bash
python -X utf8 - <<'PY'
import io, re
E3 = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\budget-calc-oct-dec-2026-20c998\docs\superpowers\plans\_p1_notes\E3_tochki_vstraivaniya.md"
Z = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\budget-calc-oct-dec-2026-20c998\_Rarzrabotki\Python\test\spec_domov\zaprosy"
lines = io.open(E3, encoding="utf-8").read().split("\n")
def block(a, b): return "\n".join(lines[a - 1:b]) + "\n"
kand = block(150, 340)          # строки 149 и 341 — ограждения ``` блока, не текст запроса
assert kand.startswith("ВЫБРАТЬ") and kand.rstrip().endswith("НомерСтроки"), kand[-60:]
anchor = "\tПоз.ЕдиницаСС КАК ЕдиницаСС,\n\tПоз.КоличествоСС_Оригинал"
assert kand.count(anchor) == 1
kand = kand.replace(anchor, "\tПоз.ЕдиницаСС КАК ЕдиницаСС,\n\tПоз.Единица КАК Единица,\n\tПоз.КоличествоСС_Оригинал")
io.open(Z + r"\kandidaty_kartochek.sql", "w", encoding="utf-8").write(kand)
dok = block(385, 403); assert dok.startswith("ВЫБРАТЬ") and "Док.Дата" in dok
io.open(Z + r"\dokumenty_komplektacii.sql", "w", encoding="utf-8").write(dok)
sos = block(411, 423); assert sos.startswith("ВЫБРАТЬ") and "НомерСтроки" in sos
io.open(Z + r"\sostav_dokumenta_komplektacii.sql", "w", encoding="utf-8").write(sos)
print("OK: 3 текста скопированы")
PY
```
`zaprosy/nulevye_pozicii.sql` создать из блока выше (инструментом Write). Expected: `OK: 3 текста скопированы`. (Если номера строк в E3 сдвинулись — найти блоки по заголовкам «### 3.4», «### 3.6» и скорректировать номера; ассерты не дадут скопировать не тот текст.)

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2_zaprosy.py`
Expected: все строки `OK …`, последняя `OK test_e2_zaprosy`. Расхождение `fakt`/`ostatki` со свежим снимком этапа 0 — ошибка запроса (оба считаются в одну минуту).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/zaprosy/*.sql _Rarzrabotki/Python/test/spec_domov/test_e2_zaprosy.py
git commit -m "test(spec_domov): Rule #-1 — 11 запросов этапа 2 (факт домов = Z6, цены Т2, склад проекта, единицы, КН, нулевые)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 30: Единицы этапа 2 — кэш, единица по тексту, подсказка k, размер из названия (Р23, Р25)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_40_edinicy2.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2a_edinicy.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2a_edinicy.py`:

```python
# -*- coding: utf-8 -*-
"""Единицы этапа 2 через COM: ЕдиницаПоТексту (кэш классификатора), ПодсказкаКоэффициента (Р23, E3 §4.2),
РазмерНеПротиворечит (Р25)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect  # noqa: E402
from fixtures import card, srv, unit  # noqa: E402


def main():
    B = connect()
    s = srv(B)
    for text, code in ((" т", "0306"), ("т", "0306"), ("шт.", "2009"), ("пач", "2110"), ("м.п.", "0101"), ("м", "0101"), ("балон", "2009")):
        ref = s.ЕдиницаПоТексту(text)
        assert ref is not None and ref.Код.strip() == code, (text, ref.Код if ref else None)
    assert s.ЕдиницаПоТексту("непонятно") is None
    print("OK ЕдиницаПоТексту")
    m2, m = unit(B, "0123"), unit(B, "0101")
    for code, row_unit, k in (("00-00006711", m2, 75), ("00-00009980", m, 3), ("00-00013067", m, 3),
                              ("00-00021890", m2, 12.2), ("00-00022590", m2, 5), ("00-00022213", m, 1)):
        h = s.ПодсказкаКоэффициента(card(B, code), row_unit)
        assert abs(h.Коэффициент - k) < 1e-9, (code, h.Коэффициент, h.Текст)
    assert s.ПодсказкаКоэффициента(card(B, "00-00010526"), m2).Коэффициент == 0
    print("OK ПодсказкаКоэффициента: мембрана 75, плінтус 3, труба 3, EKOBOARD 12,2, Кнауф 5")
    assert s.РазмерНеПротиворечит("Мінеральна вата (50 мм)", "EKOBOARD 50*610*1250 мм 12.2 м2 1/36")
    assert not s.РазмерНеПротиворечит("Мінеральна вата (50 мм)", "EKOBOARD 100*610*1250 мм 6,1 м2 1/36")
    assert s.РазмерНеПротиворечит("Замок", "Замок врізний 85 мм")
    print("OK РазмерНеПротиворечит")
    print("OK test_e2a_edinicy")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2a_edinicy.py`
Expected: FAIL — `AttributeError: …ЕдиницаПоТексту` (функции этапа 2 нет в базе).

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_40_edinicy2.bsl`:

```bsl
#Область ЕдиницыЭтап2

// Кэш «ключ единицы спецификации → элемент классификатора» (запрос kesh_edinic): на ключ — непомеченный элемент с наибольшим
// числом карточек, затем код (как ПостроитьКэшЕдиниц СС, E3 §1.1): «т» → 0306 « т», «шт» → 2009, «м» → 0101, «пак» → 2110.
Функция КэшЕдиницСпецификации() Экспорт

	Кэш = Новый Соответствие;
	Запрос = Новый Запрос(ТекстЗапросаКэшаЕдиниц());
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Ключ = КлючЕдиницыСпецификации(Выборка.Наименование);
		Если Не ПустаяСтрока(Ключ) И Кэш.Получить(Ключ) = Неопределено Тогда
			Кэш.Вставить(Ключ, Выборка.Ссылка);
		КонецЕсли;
	КонецЦикла;
	Возврат Кэш;

КонецФункции

// Элемент классификатора по тексту единицы СС («пач», « т», «м.п.»): Неопределено — не распознан.
Функция ЕдиницаПоТексту(Знач ТекстЕдиницы, КэшЕдиниц = Неопределено) Экспорт

	Если КэшЕдиниц = Неопределено Тогда
		КэшЕдиниц = КэшЕдиницСпецификации();
	КонецЕсли;
	Ключ = КлючЕдиницыСпецификации(ТекстЕдиницы);
	Если ПустаяСтрока(Ключ) Тогда
		Возврат Неопределено;
	КонецЕсли;
	Возврат КэшЕдиниц.Получить(Ключ);

КонецФункции

// Подсказка k «1 [од. картки] = k [од. рядка]» из названия карточки (Р23, E3 §4.2): «75 м кв», «12,2 м2», «5,0м2» — для строки
// в м2; «3м», «50м.п» и «/3000» (мм) — для строки в м. Возврат: Структура(Коэффициент — 0, если подсказки нет; Текст).
Функция ПодсказкаКоэффициента(Номенклатура, ЕдиницаСтроки) Экспорт

	Результат = Новый Структура("Коэффициент, Текст", 0, "");
	Если Не ЗначениеЗаполнено(Номенклатура) Тогда
		Возврат Результат;
	КонецЕсли;
	Реквизиты = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(Номенклатура, "Наименование, БазоваяЕдиницаИзмерения");
	КлючРядка = КлючЕдиницыСпецификации(НаименованиеЕдиницы(ЕдиницаСтроки));
	Если КлючРядка = КлючЕдиницыСпецификации(НаименованиеЕдиницы(Реквизиты.БазоваяЕдиницаИзмерения)) Тогда
		Результат.Коэффициент = 1;
		Результат.Текст = НСтр("ru = 'единицы совпадают'; uk = 'одиниці збігаються'");
		Возврат Результат;
	КонецЕсли;
	Для Каждого Токен Из ТокеныРазмеров(Реквизиты.Наименование) Цикл
		Если КлючРядка = "м2" И Токен.Суффикс = "м2" Тогда
			Результат.Коэффициент = Токен.Число;
		ИначеЕсли КлючРядка = "м" И (Токен.Суффикс = "м" Или Токен.Суффикс = "мп") Тогда
			Результат.Коэффициент = Токен.Число;
		ИначеЕсли КлючРядка = "м" И Токен.ПослеКосой И Токен.Суффикс = "" И Токен.Число >= 1000 Тогда
			Результат.Коэффициент = Токен.Число / 1000;
		КонецЕсли;
		Если Результат.Коэффициент > 0 Тогда
			Результат.Текст = СтрШаблон("1 %1 = %2 %3", НаименованиеЕдиницы(Реквизиты.БазоваяЕдиницаИзмерения),
				Формат(Результат.Коэффициент, "ЧГ=0"), НаименованиеЕдиницы(ЕдиницаСтроки));
			Прервать;
		КонецЕсли;
	КонецЦикла;
	Возврат Результат;

КонецФункции

// Р25: размер из названия карточки не противоречит тексту позиции СС — у обоих нет значимых чисел (> 1)
// или есть хотя бы одно общее («(50 мм)» и «50*610*1250» — да; «(50 мм)» и «100*610*1250» — нет).
Функция РазмерНеПротиворечит(Знач ТекстПозиции, Знач ИмяКарточки) Экспорт

	ЧислаПозиции = ЗначимыеЧисла(ТекстПозиции);
	ЧислаКарточки = ЗначимыеЧисла(ИмяКарточки);
	Если ЧислаПозиции.Количество() = 0 Или ЧислаКарточки.Количество() = 0 Тогда
		Возврат Истина;
	КонецЕсли;
	Для Каждого ЧислоПозиции Из ЧислаПозиции Цикл
		Если ЧислаКарточки.Найти(ЧислоПозиции) <> Неопределено Тогда
			Возврат Истина;
		КонецЕсли;
	КонецЦикла;
	Возврат Ложь;

КонецФункции

Функция ЗначимыеЧисла(Знач Текст)

	Числа = Новый Массив;
	Для Каждого Токен Из ТокеныРазмеров(Текст) Цикл
		Если Токен.Число > 1 И Числа.Найти(Токен.Число) = Неопределено Тогда
			Числа.Добавить(Токен.Число);
		КонецЕсли;
	КонецЦикла;
	Возврат Числа;

КонецФункции

// Числа названия с единицей сразу после них: «75 м кв» → (75, «м2»), «3м» → (3, «м»), «/3000» → (3000, «», после «/»).
Функция ТокеныРазмеров(Знач Текст)

	Текст = НРег(Текст);
	Текст = СтрЗаменить(Текст, "м.п.", "мп");
	Текст = СтрЗаменить(Текст, "м.п", "мп");
	Текст = СтрЗаменить(Текст, "п.м.", "мп");
	Текст = СтрЗаменить(Текст, "м кв", "м2");
	Текст = СтрЗаменить(Текст, "м.кв", "м2");
	Текст = СтрЗаменить(Текст, "мкв", "м2");
	Текст = СтрЗаменить(Текст, "м²", "м2");
	Токены = Новый Массив;
	Длина = СтрДлина(Текст);
	Поз = 1;
	Пока Поз <= Длина Цикл
		Если Не ЭтоЦифра(Сред(Текст, Поз, 1)) Тогда
			Поз = Поз + 1;
			Продолжить;
		КонецЕсли;
		ПослеКосой = (Поз > 1 И Сред(Текст, Поз - 1, 1) = "/");
		Начало = Поз;
		Пока Поз <= Длина Цикл
			Символ = Сред(Текст, Поз, 1);
			Если ЭтоЦифра(Символ) Или ((Символ = "," Или Символ = ".") И ЭтоЦифра(Сред(Текст, Поз + 1, 1))) Тогда
				Поз = Поз + 1;
			Иначе
				Прервать;
			КонецЕсли;
		КонецЦикла;
		ЧислоТекстом = СтрЗаменить(Сред(Текст, Начало, Поз - Начало), ",", ".");
		Пока Поз <= Длина И Сред(Текст, Поз, 1) = " " Цикл
			Поз = Поз + 1;
		КонецЦикла;
		Суффикс = "";
		Пока Поз <= Длина И СтрДлина(Суффикс) < 4 Цикл
			Символ = Сред(Текст, Поз, 1);
			Если ЭтоБуква(Символ) Или (Символ = "2" И Прав(Суффикс, 1) = "м") Тогда
				Суффикс = Суффикс + Символ;
				Поз = Поз + 1;
			Иначе
				Прервать;
			КонецЕсли;
		КонецЦикла;
		Токены.Добавить(Новый Структура("Число, Суффикс, ПослеКосой", ЧислоИзТекста(ЧислоТекстом), Суффикс, ПослеКосой));
	КонецЦикла;
	Возврат Токены;

КонецФункции

Функция ЧислоИзТекста(Знач Текст)

	Попытка
		Возврат Число(Текст);
	Исключение
		Возврат 0;
	КонецПопытки;

КонецФункции

Функция ТекстЗапросаКэшаЕдиниц()

	Возврат
	//@ЗАПРОС kesh_edinic
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_40_edinicy2.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 33 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_40_edinicy2.bsl _Rarzrabotki/Python/test/spec_domov/test_e2a_edinicy.py
git commit -m "feat(spec_domov): единица по тексту СС, подсказка k из названия карточки (Р23), размер из названия (Р25)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 31: Факт домов, цены последней закупки, склад проекта, кандидаты, документы комплектации (Р31)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_50_fakt.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2a_zapolnit.py` (часть `t_fakt`)

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2a_zapolnit.py`:

```python
# -*- coding: utf-8 -*-
"""Этап 2, партия 2A через COM: факт и цены (t_fakt), «Склад з документа комплектації» Р31 (t_kn), «Заповнити з СС» на 30 м
(t_fill: 170 позиций, 17 зелёных, ни одной жёлтой на ОН с 2+ позициями), «Позначити зелені перевіреними», «Взяти з оригіналу
Excel» (кутник 10 м), подбор вне ОН позиции (Р35)."""
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import OUT, arr, connect, q, same, tx  # noqa: E402
from fixtures import card, kn, new_spec, row_index_by_ss, srv, ss  # noqa: E402


def e0():
    return json.load(open(sorted(glob.glob(os.path.join(OUT, "e0_snimok_*.json")))[-1], encoding="utf-8"))["control_sums"]


def t_fakt(B):
    s = srv(B)
    s9, s8 = ss(B, "000000009"), ss(B, "000000008")
    d9, d8 = s.ДомаТипа(s9), s.ДомаТипа(s8)
    assert (d9.Количество(), d8.Количество()) == (7, 6)
    f8 = s.ФактСписанийДома(d8)
    total = round(sum(f8.Получить(i).Сумма for i in range(f8.Количество())), 2)
    assert abs(total - e0()["written_off_15_uah_net"]) < 0.01, total
    ceny = s.ЦеныПоследнейЗакупки(arr(B, [card(B, "00-00010526")]), s9.ПодразделениеЗакупки)
    c = ceny.Получить(card(B, "00-00010526"))
    assert round(c.ЦенаБезНДС, 2) == 2679.16 and c.ЦенаСНДС > c.ЦенаБезНДС
    ost = s.ОстаткиСкладаПроекта(s9.ПодразделениеЗакупки)
    assert ost.Количество() > 0
    kand = s.КандидатыКарточек(s9)
    assert kand.Количество() == 170
    print(f"OK факт 15 м {total}, цена днища 2679,16, склад проекта карточек {ost.Количество()}")


def t_kn(B):
    s = srv(B)
    sost = s.СоставИзДокументаКомплектации(kn(B, "ІБ00-000174"))
    assert round(sost.КоличествоВыпускаДокумента, 3) == 0.5
    qty = sorted(round(sost.Строки.Получить(i).Количество, 6) for i in range(sost.Строки.Количество()))
    assert qty == [1, 1, 1, 2], qty
    half = s.СоставИзДокументаКомплектации(kn(B, "ІБ00-000174"), 2)
    assert sorted(round(half.Строки.Получить(i).Количество, 6) for i in range(half.Строки.Количество())) == [0.5, 0.5, 0.5, 1]
    docs = s.ДокументыКомплектацииПозиций(ss(B, "000000009"))
    assert any(int(docs.Получить(i).НомерСтроки) == 170 for i in range(docs.Количество()))
    print("OK Р31: ІБ00-000174 → 1/2/1/1 (выпуск 0,5 не делит состав)")


def t_fill(B):
    s = srv(B)
    with tx(B):
        h = new_spec(B, "Дом", "еталон 30 м", ss_ref=ss(B, "000000009"))
        p = s.ЗаполнитьИзСС(h)
        assert (p.Строк, p.Зеленых, p.ПропущеноНулевых) == (170, 17, 0), (p.Строк, p.Зеленых, p.ПропущеноНулевых)
        assert round(p.Сумма, 2) == 1009864.74 and p.Групп == 9, (p.Сумма, p.Групп)
        kand = s.КандидатыКарточек(ss(B, "000000009"))
        on_positions = {}
        for i in range(kand.Количество()):
            r = kand.Получить(i)
            if r.Количество != 0:
                on_positions[B.String(r.ОбщееНазвание)] = on_positions.get(B.String(r.ОбщееНазвание), 0) + 1
        yellow = [h.Состав.Получить(i) for i in range(h.Состав.Количество())
                  if B.String(h.Состав.Получить(i).ИсточникКарточки) == "Списувалась на будинки типу"]
        assert yellow and all(on_positions[B.String(r.ОбщееНазвание)] == 1 for r in yellow), "жёлтая на ОН с 2+ позициями"
        marked = s.ОтметитьЗеленыеПроверенными(h)
        assert 0 < marked <= 17
        i13 = row_index_by_ss(B, h, "000000009", 13)
        h.Состав.Получить(i13).Номенклатура = card(B, "00-00022213")
        assert s.ВзятьИзОригиналаExcel(h, i13)
        r13 = h.Состав.Получить(i13)
        assert round(r13.Количество, 6) == 10 and r13.Коэффициент == 1 and B.String(r13.ИсточникКоэффициента) == "Одиниці збігаються", \
            (r13.Количество, r13.Коэффициент, B.String(r13.ИсточникКоэффициента))
        print(f"OK «Заповнити з СС»: 170 / 17 зелёных / жёлтых {len(yellow)} (все на ОН с одной позицией) / отмечено {marked}; кутник 10 м")


def t_podbor(B):
    s = srv(B)
    s9 = ss(B, "000000009")
    on42 = q(B, "ВЫБРАТЬ К.ОбщееНазвание КАК ОН ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К ГДЕ К.Ссылка = &С И К.НомерСтроки = 42", С=s9)[0]["ОН"]
    only_on = s.ПредложитьКарточки(s9, on42, "Вітробарєр", True, "")
    assert not any(same(B, only_on.Получить(i).Номенклатура, card(B, "00-00006711")) for i in range(only_on.Количество()))
    wide = s.ПредложитьКарточки(s9, on42, "Вітробарєр", False, "Мембрана супердифузійна")
    rows = [wide.Получить(i) for i in range(wide.Количество())]
    assert any(same(B, r.Номенклатура, card(B, "00-00006711")) for r in rows), "подбор вне ОН (E3 §3.5)"
    qty = [r.СписаноКоличество for r in rows]
    assert qty == sorted(qty, reverse=True), "сортировка по факту"
    print("OK подбор: мембрана находится только вне ОН позиции")


if __name__ == "__main__":
    B = connect()
    t_fakt(B)
    t_kn(B)
    t_fill(B)
    t_podbor(B)
    print("OK test_e2a_zapolnit")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2a_zapolnit.py`
Expected: FAIL — `AttributeError: …ДомаТипа`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_50_fakt.bsl`:

```bsl
#Область ФактИЦены

// Дома типа — подразделения-дети СС.ПодразделениеОбъекта (E3 §3.1). Возврат: Массив.
Функция ДомаТипа(СсылкаСС) Экспорт

	Запрос = Новый Запрос(ТекстЗапросаДомовТипа());
	Запрос.УстановитьПараметр("ПодразделениеОбъекта", ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "ПодразделениеОбъекта"));
	Возврат Запрос.Выполнить().Выгрузить().ВыгрузитьКолонку("Ссылка");

КонецФункции

// «Списано на дом» (E3 §3.3; общий для П1, П2 и регистратора П4): Кт 20/22/28 со складов домов, корсчёт не 20/22/28,
// за всё время; склад — через ВидыСубконто [Номенклатура, Склады] (E3 §3.2). Сверено с Z6 до копейки (Task 29).
// Дома — Массив; ОтборНоменклатуры — Массив или Неопределено.
// Возврат: ТаблицаЗначений(Дом, Номенклатура, ОбщееНазвание, Количество, Сумма).
Функция ФактСписанийДома(Дома, ОтборНоменклатуры = Неопределено) Экспорт

	Запрос = Новый Запрос(ТекстЗапросаФактаСписаний());
	Запрос.УстановитьПараметр("Дома", Дома);
	Запрос.УстановитьПараметр("ВидыСубконто", ВидыСубконтоНоменклатураСклад());
	Запрос.УстановитьПараметр("БезОтбораНоменклатуры", ОтборНоменклатуры = Неопределено);
	Запрос.УстановитьПараметр("Номенклатура", ?(ОтборНоменклатуры = Неопределено, Новый Массив, ОтборНоменклатуры));
	Возврат Запрос.Выполнить().Выгрузить();

КонецФункции

// Последняя закупка карточки (Т2): ПТУ на склады подразделения закупки СС, средневзвешенно по документу.
// Возврат: Соответствие — Номенклатура → Структура(ЦенаБезНДС, ЦенаСНДС, Документ).
Функция ЦеныПоследнейЗакупки(МассивНоменклатуры, ПодразделениеЗакупки, НаДату = Неопределено) Экспорт

	Цены = Новый Соответствие;
	Если МассивНоменклатуры.Количество() = 0 Тогда
		Возврат Цены;
	КонецЕсли;
	Запрос = Новый Запрос(ТекстЗапросаЦенПоследнейЗакупки());
	Запрос.УстановитьПараметр("Номенклатура", МассивНоменклатуры);
	Запрос.УстановитьПараметр("ПодразделениеЗакупки", ПодразделениеЗакупки);
	Запрос.УстановитьПараметр("НаДату", ?(НаДату = Неопределено, ТекущаяДатаСеанса(), НаДату));
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Если Выборка.Количество > 0 Тогда
			Цены.Вставить(Выборка.Номенклатура, Новый Структура("ЦенаБезНДС, ЦенаСНДС, Документ",
				Окр(Выборка.СуммаБезНДС / Выборка.Количество, 6), Окр(Выборка.СуммаСНДС / Выборка.Количество, 6), Выборка.Документ));
		КонецЕсли;
	КонецЦикла;
	Возврат Цены;

КонецФункции

// Остаток склада проекта (Склады.Подразделение = СС.ПодразделениеЗакупки), счета 20/22/28.
// Возврат: Соответствие — Номенклатура → Структура(Количество, Сумма).
Функция ОстаткиСкладаПроекта(ПодразделениеЗакупки, МассивНоменклатуры = Неопределено, НаДату = Неопределено) Экспорт

	Запрос = Новый Запрос(ТекстЗапросаОстатковСкладаПроекта());
	Запрос.УстановитьПараметр("ПодразделениеЗакупки", ПодразделениеЗакупки);
	Запрос.УстановитьПараметр("НаДату", Новый Граница(?(НаДату = Неопределено, ТекущаяДатаСеанса(), НаДату), ВидГраницы.Включая));
	Запрос.УстановитьПараметр("ВидыСубконто", ВидыСубконтоНоменклатураСклад());
	Запрос.УстановитьПараметр("БезОтбораНоменклатуры", МассивНоменклатуры = Неопределено);
	Запрос.УстановитьПараметр("Номенклатура", ?(МассивНоменклатуры = Неопределено, Новый Массив, МассивНоменклатуры));
	Остатки = Новый Соответствие;
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Остатки.Вставить(Выборка.Номенклатура, Новый Структура("Количество, Сумма", Выборка.Количество, Выборка.Сумма));
	КонецЦикла;
	Возврат Остатки;

КонецФункции

// Кандидаты карточек по позициям СС (E3 §3.4): карточки ОН, факт списаний на дома типа, число ненулевых позиций ОН.
Функция КандидатыКарточек(СсылкаСС) Экспорт

	Запрос = Новый Запрос(ТекстЗапросаКандидатов());
	Запрос.УстановитьПараметр("СС", СсылкаСС);
	Запрос.УстановитьПараметр("ПодразделениеОбъекта", ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "ПодразделениеОбъекта"));
	Запрос.УстановитьПараметр("ВидыСубконто", ВидыСубконтоНоменклатураСклад());
	Возврат Запрос.Выполнить().Выгрузить();

КонецФункции

// Позиции-комплекты (E3 §3.6): документы комплектации, выпуск которых на складах домов типа лежит в ОН позиции.
// Возврат: ТЗ(НомерСтроки, НоменклатураСС, ДокументКомплектации, НоменклатураВузла, КоличествоВыпуска, Проведен, Дом).
Функция ДокументыКомплектацииПозиций(СсылкаСС) Экспорт

	Запрос = Новый Запрос(ТекстЗапросаДокументовКомплектации());
	Запрос.УстановитьПараметр("СС", СсылкаСС);
	Запрос.УстановитьПараметр("ПодразделениеОбъекта", ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "ПодразделениеОбъекта"));
	Возврат Запрос.Выполнить().Выгрузить();

КонецФункции

// «Склад з документа комплектації» (Р31): состав НЕ делится на выпуск документа сам — только на УзловВДокументе
// («склад документа = N вузлів», по умолчанию 1). ІБ00-000174: выпуск 0,5, детали 1/2/1/1.
// Возврат: Структура(КоличествоВыпускаДокумента, НоменклатураВузла, Строки — ТЗ(НомерСтроки, Номенклатура, ОбщееНазвание,
// Единица, Количество, Сумма)).
Функция СоставИзДокументаКомплектации(ДокументКомплектации, УзловВДокументе = 1) Экспорт

	Реквизиты = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(ДокументКомплектации, "Количество, Номенклатура");
	Запрос = Новый Запрос(ТекстЗапросаСоставаДокументаКомплектации());
	Запрос.УстановитьПараметр("ДокументКомплектации", ДокументКомплектации);
	Строки = Запрос.Выполнить().Выгрузить();
	Строки.Колонки.Добавить("КоличествоНаУзел", Новый ОписаниеТипов("Число"));
	Делитель = ?(УзловВДокументе > 0, УзловВДокументе, 1);
	Для Каждого СтрокаДетали Из Строки Цикл
		СтрокаДетали.КоличествоНаУзел = СтрокаДетали.Количество / Делитель;
	КонецЦикла;
	Строки.Колонки.Удалить("Количество");
	Строки.Колонки.КоличествоНаУзел.Имя = "Количество";
	Возврат Новый Структура("КоличествоВыпускаДокумента, НоменклатураВузла, Строки", Реквизиты.Количество, Реквизиты.Номенклатура, Строки);

КонецФункции

Функция ВидыСубконтоНоменклатураСклад()

	Виды = Новый Массив;
	Виды.Добавить(ПланыВидовХарактеристик.ВидыСубконтоХозрасчетные.Номенклатура);
	Виды.Добавить(ПланыВидовХарактеристик.ВидыСубконтоХозрасчетные.Склады);
	Возврат Виды;

КонецФункции

Функция ТекстЗапросаДомовТипа()

	Возврат
	//@ЗАПРОС doma_tipa
	;

КонецФункции

Функция ТекстЗапросаФактаСписаний()

	Возврат
	//@ЗАПРОС fakt_spisaniy_domov
	;

КонецФункции

Функция ТекстЗапросаЦенПоследнейЗакупки()

	Возврат
	//@ЗАПРОС ceny_posledney_zakupki
	;

КонецФункции

Функция ТекстЗапросаОстатковСкладаПроекта()

	Возврат
	//@ЗАПРОС ostatki_sklada_proekta
	;

КонецФункции

Функция ТекстЗапросаКандидатов()

	Возврат
	//@ЗАПРОС kandidaty_kartochek
	;

КонецФункции

Функция ТекстЗапросаДокументовКомплектации()

	Возврат
	//@ЗАПРОС dokumenty_komplektacii
	;

КонецФункции

Функция ТекстЗапросаСоставаДокументаКомплектации()

	Возврат
	//@ЗАПРОС sostav_dokumenta_komplektacii
	;

КонецФункции

#КонецОбласти
```

(Колонка `КоличествоНаУзел` вместо деления «на месте»: у колонки результата запроса квалификатор Число(15,3), и 1/3 обрезалась бы.)

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_50_fakt.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 33 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_50_fakt.bsl _Rarzrabotki/Python/test/spec_domov/test_e2a_zapolnit.py
git commit -m "feat(spec_domov): факт списаний на дома (= Z6), цены последней закупки Т2, склад проекта, кандидаты, состав КН (Р31)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 32: «Заповнити з СС», «Позначити зелені перевіреними», «Взяти з оригіналу Excel», подбор карточки (Р25, Р35)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_55_zapolnenie.bsl`
- Test: `test_e2a_zapolnit.py::t_fill`, `::t_podbor` (Task 31)

- [ ] **Step 1: Тест уже написан.** **Step 2: FAIL** — `AttributeError`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_55_zapolnenie.bsl`:

```bsl
#Область ЗаполнениеЭталона

// «Заповнити з СС» ([M5] §3.1, Р10, Р22, Р25): строка на каждую ненулевую позицию СС (нулевые — в протокол, их список —
// НулевыеПозицииСФактом). Количество и единица — из СС (единица — по ТЕКСТУ ЕдиницаСС: у 30 м поз. 42 ссылка «рул», а 81,975 — м²).
// Карточка только из базы: у ОН одна активная — зелёная «Єдина картка ОН»; на дома типа списывалась ровно одна карточка ОН,
// ОН стоит на одной ненулевой позиции и размер из названия не противоречит — жёлтая «Списувалась на будинки типу»; иначе пусто.
// ОН из списка «групою» (по умолчанию «Гвинтові кріплення») — строка «Група за загальною назвою». Состав очищается.
// Возврат: Структура(Строк, Зеленых, Желтых, Пустых, Групп, ПропущеноНулевых, Сумма, Комплекты — Массив строк «поз. N: документ»).
Функция ЗаполнитьИзСС(СпецОбъект, ОбщиеНазванияГруппой = Неопределено) Экспорт

	СсылкаСС = СпецОбъект.СтруктураСебестоимости;
	Если Не ЗначениеЗаполнено(СсылкаСС) Тогда
		ВызватьИсключение НСтр("ru = 'Не указан тип дома (СС).'; uk = 'Не вказано тип будинку (СС).'");
	КонецЕсли;
	Если ОбщиеНазванияГруппой = Неопределено Тогда
		ОбщиеНазванияГруппой = Новый Массив;
		Крепления = Справочники.ОбщиеНазванияНоменклатуры.НайтиПоНаименованию("Гвинтові кріплення", Истина);
		Если ЗначениеЗаполнено(Крепления) Тогда
			ОбщиеНазванияГруппой.Добавить(Крепления);
		КонецЕсли;
	КонецЕсли;
	Комплекты = Новый Соответствие;
	Для Каждого СтрокаКН Из ДокументыКомплектацииПозиций(СсылкаСС) Цикл
		ДокументыКН = Комплекты.Получить(СтрокаКН.НомерСтроки);
		Если ДокументыКН = Неопределено Тогда
			ДокументыКН = Новый Массив;
			Комплекты.Вставить(СтрокаКН.НомерСтроки, ДокументыКН);
		КонецЕсли;
		ДокументыКН.Добавить(Строка(СтрокаКН.ДокументКомплектации));
	КонецЦикла;
	КэшЕдиниц = КэшЕдиницСпецификации();
	ВидыСтрок = Перечисления.А_ВидыСтрокСпецификацииДомов;
	Источники = Перечисления.А_ИсточникиКарточкиСпецификации;
	Протокол = Новый Структура("Строк, Зеленых, Желтых, Пустых, Групп, ПропущеноНулевых, Сумма, Комплекты",
		0, 0, 0, 0, 0, 0, 0, Новый Массив);
	СпецОбъект.Состав.Очистить();
	Для Каждого Позиция Из КандидатыКарточек(СсылкаСС) Цикл
		Если Позиция.Количество = 0 Тогда
			Протокол.ПропущеноНулевых = Протокол.ПропущеноНулевых + 1;
			Продолжить;
		КонецЕсли;
		Стр = СпецОбъект.Состав.Добавить();
		Стр.КлючСтроки = Новый УникальныйИдентификатор;
		Стр.ЭтапСС = Позиция.Этап;
		Стр.ПозицияСС = Позиция.НоменклатураСС;
		Стр.ОбщееНазвание = Позиция.ОбщееНазвание;
		Стр.Количество = Позиция.Количество;
		Стр.ЕдиницаСтроки = ЕдиницаПоТексту(Позиция.ЕдиницаСС, КэшЕдиниц);
		Если Не ЗначениеЗаполнено(Стр.ЕдиницаСтроки) Тогда
			Стр.ЕдиницаСтроки = Позиция.Единица;
		КонецЕсли;
		Стр.Проверено = Ложь;
		Протокол.Строк = Протокол.Строк + 1;
		Протокол.Сумма = Протокол.Сумма + Позиция.Сумма;
		Если ОбщиеНазванияГруппой.Найти(Позиция.ОбщееНазвание) <> Неопределено Тогда
			Стр.ВидСтроки = ВидыСтрок.ГруппаПоОбщемуНазванию;
			Протокол.Групп = Протокол.Групп + 1;
		Иначе
			Стр.ВидСтроки = ВидыСтрок.Деталь;
			Если Позиция.КарточекОН = 1 Тогда
				Стр.Номенклатура = Позиция.ЕдинственнаяКарточка;
				Стр.ИсточникКарточки = Источники.ЕдинственнаяКарточкаОН;
				Протокол.Зеленых = Протокол.Зеленых + 1;
			ИначеЕсли Позиция.КарточекСФактом = 1 И Позиция.НенулевыхПозицийОН = 1
				И РазмерНеПротиворечит(Позиция.НоменклатураСС, Строка(Позиция.КарточкаФакта)) Тогда
				Стр.Номенклатура = Позиция.КарточкаФакта;
				Стр.ИсточникКарточки = Источники.СписаннаяНаДомаТипа;
				Протокол.Желтых = Протокол.Желтых + 1;
			Иначе
				Протокол.Пустых = Протокол.Пустых + 1;
			КонецЕсли;
		КонецЕсли;
		УстановитьКоэффициентПоЕдиницам(Стр);
		ДокументыКН = Комплекты.Получить(Позиция.НомерСтроки);
		Если ДокументыКН <> Неопределено Тогда
			Протокол.Комплекты.Добавить(СтрШаблон(НСтр("ru = 'поз. %1 «%2»: %3'; uk = 'поз. %1 «%2»: %3'"),
				Позиция.НомерСтроки, Позиция.НоменклатураСС, СтрСоединить(ДокументыКН, ", ")));
		КонецЕсли;
	КонецЦикла;
	Возврат Протокол;

КонецФункции

// «Позначити зелені перевіреними» (Р35): строки «Деталь» с единственной карточкой ОН и совпавшими единицами. Возврат: число.
Функция ОтметитьЗеленыеПроверенными(СпецОбъект) Экспорт

	Отмечено = 0;
	Для Каждого Стр Из СпецОбъект.Состав Цикл
		Если Стр.ВидСтроки = Перечисления.А_ВидыСтрокСпецификацииДомов.Деталь
			И Стр.ИсточникКарточки = Перечисления.А_ИсточникиКарточкиСпецификации.ЕдинственнаяКарточкаОН
			И Стр.ИсточникКоэффициента = Перечисления.А_ИсточникиКоэффициентаПересчетаСС.ЕдиницыСовпадают
			И Не Стр.Проверено Тогда
			Стр.Проверено = Истина;
			Отмечено = Отмечено + 1;
		КонецЕсли;
	КонецЦикла;
	Возврат Отмечено;

КонецФункции

// «Взяти з оригіналу Excel» ([M5] §1.2 п. 5): количество и единица строки — из КоличествоСС_Оригинал / ЕдиницаСС_Оригинал
// позиции; k — по единицам. Единицу сама форма не меняет. Возврат: Ложь — у позиции нет оригинала.
Функция ВзятьИзОригиналаExcel(СпецОбъект, ИндексСтроки) Экспорт

	Стр = СпецОбъект.Состав[ИндексСтроки];
	Позиция = ПозицииСтруктурыСебестоимости(СпецОбъект.СтруктураСебестоимости).Получить(
		А_СтруктураСебестоимостиСервер.КлючПозиции(Стр.ЭтапСС, Стр.ПозицияСС));
	Если Позиция = Неопределено Или Позиция.КоличествоОригинал = 0 Или ПустаяСтрока(Позиция.ЕдиницаОригинал) Тогда
		Возврат Ложь;
	КонецЕсли;
	ЕдиницаОригинала = ЕдиницаПоТексту(Позиция.ЕдиницаОригинал);
	Если ЕдиницаОригинала = Неопределено Тогда
		Возврат Ложь;
	КонецЕсли;
	Стр.Количество = Позиция.КоличествоОригинал;
	Стр.ЕдиницаСтроки = ЕдиницаОригинала;
	Стр.Проверено = Ложь;
	УстановитьКоэффициентПоЕдиницам(Стр);
	Возврат Истина;

КонецФункции

// Подбор карточки с контекстом (Р35): карточки ОН позиции или — при ТолькоОН = Ложь — по строке поиска во всей номенклатуре
// (карточки приёмки лежат в чужой ОН, E3 §3.5); колонки факта по домам типа, склада проекта, цены последней закупки з ПДВ,
// «розмір з назви збігся»; сортировка по факту. Возврат: ТЗ с колонками формы подбора.
Функция ПредложитьКарточки(СсылкаСС, ОбщееНазвание, Знач ПозицияСС = "", ТолькоОН = Истина, Знач СтрокаПоиска = "") Экспорт

	Запрос = Новый Запрос(ТекстЗапросаКарточекДляПодбора());
	Запрос.УстановитьПараметр("ТолькоОН", ТолькоОН И ЗначениеЗаполнено(ОбщееНазвание));
	Запрос.УстановитьПараметр("ОбщееНазвание", ОбщееНазвание);
	Поиск = СокрЛП(СтрЗаменить(СтрЗаменить(СтрЗаменить(СтрокаПоиска, "%", ""), "_", ""), "[", ""));
	Запрос.УстановитьПараметр("БезПоиска", ПустаяСтрока(Поиск));
	Запрос.УстановитьПараметр("Шаблон", "%" + Поиск + "%");
	Карточки = Запрос.Выполнить().Выгрузить();
	Карточки.Колонки.Добавить("СписаноКоличество", Новый ОписаниеТипов("Число"));
	Карточки.Колонки.Добавить("СписаноСумма", Новый ОписаниеТипов("Число"));
	Карточки.Колонки.Добавить("ДомовСФактом", Новый ОписаниеТипов("Число"));
	Карточки.Колонки.Добавить("ФактПоДомам", Новый ОписаниеТипов("Строка"));
	Карточки.Колонки.Добавить("СкладПроекта", Новый ОписаниеТипов("Число"));
	Карточки.Колонки.Добавить("ЦенаПоследнейЗакупки", Новый ОписаниеТипов("Число"));
	Карточки.Колонки.Добавить("РазмерСовпал", Новый ОписаниеТипов("Булево"));
	Если Карточки.Количество() = 0 Тогда
		Возврат Карточки;
	КонецЕсли;
	МассивКарточек = Карточки.ВыгрузитьКолонку("Номенклатура");
	ПодразделениеЗакупки = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "ПодразделениеЗакупки");
	Факт = ФактСписанийДома(ДомаТипа(СсылкаСС), МассивКарточек);
	Цены = ЦеныПоследнейЗакупки(МассивКарточек, ПодразделениеЗакупки);
	Остатки = ОстаткиСкладаПроекта(ПодразделениеЗакупки, МассивКарточек);
	Для Каждого Карточка Из Карточки Цикл
		ПоДомам = Новый Массив;
		СтрокиФакта = Факт.НайтиСтроки(Новый Структура("Номенклатура", Карточка.Номенклатура));
		Для Каждого СтрокаФакта Из СтрокиФакта Цикл
			Карточка.СписаноКоличество = Карточка.СписаноКоличество + СтрокаФакта.Количество;
			Карточка.СписаноСумма = Карточка.СписаноСумма + СтрокаФакта.Сумма;
			ПоДомам.Добавить(СтрШаблон("%1: %2", КороткоеИмяДома(СтрокаФакта.Дом), СтрокаФакта.Количество));
		КонецЦикла;
		Карточка.ДомовСФактом = СтрокиФакта.Количество();
		Карточка.ФактПоДомам = СтрСоединить(ПоДомам, "; ");
		Цена = Цены.Получить(Карточка.Номенклатура);
		Если Цена <> Неопределено Тогда
			Карточка.ЦенаПоследнейЗакупки = Цена.ЦенаСНДС;
		КонецЕсли;
		Остаток = Остатки.Получить(Карточка.Номенклатура);
		Если Остаток <> Неопределено Тогда
			Карточка.СкладПроекта = Остаток.Количество;
		КонецЕсли;
		Карточка.РазмерСовпал = ПустаяСтрока(ПозицияСС) Или РазмерНеПротиворечит(ПозицияСС, Строка(Карточка.Номенклатура));
	КонецЦикла;
	Карточки.Сортировать("СписаноКоличество Убыв, Код");
	Возврат Карточки;

КонецФункции

Функция КороткоеИмяДома(Дом)

	Имя = Строка(Дом);
	Номер = СтрНайти(Имя, "№");
	Возврат ?(Номер > 0, Сред(Имя, Номер), Имя);

КонецФункции

Функция ТекстЗапросаКарточекДляПодбора()

	Возврат
	//@ЗАПРОС kartochki_dlya_podbora
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_55_zapolnenie.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 33 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_55_zapolnenie.bsl
git commit -m "feat(spec_domov): «Заповнити з СС» (Р10, Р25), «Позначити зелені перевіреними», «Взяти з оригіналу Excel», подбор Р35

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 33: Загрузка 2A и прогон тестов партии

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/load_lists/e2a.txt` (одна строка `CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl`)

- [ ] **Step 1: Тесты** — `test_e2a_edinicy.py`, `test_e2a_zapolnit.py`. **Step 2: FAIL** — `AttributeError` (подтверждено в задачах 30–32).

- [ ] **Step 3: Сборка, статика, копия, загрузка**

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/build_bsl.py --partiya 2a
python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2a.txt --tests
python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2a.txt
powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e2a.txt -Tag e2a
```
Expected: `частей 9`; `ошибок: 0`; `скопирован CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl`; `OK load e2a`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2a_edinicy.py && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2a_zapolnit.py && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_priemka.py`
Expected: `OK test_e2a_edinicy`, `OK test_e2a_zapolnit` (170 / 17 зелёных / жёлтые только на ОН с одной позицией; кутник 10 м; ІБ00-000174 → 1/2/1/1; мембрана — только вне ОН), регресс этапа 1 `ПРОВАЛОВ: 0`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/load_lists/e2a.txt _Rarzrabotki/Python/test/spec_domov/sync_manifest.json \
  _Rarzrabotki/Python/test/spec_domov/out/e2a_load_*.txt \
  "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"
git commit -m "feat(BASEBuh): партия 2A загружена — единицы, факт и цены, «Заповнити з СС», подбор карточки

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 34: Узлы строителя — создать, объединить, расформировать, из документа комплектации, узел в узле (§5, §9, Р31)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_60_uzly.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2b_uzly.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2b_uzly.py`:

```python
# -*- coding: utf-8 -*-
"""Узлы через COM в транзакции: «Склад з документа комплектації» ІБ00-000174 → септик 1/2/1/1 (Р31), узел в строку позиции 170,
«Обʼєднати у вузол» водостік (поз. 67–74) и «Зовнішні мережі» ⊃ «Септик» (узел в узле), «Розформувати» — норма по явным позициям
не меняется ни на одном шаге; «Входимість»; имена узлов строителя (§9)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, same, tx  # noqa: E402
from fixtures import kn, new_spec, row_index_by_ss, srv, ss, unit, write  # noqa: E402


def explicit_norm(B, spec):
    r = srv(B).РазвернутьДоКарточек(spec, True)
    m = B.А_СтруктураСебестоимостиСервер
    out = {}
    for i in range(r.ЯвныеПозиции.Количество()):
        x = r.ЯвныеПозиции.Получить(i)
        k = m.КлючПозиции(x.ЭтапСС, x.ПозицияСС)
        out[k] = round(out.get(k, 0) + x.КоличествоЕдСтроки, 6)
    return out


def row_index_by_node(B, obj, node):
    for i in range(obj.Состав.Количество()):
        if obj.Состав.Получить(i).Узел is not None and same(B, obj.Состав.Получить(i).Узел, node):
            return i
    raise AssertionError("нет строки узла")


def main():
    B = connect()
    s = srv(B)
    names = [s.ИменаУзловСтроителя().Получить(i) for i in range(s.ИменаУзловСтроителя().Количество())]
    assert len(names) == 21 and "Септик" in names and "Кріплення загального призначення" in names
    with tx(B):
        s9 = ss(B, "000000009")
        kompl = unit(B, "2938")
        h = new_spec(B, "Дом", "вузли 30 м", ss_ref=s9)
        s.ЗаполнитьИзСС(h)
        node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
        lv = s.РазвернутьДоКарточек(node).Листья
        assert sorted(round(lv.Получить(i).Количество, 6) for i in range(lv.Количество())) == [1, 1, 1, 2], "не 2/4/2/2"
        s.ПоставитьУзелВСтроку(h, row_index_by_ss(B, h, "000000009", 170), node)
        write(h)
        before = explicit_norm(B, h.Ссылка)
        idx = [row_index_by_ss(B, h, "000000009", n) for n in range(67, 75)]
        vod = s.ОбъединитьВУзел(h, arr(B, idx), "Водостік 30 м", 1, kompl)
        write(h)
        assert explicit_norm(B, h.Ссылка) == before, "водостік изменил норму"
        idx = [row_index_by_node(B, h, node)] + [row_index_by_ss(B, h, "000000009", n) for n in (169, 145, 142)]
        zm = s.ОбъединитьВУзел(h, arr(B, idx), "Зовнішні мережі", 1, kompl)
        write(h)
        assert explicit_norm(B, h.Ссылка) == before, "узел в узле изменил норму"
        vh = s.Входимость(node)
        lv = {B.String(vh.Получить(i).Спецификация): vh.Получить(i).Уровень for i in range(vh.Количество())}
        assert lv == {"Зовнішні мережі": 1, "__ТЕСТ_П1 вузли 30 м": 2}, lv
        r = s.РазвернутьДоКарточек(h.Ссылка)
        sep = [r.Листья.Получить(i) for i in range(r.Листья.Количество()) if r.Листья.Получить(i).Уровень == 2]
        assert len(sep) == 4, "детали септика на уровне 2"
        assert s.РасформироватьУзел(h, row_index_by_node(B, h, zm))
        write(h)
        assert explicit_norm(B, h.Ссылка) == before, "расформирование изменило норму"
        assert s.РасформироватьУзел(h, row_index_by_node(B, h, vod))
        write(h)
        assert explicit_norm(B, h.Ссылка) == before
    print("OK test_e2b_uzly: Р31 1/2/1/1, объединение/расформирование без изменения нормы, узел в узле, входимость")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2b_uzly.py`
Expected: FAIL — `AttributeError: …ИменаУзловСтроителя`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_60_uzly.bsl`:

```bsl
#Область УзлыСтроителя

// Имена узлов строителя (§9) — подсказка «Обʼєднати у вузол» (Р35).
Функция ИменаУзловСтроителя() Экспорт

	Возврат СтрРазделить("Каркас модуля|Болтові з'єднання і анкери каркаса|Зварювання і фарбування каркаса|Підлога|"
		+ "Стіни і перегородки|Внутрішня обшивка стін|Стеля|Покрівля|Водостік|Фасад|Вікна й двері|Електрика|Електрощит|"
		+ "Резервне живлення (генератор)|Опалення, вентиляція і клімат|Водопостачання і сантехніка|Каналізація всередині|"
		+ "Зовнішні мережі|Септик|Оснащення|Кріплення загального призначення", "|");

КонецФункции

// Новый объект узла (не записан): вид «Вузол», «Чернетка», база выпуска (по умолчанию 1 «компл»).
Функция НовыйУзел(Наименование, КоличествоВыпуска, ЕдиницаВыпуска, СсылкаСС = Неопределено, Родитель = Неопределено) Экспорт

	УзелОбъект = Справочники.А_СпецификацииДомов.СоздатьЭлемент();
	УзелОбъект.Наименование = Наименование;
	УзелОбъект.ВидСпецификации = Перечисления.А_ВидыСпецификацийДомов.Узел;
	УзелОбъект.Статус = Перечисления.А_СтатусыСпецификацийДомов.Черновик;
	УзелОбъект.КоличествоВыпуска = ?(КоличествоВыпуска > 0, КоличествоВыпуска, 1);
	УзелОбъект.ЕдиницаВыпуска = ?(ЗначениеЗаполнено(ЕдиницаВыпуска), ЕдиницаВыпуска, ЕдиницаПоТексту("компл"));
	Если ЗначениеЗаполнено(СсылкаСС) Тогда
		УзелОбъект.СтруктураСебестоимости = СсылкаСС;
	КонецЕсли;
	Если ЗначениеЗаполнено(Родитель) Тогда
		УзелОбъект.Родитель = Родитель;
	КонецЕсли;
	Возврат УзелОбъект;

КонецФункции

// Строки с индексами Индексы заменяются ОДНОЙ строкой-узлом (на месте первой); количество строки — Количество
// (в единице выпуска узла). Возврат: вставленная строка.
Функция ЗаменитьСтрокиУзлом(ЦельОбъект, Индексы, Узел, Количество, ЕдиницаВыпуска) Экспорт

	Список = Новый СписокЗначений;
	Список.ЗагрузитьЗначения(Индексы);
	Список.СортироватьПоЗначению();
	Отсортированные = Список.ВыгрузитьЗначения();
	СтрокаУзла = ЦельОбъект.Состав.Вставить(Отсортированные[0]);
	СтрокаУзла.КлючСтроки = Новый УникальныйИдентификатор;
	СтрокаУзла.ВидСтроки = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	СтрокаУзла.Узел = Узел;
	СтрокаУзла.Количество = Количество;
	СтрокаУзла.ЕдиницаСтроки = ЕдиницаВыпуска;
	СтрокаУзла.Коэффициент = 1;
	СтрокаУзла.ИсточникКоэффициента = Перечисления.А_ИсточникиКоэффициентаПересчетаСС.ЕдиницыСовпадают;
	Номер = Отсортированные.ВГраница();
	Пока Номер >= 0 Цикл
		ЦельОбъект.Состав.Удалить(Отсортированные[Номер] + 1);
		Номер = Номер - 1;
	КонецЦикла;
	Возврат СтрокаУзла;

КонецФункции

// «Обʼєднати у вузол» ([M5] §3.2): выделенные строки переезжают в новый узел вместе с позициями, КлючСтроки, ценами
// и «Перевірено»; в спецификации остаётся строка узла × КоличествоВыпуска — норма и деньги не меняются. Возврат: ссылка узла.
Функция ОбъединитьВУзел(СпецОбъект, ИндексыСтрок, ИмяУзла, КоличествоВыпуска, ЕдиницаВыпуска) Экспорт

	Если ИндексыСтрок.Количество() = 0 Тогда
		ВызватьИсключение НСтр("ru = 'Не выбраны строки для узла.'; uk = 'Не вибрано рядки для вузла.'");
	КонецЕсли;
	ЗаполнитьКлючиСтрок(СпецОбъект.Состав);
	УзелОбъект = НовыйУзел(ИмяУзла, КоличествоВыпуска, ЕдиницаВыпуска, СпецОбъект.СтруктураСебестоимости, СпецОбъект.Родитель);
	ВсеПроверены = Истина;
	Список = Новый СписокЗначений;
	Список.ЗагрузитьЗначения(ИндексыСтрок);
	Список.СортироватьПоЗначению();
	Для Каждого Индекс Из Список.ВыгрузитьЗначения() Цикл
		СтрокаИсточника = СпецОбъект.Состав[Индекс];
		ЗаполнитьЗначенияСвойств(УзелОбъект.Состав.Добавить(), СтрокаИсточника, , "НомерСтроки");
		ВсеПроверены = ВсеПроверены И СтрокаИсточника.Проверено;
	КонецЦикла;
	УзелОбъект.Записать();
	СтрокаУзла = ЗаменитьСтрокиУзлом(СпецОбъект, ИндексыСтрок, УзелОбъект.Ссылка, УзелОбъект.КоличествоВыпуска, УзелОбъект.ЕдиницаВыпуска);
	СтрокаУзла.Проверено = ВсеПроверены;
	Возврат УзелОбъект.Ссылка;

КонецФункции

// «Розформувати вузол»: строки узла возвращаются в спецификацию (количества × строка / выпуск узла; позиция строки-узла —
// пустым позициям). Только если узел больше нигде не используется. Возврат: Истина.
Функция РасформироватьУзел(СпецОбъект, ИндексСтроки) Экспорт

	СтрокаУзла = СпецОбъект.Состав[ИндексСтроки];
	Если СтрокаУзла.ВидСтроки <> Перечисления.А_ВидыСтрокСпецификацииДомов.Узел Или Не ЗначениеЗаполнено(СтрокаУзла.Узел) Тогда
		ВызватьИсключение НСтр("ru = 'Текущая строка — не узел.'; uk = 'Поточний рядок — не вузол.'");
	КонецЕсли;
	УзелСсылка = СтрокаУзла.Узел;
	Для Каждого Вход Из Входимость(УзелСсылка) Цикл
		Если Вход.Уровень = 1 И Вход.Спецификация <> СпецОбъект.Ссылка Тогда
			ВызватьИсключение СтрШаблон(НСтр("ru = 'Узел используется ещё в «%1» — расформировать нельзя.'; uk = 'Вузол використовується ще в «%1» — розформувати не можна.'"),
				Вход.Спецификация);
		КонецЕсли;
	КонецЦикла;
	Ссылок = 0;
	Для Каждого Стр Из СпецОбъект.Состав Цикл
		Если Стр.Узел = УзелСсылка Тогда
			Ссылок = Ссылок + 1;
		КонецЕсли;
	КонецЦикла;
	Если Ссылок > 1 Тогда
		ВызватьИсключение НСтр("ru = 'Узел стоит в спецификации больше одного раза — расформировать нельзя.'; uk = 'Вузол стоїть у специфікації більше одного разу — розформувати не можна.'");
	КонецЕсли;
	Выпуск = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(УзелСсылка, "КоличествоВыпуска");
	Множитель = СтрокаУзла.Количество / ?(Выпуск = 0, 1, Выпуск);
	Вставлено = 0;
	Для Каждого СтрокаВУзле Из УзелСсылка.Состав Цикл
		НоваяСтрока = СпецОбъект.Состав.Вставить(ИндексСтроки + Вставлено);
		ЗаполнитьЗначенияСвойств(НоваяСтрока, СтрокаВУзле, , "НомерСтроки");
		НоваяСтрока.Количество = СтрокаВУзле.Количество * Множитель;
		Если Не ЗначениеЗаполнено(НоваяСтрока.ПозицияСС) Тогда
			НоваяСтрока.ЭтапСС = СтрокаУзла.ЭтапСС;
			НоваяСтрока.ПозицияСС = СтрокаУзла.ПозицияСС;
		КонецЕсли;
		Вставлено = Вставлено + 1;
	КонецЦикла;
	СпецОбъект.Состав.Удалить(ИндексСтроки + Вставлено);
	Возврат Истина;

КонецФункции

// «Склад з документа комплектації» (Р31): новый узел с деталями документа (на УзловВДокументе, по умолчанию 1 — выпуск
// документа состав НЕ делит), НоменклатураВузла = выпуск документа, источник карточек «З документа комплектації». Записывается.
Функция УзелИзДокументаКомплектации(ДокументКомплектации, ИмяУзла, УзловВДокументе = 1, ЕдиницаВыпуска = Неопределено,
	СсылкаСС = Неопределено) Экспорт

	Состав = СоставИзДокументаКомплектации(ДокументКомплектации, УзловВДокументе);
	УзелОбъект = НовыйУзел(ИмяУзла, 1, ЕдиницаВыпуска, СсылкаСС);
	УзелОбъект.НоменклатураВузла = Состав.НоменклатураВузла;
	Для Каждого СтрокаДетали Из Состав.Строки Цикл
		Стр = УзелОбъект.Состав.Добавить();
		Стр.КлючСтроки = Новый УникальныйИдентификатор;
		Стр.ВидСтроки = Перечисления.А_ВидыСтрокСпецификацииДомов.Деталь;
		Стр.Номенклатура = СтрокаДетали.Номенклатура;
		Стр.ОбщееНазвание = СтрокаДетали.ОбщееНазвание;
		Стр.Количество = СтрокаДетали.Количество;
		Стр.ЕдиницаСтроки = СтрокаДетали.Единица;
		Стр.Коэффициент = 1;
		Стр.ИсточникКоэффициента = Перечисления.А_ИсточникиКоэффициентаПересчетаСС.ЕдиницыСовпадают;
		Стр.ИсточникКарточки = Перечисления.А_ИсточникиКарточкиСпецификации.ИзДокументаКомплектации;
		Стр.Комментарий = Лев(СтрШаблон(НСтр("ru = 'Из %1'; uk = 'З %1'"), ДокументКомплектации), 250);
	КонецЦикла;
	УзелОбъект.Записать();
	Возврат УзелОбъект.Ссылка;

КонецФункции

// Строка с индексом (или новая строка × 1, если индекс Неопределено) становится строкой узла; позиция строки сохраняется
// (детали узла её наследуют: септик ← поз. 170), количество — тоже; единица = единица выпуска узла, k = 1, «Перевірено» снят.
Процедура ПоставитьУзелВСтроку(СпецОбъект, ИндексСтроки, Узел) Экспорт

	Если ИндексСтроки = Неопределено Тогда
		Стр = СпецОбъект.Состав.Добавить();
		Стр.КлючСтроки = Новый УникальныйИдентификатор;
		Стр.Количество = 1;
	Иначе
		Стр = СпецОбъект.Состав[ИндексСтроки];
	КонецЕсли;
	Стр.ВидСтроки = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	Стр.Узел = Узел;
	Стр.Номенклатура = Справочники.Номенклатура.ПустаяСсылка();
	Стр.ИсточникКарточки = Перечисления.А_ИсточникиКарточкиСпецификации.ПустаяСсылка();
	Стр.ЕдиницаСтроки = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Узел, "ЕдиницаВыпуска");
	Стр.Коэффициент = 1;
	Стр.ИсточникКоэффициента = Перечисления.А_ИсточникиКоэффициентаПересчетаСС.ЕдиницыСовпадают;
	Стр.Проверено = Ложь;

КонецПроцедуры

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_60_uzly.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 37 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_60_uzly.bsl _Rarzrabotki/Python/test/spec_domov/test_e2b_uzly.py
git commit -m "feat(spec_domov): узлы строителя — из документа комплектации (Р31), объединить/расформировать, узел в узле

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 35: «Нульові позиції з фактом» и печать для экономиста (Р22)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_70_nulevye.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2b_nulevye.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2b_nulevye.py`:

```python
# -*- coding: utf-8 -*-
"""Р22 на 15 м: 22 нулевые позиции, у 18 есть факт; после строк-алиасов 131, 52, 109, 70 ([Z3] §2.6) закрыто алиасами 4
(169 → 131, 173 → 52, 183 → 109, 187 → 70); печать для экономиста содержит 18 строк."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect, q, tx, vidy_subkonto, zapros  # noqa: E402
from fixtures import new_spec, row_index_by_ss, srv, ss, write  # noqa: E402


def rows(t):
    return [t.Получить(i) for i in range(t.Количество())]


def main():
    B = connect()
    s = srv(B)
    s8 = ss(B, "000000008")
    fact = q(B, zapros("nulevye_pozicii"), СС=s8, ПодразделениеОбъекта=s8.ПодразделениеОбъекта, ВидыСубконто=vidy_subkonto(B))
    first_card = {}
    for r in fact:
        if r["Номенклатура"] is not None:
            first_card.setdefault(int(r["НомерСтроки"]), r["Номенклатура"])
    with tx(B):
        h = new_spec(B, "Дом", "еталон 15 м", ss_ref=s8)
        s.ЗаполнитьИзСС(h)
        write(h)
        t = rows(s.НулевыеПозицииСФактом(h.Ссылка))
        assert len(t) == 22 and sum(1 for r in t if r.ЕстьФакт) == 18, (len(t), sum(1 for r in t if r.ЕстьФакт))
        closed_before = {int(r.НомерСтроки) for r in t if r.ЗакрытаАлиасом}
        for zero_n, alias_n in ((169, 131), (173, 52), (183, 109), (187, 70)):
            h.Состав.Получить(row_index_by_ss(B, h, "000000008", alias_n)).Номенклатура = first_card[zero_n]
        write(h)
        t = rows(s.НулевыеПозицииСФактом(h.Ссылка))
        closed = {int(r.НомерСтроки) for r in t if r.ЗакрытаАлиасом}
        assert closed - closed_before == {169, 173, 183, 187}, (sorted(closed_before), sorted(closed))
        doc = s.НулевыеПозицииВТабличныйДокумент(h.Ссылка)
        assert doc.ВысотаТаблицы == 3 + 18, doc.ВысотаТаблицы
    print("OK test_e2b_nulevye: 22 / 18 с фактом / 4 закрыты алиасами")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2b_nulevye.py`
Expected: FAIL — `AttributeError: …НулевыеПозицииСФактом`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_70_nulevye.bsl`:

```bsl
#Область НулевыеПозиции

// «Нульові позиції з фактом» (Р22): нулевые позиции СС, по ОН которых есть списания на дома типа (запрос nulevye_pozicii);
// «закрыта алиасом» — карточка факта стоит в листе спецификации с другой позицией ([Z3] §2.6). Спецификация — ссылка или объект.
// Возврат: ТЗ(НомерСтроки, Этап, НоменклатураСС, ОбщееНазвание, Карточки, ДомовСФактом, Количество, Сумма, ЕстьФакт,
// ЗакрытаАлиасом, ПозицияАлиаса) — строка на нулевую позицию.
Функция НулевыеПозицииСФактом(Спецификация) Экспорт

	СсылкаСС = Спецификация.СтруктураСебестоимости;
	Запрос = Новый Запрос(ТекстЗапросаНулевыхПозиций());
	Запрос.УстановитьПараметр("СС", СсылкаСС);
	Запрос.УстановитьПараметр("ПодразделениеОбъекта", ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "ПодразделениеОбъекта"));
	Запрос.УстановитьПараметр("ВидыСубконто", ВидыСубконтоНоменклатураСклад());
	Факт = Запрос.Выполнить().Выгрузить();
	ПозицииКарточек = Новый Соответствие;
	Для Каждого Лист Из РазвернутьДоКарточек(Спецификация, Истина).Листья Цикл
		Если ЗначениеЗаполнено(Лист.Номенклатура) И ПозицииКарточек.Получить(Лист.Номенклатура) = Неопределено Тогда
			ПозицииКарточек.Вставить(Лист.Номенклатура, Лист.ПозицияСС);
		КонецЕсли;
	КонецЦикла;
	Результат = Новый ТаблицаЗначений;
	Результат.Колонки.Добавить("НомерСтроки", Новый ОписаниеТипов("Число"));
	Результат.Колонки.Добавить("Этап");
	Результат.Колонки.Добавить("НоменклатураСС", Новый ОписаниеТипов("Строка"));
	Результат.Колонки.Добавить("ОбщееНазвание");
	Результат.Колонки.Добавить("Карточки", Новый ОписаниеТипов("Строка"));
	Результат.Колонки.Добавить("ДомовСФактом", Новый ОписаниеТипов("Число"));
	Результат.Колонки.Добавить("Количество", Новый ОписаниеТипов("Число"));
	Результат.Колонки.Добавить("Сумма", Новый ОписаниеТипов("Число"));
	Результат.Колонки.Добавить("ЕстьФакт", Новый ОписаниеТипов("Булево"));
	Результат.Колонки.Добавить("ЗакрытаАлиасом", Новый ОписаниеТипов("Булево"));
	Результат.Колонки.Добавить("ПозицияАлиаса", Новый ОписаниеТипов("Строка"));
	ПоНомеру = Новый Соответствие;
	Для Каждого СтрокаФакта Из Факт Цикл
		СтрокаРезультата = ПоНомеру.Получить(СтрокаФакта.НомерСтроки);
		Если СтрокаРезультата = Неопределено Тогда
			СтрокаРезультата = Результат.Добавить();
			СтрокаРезультата.НомерСтроки = СтрокаФакта.НомерСтроки;
			СтрокаРезультата.Этап = СтрокаФакта.Этап;
			СтрокаРезультата.НоменклатураСС = СтрокаФакта.НоменклатураСС;
			СтрокаРезультата.ОбщееНазвание = СтрокаФакта.ОбщееНазвание;
			ПоНомеру.Вставить(СтрокаФакта.НомерСтроки, СтрокаРезультата);
		КонецЕсли;
		Если Не ЗначениеЗаполнено(СтрокаФакта.Номенклатура) Тогда
			Продолжить;
		КонецЕсли;
		СтрокаРезультата.ЕстьФакт = Истина;
		СтрокаРезультата.Карточки = ?(ПустаяСтрока(СтрокаРезультата.Карточки), "", СтрокаРезультата.Карточки + "; ")
			+ Строка(СтрокаФакта.Номенклатура);
		СтрокаРезультата.ДомовСФактом = Макс(СтрокаРезультата.ДомовСФактом, СтрокаФакта.ДомовСФактом);
		СтрокаРезультата.Количество = СтрокаРезультата.Количество + СтрокаФакта.Количество;
		СтрокаРезультата.Сумма = СтрокаРезультата.Сумма + СтрокаФакта.Сумма;
		ПозицияАлиаса = ПозицииКарточек.Получить(СтрокаФакта.Номенклатура);
		Если ПозицияАлиаса <> Неопределено Тогда
			СтрокаРезультата.ЗакрытаАлиасом = Истина;
			СтрокаРезультата.ПозицияАлиаса = ПозицияАлиаса;
		КонецЕсли;
	КонецЦикла;
	Возврат Результат;

КонецФункции

// Печать «Нульові позиції з фактом» для экономиста (Р22, §5 шаг 0): только позиции с фактом; лечится в СС — поставить количество.
Функция НулевыеПозицииВТабличныйДокумент(Спецификация) Экспорт

	Таблица = НулевыеПозицииСФактом(Спецификация);
	ТабДок = Новый ТабличныйДокумент;
	ТабДок.Область(1, 1).Текст = СтрШаблон(НСтр("ru = 'Нулевые позиции СС с фактом списаний на дома типа: %1'; uk = 'Нульові позиції СС з фактом списань на будинки типу: %1'"),
		Спецификация.СтруктураСебестоимости);
	Заголовки = СтрРазделить(НСтр("ru = '№ СС|Позиция СС|Общее название|Карточки факта|Домов|Количество|Сумма, без НДС|Закрыта алиасом (позиция)'; uk = '№ СС|Позиція СС|Загальна назва|Картки факту|Будинків|Кількість|Сума, без ПДВ|Закрита аліасом (позиція)'"), "|");
	Для Номер = 0 По Заголовки.ВГраница() Цикл
		ТабДок.Область(3, Номер + 1).Текст = Заголовки[Номер];
	КонецЦикла;
	НомерСтрокиДокумента = 4;
	Для Каждого Стр Из Таблица Цикл
		Если Не Стр.ЕстьФакт Тогда
			Продолжить;
		КонецЕсли;
		ТабДок.Область(НомерСтрокиДокумента, 1).Текст = Строка(Стр.НомерСтроки);
		ТабДок.Область(НомерСтрокиДокумента, 2).Текст = Стр.НоменклатураСС;
		ТабДок.Область(НомерСтрокиДокумента, 3).Текст = Строка(Стр.ОбщееНазвание);
		ТабДок.Область(НомерСтрокиДокумента, 4).Текст = Стр.Карточки;
		ТабДок.Область(НомерСтрокиДокумента, 5).Текст = Строка(Стр.ДомовСФактом);
		ТабДок.Область(НомерСтрокиДокумента, 6).Текст = Формат(Стр.Количество, "ЧДЦ=3; ЧРГ=' '");
		ТабДок.Область(НомерСтрокиДокумента, 7).Текст = Формат(Стр.Сумма, "ЧДЦ=2; ЧРГ=' '");
		ТабДок.Область(НомерСтрокиДокумента, 8).Текст = ?(Стр.ЗакрытаАлиасом, Стр.ПозицияАлиаса, "");
		НомерСтрокиДокумента = НомерСтрокиДокумента + 1;
	КонецЦикла;
	Возврат ТабДок;

КонецФункции

Функция ТекстЗапросаНулевыхПозиций()

	Возврат
	//@ЗАПРОС nulevye_pozicii
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_70_nulevye.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 37 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_70_nulevye.bsl _Rarzrabotki/Python/test/spec_domov/test_e2b_nulevye.py
git commit -m "feat(spec_domov): «Нульові позиції з фактом» с алиасами и печать для экономиста (Р22)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 36: «Розрахувати ціни норми» (И4) и «Порівняти з СС» (Р24, Р26, Р27)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_80_ceny.bsl`, `_Rarzrabotki/Python/test/spec_domov/bsl/srv_90_sravnenie.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2b_ceny.py`

- [ ] **Step 1: Написать падающий тест (приёмка 30 м §10.1)**

`_Rarzrabotki/Python/test/spec_domov/test_e2b_ceny.py`:

```python
# -*- coding: utf-8 -*-
"""Приёмка этапа 2 на 30 м (§10.1): 170 позиций; мембрана 1,093 рул; плінтус 30,6667 шт; труба 110/3000 1,1667 шт; кутник 10 м;
ІБ00-000174 → септик 1/2/1/1, Σ 12 000,00; Σ норм = 1 009 864,74 ± 1,00 (И4); «Порівняти з СС» — без сирот и пустых позиций."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect, tx  # noqa: E402
from fixtures import card, kn, new_spec, row_index_by_ss, srv, ss, write  # noqa: E402


def set_card(B, h, n, code, expect_k):
    r = h.Состав.Получить(row_index_by_ss(B, h, "000000009", n))
    r.Номенклатура = card(B, code)
    hint = srv(B).ПодсказкаКоэффициента(r.Номенклатура, r.ЕдиницаСтроки)
    assert abs(hint.Коэффициент - expect_k) < 1e-9, (n, code, hint.Коэффициент)
    r.Коэффициент = hint.Коэффициент
    r.ИсточникКоэффициента = B.Перечисления.А_ИсточникиКоэффициентаПересчетаСС.Вручную


def main():
    B = connect()
    s = srv(B)
    with tx(B):
        s9 = ss(B, "000000009")
        h = new_spec(B, "Дом", "еталон 30 м", ss_ref=s9)
        p = s.ЗаполнитьИзСС(h)
        assert p.Строк == 170
        set_card(B, h, 42, "00-00006711", 75)
        set_card(B, h, 53, "00-00009980", 3)
        set_card(B, h, 117, "00-00013067", 3)
        set_card(B, h, 118, "00-00013067", 3)
        i13 = row_index_by_ss(B, h, "000000009", 13)
        h.Состав.Получить(i13).Номенклатура = card(B, "00-00022213")
        assert s.ВзятьИзОригиналаExcel(h, i13)
        node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
        s.ПоставитьУзелВСтроку(h, row_index_by_ss(B, h, "000000009", 170), node)
        write(h)
        r = s.РазвернутьДоКарточек(h.Ссылка)
        by_card = {}
        for i in range(r.Листья.Количество()):
            x = r.Листья.Получить(i)
            if x.Номенклатура is not None and not x.Номенклатура.Пустая():
                by_card[x.Номенклатура.Код.strip()] = by_card.get(x.Номенклатура.Код.strip(), 0) + x.Количество
        assert round(by_card["00-00006711"], 4) == 1.093, by_card["00-00006711"]
        assert round(by_card["00-00009980"], 4) == 30.6667, by_card["00-00009980"]
        assert round(by_card["00-00013067"], 4) == 1.1667, by_card["00-00013067"]
        assert round(by_card["00-00022213"], 4) == 10, by_card["00-00022213"]
        sep = [r.Листья.Получить(i) for i in range(r.Листья.Количество()) if r.Листья.Получить(i).Уровень == 1]
        assert sorted(round(x.Количество, 6) for x in sep) == [1, 1, 1, 2], "септик 1/2/1/1, не 2/4/2/2"
        pr = s.РассчитатьЦеныНормы(h)
        write(h)
        assert round(pr.СуммаСС, 2) == 1009864.74 and abs(pr.И4) <= 1.00 and pr.БезЛистьев.Количество() == 0, (pr.СуммаСС, pr.И4)
        r2 = s.РазвернутьДоКарточек(h.Ссылка, True)
        total = round(sum(r2.Листья.Получить(i).СуммаНормы for i in range(r2.Листья.Количество())), 2)
        assert abs(total - 1009864.74) <= 1.00, total
        sep_sum = round(sum(r2.Листья.Получить(i).СуммаНормы for i in range(r2.Листья.Количество())
                            if r2.Листья.Получить(i).Уровень == 1), 2)
        assert abs(sep_sum - 12000.00) <= 0.01, sep_sum
        cmp = s.СравнитьСоСтруктуройСебестоимости(h.Ссылка)
        assert cmp.Сироты.Количество() == 0 and cmp.ПозицийБезСтрок == 0 and abs(cmp.Разница) <= 1.00, \
            (cmp.Сироты.Количество(), cmp.ПозицийБезСтрок, cmp.Разница)
        st = {int(cmp.Позиции.Получить(i).НомерСтроки): cmp.Позиции.Получить(i).Статус for i in range(cmp.Позиции.Количество())}
        # Глобальные функции языка (НСтр, Формат, ТипЗнч) через COM недоступны (AttributeError, память com_date_shift_3_hours):
        # статус «Совпадает» сравнивается с обоими языковыми вариантами НСтр модуля.
        match = {"Збіг", "Совпадает"}
        assert st[13] in match and st[42] in match and st[170] in match, (st[13], st[42], st[170])
    print(f"OK test_e2b_ceny: И4 = {round(pr.И4, 2)}, септик {sep_sum}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2b_ceny.py`
Expected: FAIL — `AttributeError: …УзелИзДокументаКомплектации` (или `…РассчитатьЦеныНормы`).

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_80_ceny.bsl`:

```bsl
#Область ЦеныНормы

// «Розрахувати ціни норми» ([M5] §4.5, §6 п. 6): сумма каждой ненулевой позиции СС раскладывается по листьям разузлованного
// эталона (вместе с «Не постачається» — И4 считается с ними, Р24) пропорционально «количество × цена последней закупки (Т2) × 1,2»;
// если закупки нет хоть у одного листа позиции — поровну. Остаток копеек — на лист с наибольшим количеством (И4 до копейки).
// Цены пишутся в строки объекта (корень) и в строки узлов (запись узлов); в узел, используемый спецификациями другой СС,
// цена пишется только в пустые строки (протокол). Нулевые позиции не трогаются (цена вручную, Р22).
// Возврат: Структура(СуммаСС, СуммаНормы, И4, Позиций, БезЛистьев — Массив, ПропущеноУзлов — Массив).
Функция РассчитатьЦеныНормы(ЭталонОбъект, НаДату = Неопределено) Экспорт

	СсылкаСС = ЭталонОбъект.СтруктураСебестоимости;
	ЗаполнитьКлючиСтрок(ЭталонОбъект.Состав);
	Позиции = ПозицииСтруктурыСебестоимости(СсылкаСС);
	Листья = РазвернутьДоКарточек(ЭталонОбъект, Истина).Листья;
	Листья.Колонки.Добавить("КлючПозиции", Новый ОписаниеТипов("Строка"));
	Карточки = Новый Массив;
	Для Каждого Лист Из Листья Цикл
		Лист.КлючПозиции = А_СтруктураСебестоимостиСервер.КлючПозиции(Лист.ЭтапСС, Лист.ПозицияСС);
		Если ЗначениеЗаполнено(Лист.Номенклатура) И Карточки.Найти(Лист.Номенклатура) = Неопределено Тогда
			Карточки.Добавить(Лист.Номенклатура);
		КонецЕсли;
	КонецЦикла;
	Цены = ЦеныПоследнейЗакупки(Карточки, ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "ПодразделениеЗакупки"), НаДату);
	Протокол = Новый Структура("СуммаСС, СуммаНормы, И4, Позиций, БезЛистьев, ПропущеноУзлов", 0, 0, 0, 0, Новый Массив, Новый Массив);
	ЦеныСтрок = Новый Соответствие;
	Для Каждого ЭлементПозиции Из Позиции Цикл
		Позиция = ЭлементПозиции.Значение;
		Если Позиция.Количество = 0 Тогда
			Продолжить;
		КонецЕсли;
		Протокол.СуммаСС = Протокол.СуммаСС + Позиция.Сумма;
		ЛистьяПозиции = Листья.НайтиСтроки(Новый Структура("КлючПозиции", ЭлементПозиции.Ключ));
		Если ЛистьяПозиции.Количество() = 0 Тогда
			Протокол.БезЛистьев.Добавить(СтрШаблон("%1 %2", Позиция.НомерСтроки, Позиция.НоменклатураСС));
			Продолжить;
		КонецЕсли;
		Протокол.Позиций = Протокол.Позиций + 1;
		РазложитьСуммуПозиции(Позиция.Сумма, ЛистьяПозиции, Цены);
		Для Каждого Лист Из ЛистьяПозиции Цикл
			Протокол.СуммаНормы = Протокол.СуммаНормы + Лист.СуммаНормы;
			КлючЦены = Строка(Лист.Спецификация.УникальныйИдентификатор()) + "|" + Строка(Лист.КлючСтроки);
			Если ЦеныСтрок.Получить(КлючЦены) = Неопределено Тогда
				ЦеныСтрок.Вставить(КлючЦены, Новый Структура("Спецификация, КлючСтроки, ЦенаНормы",
					Лист.Спецификация, Лист.КлючСтроки, Лист.ЦенаНормы));
			КонецЕсли;
		КонецЦикла;
	КонецЦикла;
	ЗаписатьЦеныСтрок(ЭталонОбъект, ЦеныСтрок, СсылкаСС, Протокол);
	ЭталонОбъект.ДатаЦенНормы = НачалоДня(ТекущаяДатаСеанса());
	Протокол.И4 = Протокол.СуммаНормы - Протокол.СуммаСС;
	Возврат Протокол;

КонецФункции

Процедура РазложитьСуммуПозиции(СуммаПозиции, ЛистьяПозиции, Цены)

	ВсеСЦеной = Истина;
	Для Каждого Лист Из ЛистьяПозиции Цикл
		Если Лист.Количество = 0 Или Цены.Получить(Лист.Номенклатура) = Неопределено Тогда
			ВсеСЦеной = Ложь;
		КонецЕсли;
	КонецЦикла;
	Веса = Новый Массив;
	ВсегоВес = 0;
	Для Каждого Лист Из ЛистьяПозиции Цикл
		Вес = ?(ВсеСЦеной, Лист.Количество * Цены.Получить(Лист.Номенклатура).ЦенаБезНДС * 1.2, 1);
		Веса.Добавить(Вес);
		ВсегоВес = ВсегоВес + Вес;
	КонецЦикла;
	Если ВсегоВес = 0 Тогда
		Для Номер = 0 По Веса.ВГраница() Цикл
			Веса[Номер] = 1;
		КонецЦикла;
		ВсегоВес = Веса.Количество();
	КонецЕсли;
	Распределено = 0;
	Лучший = Неопределено;
	Для Номер = 0 По ЛистьяПозиции.ВГраница() Цикл
		Лист = ЛистьяПозиции[Номер];
		Если Лист.Количество = 0 Тогда
			Лист.ЦенаНормы = 0;
			Лист.СуммаНормы = 0;
			Продолжить;
		КонецЕсли;
		Лист.ЦенаНормы = Окр(СуммаПозиции * Веса[Номер] / ВсегоВес / Лист.Количество, 6);
		Лист.СуммаНормы = Окр(Лист.Количество * Лист.ЦенаНормы, 2);
		Распределено = Распределено + Лист.СуммаНормы;
		Если Лучший = Неопределено Или Лист.Количество > Лучший.Количество Тогда
			Лучший = Лист;
		КонецЕсли;
	КонецЦикла;
	Остаток = СуммаПозиции - Распределено;
	Если Остаток <> 0 И Лучший <> Неопределено Тогда
		Лучший.ЦенаНормы = Окр((Лучший.СуммаНормы + Остаток) / Лучший.Количество, 6);
		Лучший.СуммаНормы = Окр(Лучший.Количество * Лучший.ЦенаНормы, 2);
	КонецЕсли;

КонецПроцедуры

Процедура ЗаписатьЦеныСтрок(ЭталонОбъект, ЦеныСтрок, СсылкаСС, Протокол)

	ПоСпецификациям = Новый Соответствие;
	Для Каждого ЭлементЦены Из ЦеныСтрок Цикл
		Цена = ЭлементЦены.Значение;
		Ключ = ?(Цена.Спецификация.Пустая() Или Цена.Спецификация = ЭталонОбъект.Ссылка, "КОРЕНЬ", Цена.Спецификация);
		Список = ПоСпецификациям.Получить(Ключ);
		Если Список = Неопределено Тогда
			Список = Новый Массив;
			ПоСпецификациям.Вставить(Ключ, Список);
		КонецЕсли;
		Список.Добавить(Цена);
	КонецЦикла;
	Для Каждого ЭлементСпец Из ПоСпецификациям Цикл
		Если ТипЗнч(ЭлементСпец.Ключ) = Тип("Строка") Тогда
			ПрименитьЦены(ЭталонОбъект.Состав, ЭлементСпец.Значение, Ложь);
			Продолжить;
		КонецЕсли;
		ДругаяСС = Ложь;
		Для Каждого Вход Из Входимость(ЭлементСпец.Ключ) Цикл
			Если ЗначениеЗаполнено(Вход.СтруктураСебестоимости) И Вход.СтруктураСебестоимости <> СсылкаСС Тогда
				ДругаяСС = Истина;
			КонецЕсли;
		КонецЦикла;
		УзелОбъект = ЭлементСпец.Ключ.ПолучитьОбъект();
		Если ПрименитьЦены(УзелОбъект.Состав, ЭлементСпец.Значение, ДругаяСС) Тогда
			УзелОбъект.Записать();
		КонецЕсли;
		Если ДругаяСС Тогда
			Протокол.ПропущеноУзлов.Добавить(Строка(ЭлементСпец.Ключ));
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

Функция ПрименитьЦены(ТабличнаяЧасть, Цены, ТолькоПустые)

	Изменено = Ложь;
	Для Каждого Цена Из Цены Цикл
		Стр = ТабличнаяЧасть.Найти(Цена.КлючСтроки, "КлючСтроки");
		Если Стр = Неопределено Или (ТолькоПустые И Стр.ЦенаНормы <> 0) Тогда
			Продолжить;
		КонецЕсли;
		Если Стр.ЦенаНормы <> Цена.ЦенаНормы Тогда
			Стр.ЦенаНормы = Цена.ЦенаНормы;
			Изменено = Истина;
		КонецЕсли;
	КонецЦикла;
	Возврат Изменено;

КонецФункции

#КонецОбласти
```

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_90_sravnenie.bsl`:

```bsl
#Область СравнениеСоСС

// «Порівняти з СС» ([M5] §3.9; Р24, Р26, Р27): по каждой позиции СС — количество строк с явной позицией (в той из двух единиц
// позиции, что совпадает с единицей строки; строка-узел сравнивается с «шт»/«компл» позиции), сумма нормы листьев (вместе
// с «Не постачається»), статус; сироты — строки с позицией, которой нет в СС; «Додано в будинку» — листья без позиции.
// Допуск количества 0,00001 (Р26). Спецификация — ссылка или объект.
// Возврат: Структура(Позиции, Сироты, ДобавленоВДоме — ТЗ; СуммаСС, СуммаНормы, СуммаНеПоставляется, Разница, ПозицийБезСтрок).
Функция СравнитьСоСтруктуройСебестоимости(Спецификация) Экспорт

	Позиции = ПозицииСтруктурыСебестоимости(Спецификация.СтруктураСебестоимости);
	Разузлование = РазвернутьДоКарточек(Спецификация, Истина);
	ТаблицаПозиций = Новый ТаблицаЗначений;
	Для Каждого ИмяКолонки Из СтрРазделить("НомерСтроки,КоличествоСС,КоличествоОригинал,СуммаСС,КоличествоСпец,КоличествоСпецОригинал,СуммаНормы", ",") Цикл
		ТаблицаПозиций.Колонки.Добавить(ИмяКолонки, Новый ОписаниеТипов("Число"));
	КонецЦикла;
	Для Каждого ИмяКолонки Из СтрРазделить("ПозицияСС,ЕдиницаСС,ЕдиницаОригинал,Статус", ",") Цикл
		ТаблицаПозиций.Колонки.Добавить(ИмяКолонки, Новый ОписаниеТипов("Строка"));
	КонецЦикла;
	Для Каждого ИмяКолонки Из СтрРазделить("ДругаяЕдиница,НеПоставляется,ЕстьСтроки", ",") Цикл
		ТаблицаПозиций.Колонки.Добавить(ИмяКолонки, Новый ОписаниеТипов("Булево"));
	КонецЦикла;
	ТаблицаПозиций.Колонки.Добавить("ЭтапСС");
	ПоКлючу = Новый Соответствие;
	Для Каждого ЭлементПозиции Из Позиции Цикл
		Позиция = ЭлементПозиции.Значение;
		СтрокаПозиции = ТаблицаПозиций.Добавить();
		СтрокаПозиции.НомерСтроки = Позиция.НомерСтроки;
		СтрокаПозиции.ЭтапСС = Позиция.Этап;
		СтрокаПозиции.ПозицияСС = Позиция.НоменклатураСС;
		СтрокаПозиции.КоличествоСС = Позиция.Количество;
		СтрокаПозиции.ЕдиницаСС = Позиция.ЕдиницаСС;
		СтрокаПозиции.КоличествоОригинал = Позиция.КоличествоОригинал;
		СтрокаПозиции.ЕдиницаОригинал = Позиция.ЕдиницаОригинал;
		СтрокаПозиции.СуммаСС = Позиция.Сумма;
		ПоКлючу.Вставить(ЭлементПозиции.Ключ, СтрокаПозиции);
	КонецЦикла;
	Сироты = Разузлование.ЯвныеПозиции.СкопироватьКолонки();
	Для Каждого Явная Из Разузлование.ЯвныеПозиции Цикл
		СтрокаПозиции = ПоКлючу.Получить(А_СтруктураСебестоимостиСервер.КлючПозиции(Явная.ЭтапСС, Явная.ПозицияСС));
		Если СтрокаПозиции = Неопределено Тогда
			ЗаполнитьЗначенияСвойств(Сироты.Добавить(), Явная);
			Продолжить;
		КонецЕсли;
		СтрокаПозиции.ЕстьСтроки = Истина;
		Если ЕдиницаСтрокиКакУПозиции(Явная, СтрокаПозиции.ЕдиницаСС) Тогда
			СтрокаПозиции.КоличествоСпец = СтрокаПозиции.КоличествоСпец + Явная.КоличествоЕдСтроки;
		ИначеЕсли Не ПустаяСтрока(СтрокаПозиции.ЕдиницаОригинал) И ЕдиницаСтрокиКакУПозиции(Явная, СтрокаПозиции.ЕдиницаОригинал) Тогда
			СтрокаПозиции.КоличествоСпецОригинал = СтрокаПозиции.КоличествоСпецОригинал + Явная.КоличествоЕдСтроки;
		Иначе
			СтрокаПозиции.ДругаяЕдиница = Истина;
		КонецЕсли;
	КонецЦикла;
	ДобавленоВДоме = Разузлование.Листья.СкопироватьКолонки();
	Итог = Новый Структура("Позиции, Сироты, ДобавленоВДоме, СуммаСС, СуммаНормы, СуммаНеПоставляется, Разница, ПозицийБезСтрок",
		ТаблицаПозиций, Сироты, ДобавленоВДоме, 0, 0, 0, 0, 0);
	Для Каждого Лист Из Разузлование.Листья Цикл
		Итог.СуммаНормы = Итог.СуммаНормы + Лист.СуммаНормы;
		Если Лист.НеПоставляется Тогда
			Итог.СуммаНеПоставляется = Итог.СуммаНеПоставляется + Лист.СуммаНормы;
		КонецЕсли;
		Если Не ЗначениеЗаполнено(Лист.ПозицияСС) Тогда
			ЗаполнитьЗначенияСвойств(ДобавленоВДоме.Добавить(), Лист);
			Продолжить;
		КонецЕсли;
		СтрокаПозиции = ПоКлючу.Получить(А_СтруктураСебестоимостиСервер.КлючПозиции(Лист.ЭтапСС, Лист.ПозицияСС));
		Если СтрокаПозиции <> Неопределено Тогда
			СтрокаПозиции.СуммаНормы = СтрокаПозиции.СуммаНормы + Лист.СуммаНормы;
			СтрокаПозиции.НеПоставляется = СтрокаПозиции.НеПоставляется Или Лист.НеПоставляется;
		КонецЕсли;
	КонецЦикла;
	Совпадает = НСтр("ru = 'Совпадает'; uk = 'Збіг'");
	Для Каждого СтрокаПозиции Из ТаблицаПозиций Цикл
		Итог.СуммаСС = Итог.СуммаСС + СтрокаПозиции.СуммаСС;
		Если СтрокаПозиции.КоличествоСС = 0 И Не СтрокаПозиции.ЕстьСтроки Тогда
			СтрокаПозиции.Статус = НСтр("ru = 'Нулевая'; uk = 'Нульова'");
		ИначеЕсли Не СтрокаПозиции.ЕстьСтроки Тогда
			СтрокаПозиции.Статус = НСтр("ru = 'Нет строк (не норма)'; uk = 'Немає рядків (не норма)'");
			Итог.ПозицийБезСтрок = Итог.ПозицийБезСтрок + 1;
		ИначеЕсли СтрокаПозиции.ДругаяЕдиница Тогда
			СтрокаПозиции.Статус = НСтр("ru = 'Другая единица — сверьте'; uk = 'Інша одиниця — звірте'");
		ИначеЕсли СтрокаПозиции.КоличествоСпец = 0 Тогда
			СтрокаПозиции.Статус = ?(Макс(СтрокаПозиции.КоличествоСпецОригинал - СтрокаПозиции.КоличествоОригинал,
				СтрокаПозиции.КоличествоОригинал - СтрокаПозиции.КоличествоСпецОригинал) <= 0.00001, Совпадает,
				НСтр("ru = 'Количество другое (по оригиналу)'; uk = 'Кількість інша (за оригіналом)'"));
		Иначе
			СтрокаПозиции.Статус = ?(Макс(СтрокаПозиции.КоличествоСпец - СтрокаПозиции.КоличествоСС,
				СтрокаПозиции.КоличествоСС - СтрокаПозиции.КоличествоСпец) <= 0.00001, Совпадает,
				НСтр("ru = 'Количество другое'; uk = 'Кількість інша'"));
		КонецЕсли;
		Если СтрокаПозиции.НеПоставляется Тогда
			СтрокаПозиции.Статус = СтрокаПозиции.Статус + НСтр("ru = ' (не поставляется)'; uk = ' (не постачається)'");
		КонецЕсли;
	КонецЦикла;
	Итог.Разница = Итог.СуммаНормы - Итог.СуммаСС;
	Возврат Итог;

КонецФункции

Функция ЕдиницаСтрокиКакУПозиции(Явная, ТекстЕдиницыПозиции)

	Если ЕдиницыСовпадают(Явная.ЕдиницаСтроки, ТекстЕдиницыПозиции) Тогда
		Возврат Истина;
	КонецЕсли;
	КлючПозиции = КлючЕдиницыСпецификации(ТекстЕдиницыПозиции);
	Возврат Явная.ЭтоУзел И (КлючПозиции = "шт" Или КлючПозиции = "компл");

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_80_ceny.bsl _Rarzrabotki/Python/test/spec_domov/bsl/srv_90_sravnenie.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 37 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_80_ceny.bsl _Rarzrabotki/Python/test/spec_domov/bsl/srv_90_sravnenie.bsl _Rarzrabotki/Python/test/spec_domov/test_e2b_ceny.py
git commit -m "feat(spec_domov): ціни норми с контролем И4 до копейки и «Порівняти з СС» (Р24, Р26, Р27)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 37: Загрузка 2B и прогон тестов партии

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/load_lists/e2b.txt` (одна строка `CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl`)

- [ ] **Step 1–2:** тесты `test_e2b_uzly.py`, `test_e2b_nulevye.py`, `test_e2b_ceny.py` падают (`AttributeError`, проверено в задачах 34–36).

- [ ] **Step 3: Сборка и загрузка**

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/build_bsl.py --partiya 2b
python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2b.txt --tests
python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2b.txt
powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e2b.txt -Tag e2b
```
Expected: `частей 13`; `ошибок: 0`; `OK load e2b`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2b_uzly.py && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2b_nulevye.py && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2b_ceny.py && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2a_zapolnit.py`
Expected: `OK test_e2b_uzly …`, `OK test_e2b_nulevye: 22 / 18 с фактом / 4 закрыты алиасами`, `OK test_e2b_ceny: И4 = …, септик 12000.0`, регресс 2A — `OK`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/load_lists/e2b.txt _Rarzrabotki/Python/test/spec_domov/sync_manifest.json \
  _Rarzrabotki/Python/test/spec_domov/out/e2b_load_*.txt \
  "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"
git commit -m "feat(BASEBuh): партия 2B загружена — узлы, нулевые позиции, ціни норми, порівняння з СС; приёмка 30 м пройдена

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 38: «Перепривʼязати позиції СС» ([M5] §3.10, Т10)


**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_A0_perepriv.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2c_perepriv.py`

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2c_perepriv.py`:

```python
# -*- coding: utf-8 -*-
"""Перепривязка ([M5] §3.10, Т10): мягкий ключ, расстояние правки; копия СС 009 в транзакции с «Комплект септика» вместо
«Копмлпект септика» → ровно 1 сирота, найдено зелёным предложением; «Застосувати» снимает сироту."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, tx  # noqa: E402
from fixtures import kn, new_spec, row_index_by_ss, srv, ss, write  # noqa: E402


def main():
    B = connect()
    s = srv(B)
    assert s.МягкийКлючПозиции("кутик 32*32 *3") == s.МягкийКлючПозиции("Кутик 32х32х3") == s.МягкийКлючПозиции("кутик 32x32x3")
    assert s.РасстояниеПравки("Копмлпект септика", "Комплект септика") == 2
    assert s.РасстояниеПравки("абв", "абв") == 0 and s.РасстояниеПравки("а", "абвгд") == 3
    with tx(B):
        s9 = ss(B, "000000009")
        h = new_spec(B, "Дом", "еталон 30 м", ss_ref=s9)
        s.ЗаполнитьИзСС(h)
        node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
        s.ПоставитьУзелВСтроку(h, row_index_by_ss(B, h, "000000009", 170), node)
        h_ref = write(h)
        assert s.ПредложенияПерепривязки(s9, arr(B, [h_ref])).Количество() == 0, "до перезагрузки сирот нет"
        sso = s9.ПолучитьОбъект()
        r170 = sso.Комплектующие.Получить(169)
        assert r170.НоменклатураСС == "Копмлпект септика"
        r170.НоменклатураСС = "Комплект септика"
        sso.ОбменДанными.Загрузка = True
        sso.ДополнительныеСвойства.Вставить("ОтключитьМеханизмРегистрацииОбъектов", True)
        sso.Записать()
        pr = s.ПредложенияПерепривязки(s9, arr(B, [h_ref]))
        assert pr.Количество() == 1, pr.Количество()
        p = pr.Получить(0)
        assert p.Цвет == "Зелений" and p.НоваяПозиция == "Комплект септика" and p.Применить, (p.Цвет, p.НоваяПозиция, p.Причина)
        assert s.ПерепривязатьПозицииСС(pr) == 1
        assert s.ПредложенияПерепривязки(s9, arr(B, [h_ref])).Количество() == 0
    print("OK test_e2c_perepriv: 1 сирота, зелёное «єдина сирота й єдина нова позиція етапу» (Т10)")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2c_perepriv.py`
Expected: FAIL — `AttributeError: …МягкийКлючПозиции`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_A0_perepriv.bsl`:

```bsl
#Область Перепривязка

// Мягкий ключ позиции ([M5] §3.10): нижний регистр, без пробелов и знаков; «*», латинская «x» и кириллическая «х» → «х»;
// латинские буквы-двойники → кириллица.
Функция МягкийКлючПозиции(Знач ПозицияСС) Экспорт

	Текст = НРег(СокрЛП(ПозицияСС));
	Замены = Новый Соответствие;
	Замены.Вставить("*", "х");
	Замены.Вставить("x", "х");
	Замены.Вставить("a", "а");
	Замены.Вставить("c", "с");
	Замены.Вставить("e", "е");
	Замены.Вставить("i", "і");
	Замены.Вставить("o", "о");
	Замены.Вставить("p", "р");
	Замены.Вставить("k", "к");
	Замены.Вставить("m", "м");
	Замены.Вставить("t", "т");
	Замены.Вставить("y", "у");
	Замены.Вставить("h", "н");
	Замены.Вставить("b", "в");
	Результат = "";
	Для Номер = 1 По СтрДлина(Текст) Цикл
		Символ = Сред(Текст, Номер, 1);
		Замена = Замены.Получить(Символ);
		Если Замена <> Неопределено Тогда
			Символ = Замена;
		КонецЕсли;
		Если ЭтоЦифра(Символ) Или ЭтоБуква(Символ) Тогда
			Результат = Результат + Символ;
		КонецЕсли;
	КонецЦикла;
	Возврат Результат;

КонецФункции

// Расстояние Дамерау — Левенштейна (OSA) без учёта регистра; при разнице длин больше Предел — Предел + 1 без расчёта.
Функция РасстояниеПравки(Знач СтрокаА, Знач СтрокаБ, Предел = 2) Экспорт

	ТекстА = НРег(СокрЛП(СтрокаА));
	ТекстБ = НРег(СокрЛП(СтрокаБ));
	ДлинаА = СтрДлина(ТекстА);
	ДлинаБ = СтрДлина(ТекстБ);
	Если Макс(ДлинаА - ДлинаБ, ДлинаБ - ДлинаА) > Предел Тогда
		Возврат Предел + 1;
	КонецЕсли;
	Д = Новый Массив(ДлинаА + 1, ДлинаБ + 1);
	Для Сч1 = 0 По ДлинаА Цикл
		Д[Сч1][0] = Сч1;
	КонецЦикла;
	Для Сч2 = 0 По ДлинаБ Цикл
		Д[0][Сч2] = Сч2;
	КонецЦикла;
	Для Сч1 = 1 По ДлинаА Цикл
		Для Сч2 = 1 По ДлинаБ Цикл
			Стоимость = ?(Сред(ТекстА, Сч1, 1) = Сред(ТекстБ, Сч2, 1), 0, 1);
			Значение = Мин(Д[Сч1 - 1][Сч2] + 1, Д[Сч1][Сч2 - 1] + 1, Д[Сч1 - 1][Сч2 - 1] + Стоимость);
			Если Сч1 > 1 И Сч2 > 1 И Сред(ТекстА, Сч1, 1) = Сред(ТекстБ, Сч2 - 1, 1) И Сред(ТекстА, Сч1 - 1, 1) = Сред(ТекстБ, Сч2, 1) Тогда
				Значение = Мин(Значение, Д[Сч1 - 2][Сч2 - 2] + 1);
			КонецЕсли;
			Д[Сч1][Сч2] = Значение;
		КонецЦикла;
	КонецЦикла;
	Возврат Д[ДлинаА][ДлинаБ];

КонецФункции

// «Перепривʼязати позиції СС…» ([M5] §3.10, Т10): строки с позицией, которой нет в СС (сироты), по всем спецификациям СС и узлам
// их деревьев (ТолькоСпецификации — для теста: только эти и их узлы). Предложения среди позиций без строк: зелёные — тот же
// текст в другом этапе; совпал мягкий ключ; единственная сирота и единственная свободная позиция этапа с тем же количеством
// (Т10); жёлтое — до 2 правок и то же количество. Узел другой СС — не применяется сам.
// Возврат: ТЗ(Применить, Спецификация, КлючСтроки, СтарыйЭтап, СтараяПозиция, НовыйЭтап, НоваяПозиция, НомерСтрокиСС, Цвет,
// Причина, ДругаяСС) — колонки формы перепривязки.
Функция ПредложенияПерепривязки(СсылкаСС, ТолькоСпецификации = Неопределено) Экспорт

	Позиции = ПозицииСтруктурыСебестоимости(СсылкаСС);
	Спецификации = ДобавитьУзлыДеревьев(?(ТолькоСпецификации = Неопределено, СпецификацииСС(СсылкаСС), ТолькоСпецификации));
	Результат = НоваяТаблицаПредложений();
	Занятые = Новый Соответствие;
	Сироты = Новый Массив;
	Для Каждого Спецификация Из Спецификации Цикл
		СсылкаССУзла = Спецификация.СтруктураСебестоимости;
		ДругаяСС = ЗначениеЗаполнено(СсылкаССУзла) И СсылкаССУзла <> СсылкаСС;
		Для Каждого Стр Из Спецификация.Состав Цикл
			Если Не ЗначениеЗаполнено(Стр.ПозицияСС) Тогда
				Продолжить;
			КонецЕсли;
			Ключ = А_СтруктураСебестоимостиСервер.КлючПозиции(Стр.ЭтапСС, Стр.ПозицияСС);
			Если Позиции.Получить(Ключ) <> Неопределено Тогда
				Занятые.Вставить(Ключ, Истина);
			Иначе
				Сироты.Добавить(Новый Структура("Спецификация, КлючСтроки, ЭтапСС, ПозицияСС, Количество, ЕдиницаСтроки, ЭтоУзел, ДругаяСС, Ключ",
					Спецификация, Стр.КлючСтроки, Стр.ЭтапСС, Стр.ПозицияСС, Стр.Количество, Стр.ЕдиницаСтроки,
					Стр.ВидСтроки = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел, ДругаяСС, Ключ));
			КонецЕсли;
		КонецЦикла;
	КонецЦикла;
	// Свободные — ненулевые позиции без строк: нулевые не норма (Р22), строку с количеством на них не переносим.
	Свободные = Новый Массив;
	Для Каждого ЭлементПозиции Из Позиции Цикл
		Если Занятые.Получить(ЭлементПозиции.Ключ) = Неопределено И ЭлементПозиции.Значение.Количество <> 0 Тогда
			Свободные.Добавить(ЭлементПозиции.Значение);
		КонецЕсли;
	КонецЦикла;
	Для Каждого Сирота Из Сироты Цикл
		НоваяСтрока = Результат.Добавить();
		НоваяСтрока.Спецификация = Сирота.Спецификация;
		НоваяСтрока.КлючСтроки = Сирота.КлючСтроки;
		НоваяСтрока.СтарыйЭтап = Сирота.ЭтапСС;
		НоваяСтрока.СтараяПозиция = Сирота.ПозицияСС;
		НоваяСтрока.ДругаяСС = Сирота.ДругаяСС;
		Предложение = ПредложениеДляСироты(Сирота, Сироты, Свободные);
		Если Предложение <> Неопределено Тогда
			НоваяСтрока.НовыйЭтап = Предложение.Позиция.Этап;
			НоваяСтрока.НоваяПозиция = Предложение.Позиция.НоменклатураСС;
			НоваяСтрока.НомерСтрокиСС = Предложение.Позиция.НомерСтроки;
			НоваяСтрока.Цвет = Предложение.Цвет;
			НоваяСтрока.Причина = Предложение.Причина;
		КонецЕсли;
		НоваяСтрока.Применить = (НоваяСтрока.Цвет = "Зелений") И Не Сирота.ДругаяСС;
		Если Сирота.ДругаяСС Тогда
			НоваяСтрока.Причина = СокрЛП(НоваяСтрока.Причина + " " + НСтр("ru = '(узел другой СС — подтвердите)'; uk = '(вузол іншої СС — підтвердіть)'"));
		КонецЕсли;
	КонецЦикла;
	Возврат Результат;

КонецФункции

// Применяет отмеченные предложения: ЭтапСС и ПозицияСС строки по КлючСтроки; каждая спецификация записывается один раз.
// Возврат: число изменённых строк.
Функция ПерепривязатьПозицииСС(Предложения) Экспорт

	Изменено = 0;
	Объекты = Новый Соответствие;
	Для Каждого Предложение Из Предложения Цикл
		Если Не Предложение.Применить Или ПустаяСтрока(Предложение.НоваяПозиция) Тогда
			Продолжить;
		КонецЕсли;
		СпецОбъект = Объекты.Получить(Предложение.Спецификация);
		Если СпецОбъект = Неопределено Тогда
			СпецОбъект = Предложение.Спецификация.ПолучитьОбъект();
			Объекты.Вставить(Предложение.Спецификация, СпецОбъект);
		КонецЕсли;
		Стр = СпецОбъект.Состав.Найти(Предложение.КлючСтроки, "КлючСтроки");
		Если Стр = Неопределено Тогда
			Продолжить;
		КонецЕсли;
		Стр.ЭтапСС = Предложение.НовыйЭтап;
		Стр.ПозицияСС = Предложение.НоваяПозиция;
		Изменено = Изменено + 1;
	КонецЦикла;
	Для Каждого ЭлементОбъекта Из Объекты Цикл
		ЭлементОбъекта.Значение.Записать();
	КонецЦикла;
	Возврат Изменено;

КонецФункции

Функция ПредложениеДляСироты(Сирота, Сироты, Свободные)

	ТекстСироты = А_СтруктураСебестоимостиСервер.КлючОригинала(Сирота.ПозицияСС);
	МягкийСироты = МягкийКлючПозиции(Сирота.ПозицияСС);
	ТотЖеТекст = Новый Массив;
	ТотЖеМягкий = Новый Массив;
	ВЭтапе = Новый Массив;
	Похожие = Новый Массив;
	Для Каждого Позиция Из Свободные Цикл
		Если А_СтруктураСебестоимостиСервер.КлючОригинала(Позиция.НоменклатураСС) = ТекстСироты Тогда
			ТотЖеТекст.Добавить(Позиция);
		КонецЕсли;
		Если Позиция.Этап = Сирота.ЭтапСС Тогда
			ВЭтапе.Добавить(Позиция);
			Если МягкийКлючПозиции(Позиция.НоменклатураСС) = МягкийСироты Тогда
				ТотЖеМягкий.Добавить(Позиция);
			ИначеЕсли РасстояниеПравки(Позиция.НоменклатураСС, Сирота.ПозицияСС) <= 2 И КоличествоСовпадает(Сирота, Позиция) Тогда
				Похожие.Добавить(Позиция);
			КонецЕсли;
		КонецЕсли;
	КонецЦикла;
	Если ТотЖеТекст.Количество() = 1 Тогда
		Возврат Новый Структура("Позиция, Цвет, Причина", ТотЖеТекст[0], "Зелений",
			НСтр("ru = 'тот же текст в другом этапе'; uk = 'той самий текст в іншому етапі'"));
	КонецЕсли;
	Если ТотЖеМягкий.Количество() = 1 Тогда
		Возврат Новый Структура("Позиция, Цвет, Причина", ТотЖеМягкий[0], "Зелений",
			НСтр("ru = 'совпал мягкий ключ (регистр, пробелы, знаки, х/x/*)'; uk = 'збіг мʼякого ключа (регістр, пробіли, знаки, х/x/*)'"));
	КонецЕсли;
	КлючиЭтапа = Новый Соответствие;
	Для Каждого Другая Из Сироты Цикл
		Если Другая.ЭтапСС = Сирота.ЭтапСС Тогда
			КлючиЭтапа.Вставить(Другая.Ключ, Истина);
		КонецЕсли;
	КонецЦикла;
	Если КлючиЭтапа.Количество() = 1 И ВЭтапе.Количество() = 1 И КоличествоСовпадает(Сирота, ВЭтапе[0]) Тогда
		Возврат Новый Структура("Позиция, Цвет, Причина", ВЭтапе[0], "Зелений",
			НСтр("ru = 'единственная сирота и единственная новая позиция этапа, количество совпало (Т10)'; uk = 'єдина сирота й єдина нова позиція етапу, кількість збігається (Т10)'"));
	КонецЕсли;
	Если Похожие.Количество() = 1 Тогда
		Возврат Новый Структура("Позиция, Цвет, Причина", Похожие[0], "Жовтий",
			НСтр("ru = 'похожий текст (до 2 правок), то же количество'; uk = 'схожий текст (до 2 правок), та сама кількість'"));
	КонецЕсли;
	Возврат Неопределено;

КонецФункции

Функция КоличествоСовпадает(Сирота, Позиция)

	КлючЕдиницыПозиции = КлючЕдиницыСпецификации(Позиция.ЕдиницаСС);
	Если ЕдиницыСовпадают(Сирота.ЕдиницаСтроки, Позиция.ЕдиницаСС)
		Или (Сирота.ЭтоУзел И (КлючЕдиницыПозиции = "шт" Или КлючЕдиницыПозиции = "компл")) Тогда
		Возврат Макс(Сирота.Количество - Позиция.Количество, Позиция.Количество - Сирота.Количество) <= 0.00001;
	КонецЕсли;
	Если Не ПустаяСтрока(Позиция.ЕдиницаОригинал) И ЕдиницыСовпадают(Сирота.ЕдиницаСтроки, Позиция.ЕдиницаОригинал) Тогда
		Возврат Макс(Сирота.Количество - Позиция.КоличествоОригинал, Позиция.КоличествоОригинал - Сирота.Количество) <= 0.00001;
	КонецЕсли;
	Возврат Ложь;

КонецФункции

Функция СпецификацииСС(СсылкаСС)

	Запрос = Новый Запрос(ТекстЗапросаСпецификацийСС());
	Запрос.УстановитьПараметр("СсылкаСС", СсылкаСС);
	Возврат Запрос.Выполнить().Выгрузить().ВыгрузитьКолонку("Ссылка");

КонецФункции

Функция ДобавитьУзлыДеревьев(Спецификации)

	Все = Новый Массив;
	Для Каждого Спецификация Из Спецификации Цикл
		Дерево = ЗагрузитьДеревоСпецификаций(Спецификация);
		Для Каждого ЭлементКэша Из Дерево.Кэш Цикл
			Если ТипЗнч(ЭлементКэша.Ключ) <> Тип("Строка") И Все.Найти(ЭлементКэша.Ключ) = Неопределено Тогда
				Все.Добавить(ЭлементКэша.Ключ);
			КонецЕсли;
		КонецЦикла;
	КонецЦикла;
	Возврат Все;

КонецФункции

Функция НоваяТаблицаПредложений()

	Таблица = Новый ТаблицаЗначений;
	Колонки = Таблица.Колонки;
	Колонки.Добавить("Применить", Новый ОписаниеТипов("Булево"));
	Колонки.Добавить("Спецификация", Новый ОписаниеТипов("СправочникСсылка.А_СпецификацииДомов"));
	Колонки.Добавить("КлючСтроки", Новый ОписаниеТипов("УникальныйИдентификатор"));
	Колонки.Добавить("СтарыйЭтап", Новый ОписаниеТипов("СправочникСсылка.ЭтапыРабот"));
	Колонки.Добавить("СтараяПозиция", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(200)));
	Колонки.Добавить("НовыйЭтап", Новый ОписаниеТипов("СправочникСсылка.ЭтапыРабот"));
	Колонки.Добавить("НоваяПозиция", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(200)));
	Колонки.Добавить("НомерСтрокиСС", Новый ОписаниеТипов("Число"));
	Колонки.Добавить("Цвет", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(10)));
	Колонки.Добавить("Причина", Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(250)));
	Колонки.Добавить("ДругаяСС", Новый ОписаниеТипов("Булево"));
	Возврат Таблица;

КонецФункции

Функция ТекстЗапросаСпецификацийСС()

	Возврат
	//@ЗАПРОС specifikacii_ss
	;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_A0_perepriv.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 40 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_A0_perepriv.bsl _Rarzrabotki/Python/test/spec_domov/test_e2c_perepriv.py
git commit -m "feat(spec_domov): «Перепривʼязати позиції СС» — мягкий ключ, Дамерау–Левенштейн, зелёное правило Т10

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 39: «Скопіювати картки й дерево вузлів з іншої специфікації» (Р32, Т11) и приёмка 15 м

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/srv_B0_kopiya.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2c_kopiya.py`

- [ ] **Step 1: Написать падающий тест (приёмка 15 м §10.1)**

`_Rarzrabotki/Python/test/spec_domov/test_e2c_kopiya.py`:

```python
# -*- coding: utf-8 -*-
"""Приёмка этапа 2 на 15 м (§10.1, Р32, Т11): 15 м копией из 30 м — каждая скопированная карточка, расходящаяся с фактом 15 м,
жёлтая (независимая сверка с фактом); поз. 44, 45 (пач/пак) и 27 (балон) — k = 1 без ручного ввода; Септик переиспользован в
поз. 168 (похожий текст), «Водостік 15 м» скопирован; норма по явным позициям не изменилась; Σ = 704 612,01, И4 = 0,00;
строки «Не постачається» не мешают «Діє»."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, q, same, tx  # noqa: E402
from fixtures import (add_row, card, enum, kn, new_spec, row_index_by_ss, srv, ss, ss_first_row_with_card, ss_row,  # noqa: E402
                      unit, write)
from test_e2b_uzly import explicit_norm  # noqa: E402


def names(B):
    """Представления значений перечислений в языке сеанса — сравнение не зависит от языка базы."""
    return (B.String(enum(B, "А_ИсточникиКарточкиСпецификации", "КопияИзДругойСпецификации")),
            B.String(enum(B, "А_ИсточникиКарточкиСпецификации", "Вручную")),
            B.String(enum(B, "А_ИсточникиКоэффициентаПересчетаСС", "ЕдиницыСовпадают")),
            B.String(enum(B, "А_ВидыСтрокСпецификацииДомов", "Узел")))


def build_source(B, s, s9):
    src = new_spec(B, "Дом", "джерело 30 м", ss_ref=s9)
    s.ЗаполнитьИзСС(src)
    for n30, code in ((27, "00000000449"), (45, "00-00006990"), (46, "00-00021890")):
        r = src.Состав.Получить(row_index_by_ss(B, src, "000000009", n30))
        r.Номенклатура = card(B, code)
        s.УстановитьКоэффициентПоЕдиницам(r)
    node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
    s.ПоставитьУзелВСтроку(src, row_index_by_ss(B, src, "000000009", 170), node)
    idx = [row_index_by_ss(B, src, "000000009", n) for n in range(67, 75)]
    vod_src = s.ОбъединитьВУзел(src, arr(B, idx), "Водостік 30 м", 1, unit(B, "2938"))
    return write(src), node, vod_src


def t_copy(B):
    s = srv(B)
    YELLOW, MANUAL, MATCH, NODE = names(B)
    with tx(B):
        s9, s8 = ss(B, "000000009"), ss(B, "000000008")
        src, septik, vod_src = build_source(B, s, s9)
        dst = new_spec(B, "Дом", "еталон 15 м", ss_ref=s8)
        s.ЗаполнитьИзСС(dst)
        before = explicit_norm(B, dst)
        p = s.СкопироватьИзДругойСпецификации(dst, src)
        assert explicit_norm(B, dst) == before, "копия изменила норму по позициям"
        assert p.УзловПереиспользовано >= 1 and p.УзловСкопировано >= 1, (p.УзловПереиспользовано, p.УзловСкопировано)
        r168 = dst.Состав.Получить(row_index_by_ss(B, dst, "000000008", 168))
        assert B.String(r168.ВидСтроки) == NODE and same(B, r168.Узел, septik), "Т11: септик в поз. 168"
        vod = q(B, """ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
ГДЕ
	Спец.Наименование = &Имя
	И Спец.Основание = &Основание""", Имя="Водостік 15 м", Основание=vod_src)
        assert len(vod) == 1, "копия «Водостік 15 м» с основанием «Водостік 30 м» теста (боевые узлы бухгалтера не мешают)"
        for n in (27, 44, 45):
            r = dst.Состав.Получить(row_index_by_ss(B, dst, "000000008", n))
            assert not r.Номенклатура.Пустая() and r.Коэффициент == 1 and B.String(r.ИсточникКоэффициента) == MATCH, \
                (n, B.String(r.Номенклатура), r.Коэффициент, B.String(r.ИсточникКоэффициента))
        fakt = s.ФактСписанийДома(s.ДомаТипа(s8))
        by_on = {}
        for i in range(fakt.Количество()):
            f = fakt.Получить(i)
            by_on.setdefault(B.String(f.ОбщееНазвание), set()).add(f.Номенклатура.Код.strip())
        checked = 0
        for i in range(dst.Состав.Количество()):
            r = dst.Состав.Получить(i)
            src_name = B.String(r.ИсточникКарточки)
            if src_name not in (YELLOW, MANUAL):
                continue
            cards = by_on.get(B.String(r.ОбщееНазвание))
            must_be_yellow = bool(cards) and r.Номенклатура.Код.strip() not in cards
            assert (src_name == YELLOW) == must_be_yellow, (r.ПозицияСС, B.String(r.Номенклатура), src_name)
            checked += 1
        assert checked > 0
        write(dst)
        pc = s.РассчитатьЦеныНормы(dst)
        assert round(pc.СуммаСС, 2) == 704612.01 and round(pc.И4, 2) == 0.00, (pc.СуммаСС, pc.И4)
    print(f"OK копия: проверено скопированных карточек {checked}, совпало {p.Совпало}, жёлтых {p.Отличий}; Σ 704 612,01, И4 0,00")


def t_diye_ne_postach(B):
    with tx(B):
        s8 = ss(B, "000000008")
        h = new_spec(B, "Дом", "діє з не постачається", ss_ref=s8, st="Действует")
        r1 = ss_first_row_with_card(B, "000000008")
        r2 = ss_row(B, "000000008", 2 if int(r1["НомерСтроки"]) == 1 else 1)
        add_row(B, h, card_ref=r1["Номенклатура"], qty=r1["Количество"], stage=r1["Этап"], pos=r1["НоменклатураСС"], checked=True)
        add_row(B, h, qty=r2["Количество"] or 1, stage=r2["Этап"], pos=r2["НоменклатураСС"], ne_post=True,
                comment="не постачається — тест приймання", ksrc="НеНайден")
        write(h)
    print("OK «Не постачається» не мешает «Діє»")



if __name__ == "__main__":
    B = connect()
    t_copy(B)
    t_diye_ne_postach(B)
    print("OK test_e2c_kopiya")
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2c_kopiya.py`
Expected: FAIL — `AttributeError: …СкопироватьИзДругойСпецификации`.

- [ ] **Step 3: Реализовать**

`_Rarzrabotki/Python/test/spec_domov/bsl/srv_B0_kopiya.bsl`:

```bsl
#Область КопияИзДругойСпецификации

// «Скопіювати картки й дерево вузлів з іншої специфікації» (Р32, Т11). СпецОбъект — объект-получатель (эталон другого типа
// после «Заповнити з СС»), Источник — ссылка. 1) Карточки: по ключу позиции в пустые строки «Деталь»; карточка, которой нет
// среди списанных на дома типа-получателя при наличии факта по её ОН, — жёлтая «Скопійовано з іншої специфікації», иначе
// «Вручну»; k копируется, если ключи единиц строк совпали, иначе — по единицам; «Не постачається» с причиной — в пустые строки.
// 2) Узлы верхнего уровня источника: узел-позиция (септик) ставится в строку с той же позицией или — в том же этапе —
// с похожим текстом (до 2 правок, Т11); узел из нескольких позиций переиспользуется, если все его позиции нашлись с теми же
// количествами, иначе копируется с суффиксом типа («Водостік 30 м» → «Водостік 15 м»), вложенные узлы — так же.
// Норма по явным позициям не меняется. Возврат: Структура(Совпало, Отличий, НеЗакрыто, НеПоставляетсяСкопировано,
// УзловПереиспользовано, УзловСкопировано, Отличия — Массив строк).
Функция СкопироватьИзДругойСпецификации(СпецОбъект, Источник, Знач СуффиксТипа = "") Экспорт

	Протокол = Новый Структура("Совпало, Отличий, НеЗакрыто, НеПоставляетсяСкопировано, УзловПереиспользовано, УзловСкопировано, Отличия",
		0, 0, 0, 0, 0, 0, Новый Массив);
	СсылкаСС = СпецОбъект.СтруктураСебестоимости;
	Если ПустаяСтрока(СуффиксТипа) Тогда
		СуффиксТипа = СуффиксТипаСС(СсылкаСС);
	КонецЕсли;
	СуффиксИсточника = СуффиксТипаСС(ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Источник, "СтруктураСебестоимости"));
	ЗаполнитьКлючиСтрок(СпецОбъект.Состав);
	СкопироватьКарточки(СпецОбъект, Источник, СсылкаСС, Протокол);
	ПрименитьУзлыИсточника(СпецОбъект, Источник, СуффиксИсточника, СуффиксТипа, СсылкаСС, Протокол);
	Возврат Протокол;

КонецФункции

Процедура СкопироватьКарточки(СпецОбъект, Источник, СсылкаСС, Протокол)

	ВидыСтрок = Перечисления.А_ВидыСтрокСпецификацииДомов;
	ЛистьяПоКлючу = Новый Соответствие;
	Для Каждого Лист Из РазвернутьДоКарточек(Источник, Истина).Листья Цикл
		Если Не ЗначениеЗаполнено(Лист.ПозицияСС) Тогда
			Продолжить;
		КонецЕсли;
		Ключ = А_СтруктураСебестоимостиСервер.КлючПозиции(Лист.ЭтапСС, Лист.ПозицияСС);
		Список = ЛистьяПоКлючу.Получить(Ключ);
		Если Список = Неопределено Тогда
			Список = Новый Массив;
			ЛистьяПоКлючу.Вставить(Ключ, Список);
		КонецЕсли;
		Список.Добавить(Лист);
	КонецЦикла;
	ФактПоОН = ФактКарточекПоОН(СсылкаСС);
	Для Каждого Стр Из СпецОбъект.Состав Цикл
		Если Стр.ВидСтроки <> ВидыСтрок.Деталь Или Не ЗначениеЗаполнено(Стр.ПозицияСС) Тогда
			Продолжить;
		КонецЕсли;
		Найденные = ЛистьяПоКлючу.Получить(А_СтруктураСебестоимостиСервер.КлючПозиции(Стр.ЭтапСС, Стр.ПозицияСС));
		Если Найденные = Неопределено Или Найденные.Количество() <> 1 Тогда
			Если Не ЗначениеЗаполнено(Стр.Номенклатура) Тогда
				Протокол.НеЗакрыто = Протокол.НеЗакрыто + 1;
			КонецЕсли;
			Продолжить;
		КонецЕсли;
		Лист = Найденные[0];
		Если Лист.НеПоставляется Тогда
			Если Не ЗначениеЗаполнено(Стр.Номенклатура) И Не Стр.НеПоставляется Тогда
				Стр.НеПоставляется = Истина;
				Стр.Комментарий = Лист.Комментарий;
				Протокол.НеПоставляетсяСкопировано = Протокол.НеПоставляетсяСкопировано + 1;
			КонецЕсли;
			Продолжить;
		КонецЕсли;
		Если Лист.ВидСтроки <> ВидыСтрок.Деталь Или Не ЗначениеЗаполнено(Лист.Номенклатура) Тогда
			Если Не ЗначениеЗаполнено(Стр.Номенклатура) Тогда
				Протокол.НеЗакрыто = Протокол.НеЗакрыто + 1;
			КонецЕсли;
			Продолжить;
		КонецЕсли;
		Если ЗначениеЗаполнено(Стр.Номенклатура) Тогда
			Если Стр.Номенклатура = Лист.Номенклатура Тогда
				Протокол.Совпало = Протокол.Совпало + 1;
			Иначе
				Протокол.Отличия.Добавить(СтрШаблон(НСтр("ru = '%1: в строке %2, в источнике %3 — не заменено'; uk = '%1: у рядку %2, у джерелі %3 — не замінено'"),
					Стр.ПозицияСС, Стр.Номенклатура, Лист.Номенклатура));
			КонецЕсли;
			Продолжить;
		КонецЕсли;
		Стр.Номенклатура = Лист.Номенклатура;
		Если Не ЗначениеЗаполнено(Стр.ОбщееНазвание) Тогда
			Стр.ОбщееНазвание = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Лист.Номенклатура, "ОбщееНазваниеНоменклатуры");
		КонецЕсли;
		Стр.Проверено = Ложь;
		КарточкиФакта = ФактПоОН.Получить(Стр.ОбщееНазвание);
		Если КарточкиФакта <> Неопределено И КарточкиФакта.Найти(Лист.Номенклатура) = Неопределено Тогда
			Стр.ИсточникКарточки = Перечисления.А_ИсточникиКарточкиСпецификации.КопияИзДругойСпецификации;
			Протокол.Отличий = Протокол.Отличий + 1;
			Протокол.Отличия.Добавить(СтрШаблон(НСтр("ru = '%1: скопирована %2, на домах типа списывали %3'; uk = '%1: скопійовано %2, на будинках типу списували %3'"),
				Стр.ПозицияСС, Лист.Номенклатура, СтрСоединить(ПредставленияЗначений(КарточкиФакта), ", ")));
		Иначе
			Стр.ИсточникКарточки = Перечисления.А_ИсточникиКарточкиСпецификации.Вручную;
			Протокол.Совпало = Протокол.Совпало + 1;
		КонецЕсли;
		Если КлючЕдиницыСпецификации(НаименованиеЕдиницы(Стр.ЕдиницаСтроки)) = КлючЕдиницыСпецификации(НаименованиеЕдиницы(Лист.ЕдиницаСтроки)) Тогда
			Стр.Коэффициент = Лист.Коэффициент;
			Стр.ИсточникКоэффициента = Лист.ИсточникКоэффициента;
		Иначе
			УстановитьКоэффициентПоЕдиницам(Стр);
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

Процедура ПрименитьУзлыИсточника(ЦельОбъект, ИсточникСсылка, СуффиксИсточника, СуффиксТипа, СсылкаСС, Протокол)

	ВидУзел = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	Для Каждого СтрокаИсточника Из ИсточникСсылка.Состав Цикл
		Если СтрокаИсточника.ВидСтроки <> ВидУзел Или СтрокаИсточника.ПокупаетсяГотовым Или Не ЗначениеЗаполнено(СтрокаИсточника.Узел) Тогда
			Продолжить;
		КонецЕсли;
		Если ЗначениеЗаполнено(СтрокаИсточника.ПозицияСС) Тогда
			ПеренестиУзелПозицию(ЦельОбъект, СтрокаИсточника, Протокол);
		Иначе
			ПеренестиУзелСоставной(ЦельОбъект, СтрокаИсточника, СуффиксИсточника, СуффиксТипа, СсылкаСС, Протокол);
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

Процедура ПеренестиУзелПозицию(ЦельОбъект, СтрокаИсточника, Протокол)

	ВидУзел = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	КлючИсточника = А_СтруктураСебестоимостиСервер.КлючПозиции(СтрокаИсточника.ЭтапСС, СтрокаИсточника.ПозицияСС);
	Индекс = Неопределено;
	Похожие = Новый Массив;
	Для Каждого Стр Из ЦельОбъект.Состав Цикл
		Если Стр.ВидСтроки = ВидУзел Или Не ЗначениеЗаполнено(Стр.ПозицияСС) Тогда
			Продолжить;
		КонецЕсли;
		Если А_СтруктураСебестоимостиСервер.КлючПозиции(Стр.ЭтапСС, Стр.ПозицияСС) = КлючИсточника Тогда
			Индекс = ЦельОбъект.Состав.Индекс(Стр);
			Прервать;
		КонецЕсли;
		Если Стр.ЭтапСС = СтрокаИсточника.ЭтапСС И РасстояниеПравки(Стр.ПозицияСС, СтрокаИсточника.ПозицияСС) <= 2 Тогда
			Похожие.Добавить(ЦельОбъект.Состав.Индекс(Стр));
		КонецЕсли;
	КонецЦикла;
	Если Индекс = Неопределено И Похожие.Количество() = 1 Тогда
		Индекс = Похожие[0];
	КонецЕсли;
	Если Индекс = Неопределено Тогда
		Протокол.Отличия.Добавить(СтрШаблон(НСтр("ru = 'Узел «%1»: позиции «%2» в СС получателя нет.'; uk = 'Вузол «%1»: позиції «%2» у СС одержувача немає.'"),
			СтрокаИсточника.Узел, СтрокаИсточника.ПозицияСС));
		Возврат;
	КонецЕсли;
	ПоставитьУзелВСтроку(ЦельОбъект, Индекс, СтрокаИсточника.Узел);
	Протокол.УзловПереиспользовано = Протокол.УзловПереиспользовано + 1;

КонецПроцедуры

Процедура ПеренестиУзелСоставной(ЦельОбъект, СтрокаИсточника, СуффиксИсточника, СуффиксТипа, СсылкаСС, Протокол)

	ВидУзел = Перечисления.А_ВидыСтрокСпецификацииДомов.Узел;
	УзелИсточника = СтрокаИсточника.Узел;
	РеквизитыУзла = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(УзелИсточника,
		"Наименование, КоличествоВыпуска, ЕдиницаВыпуска, НоменклатураВузла, Родитель");
	// Количества позиций узла в масштабе контейнера: разузлование узла даёт q / выпуск, строка источника — × Количество.
	НужноПоКлючу = Новый Соответствие;
	Для Каждого Явная Из РазвернутьДоКарточек(УзелИсточника, Истина).ЯвныеПозиции Цикл
		Ключ = А_СтруктураСебестоимостиСервер.КлючПозиции(Явная.ЭтапСС, Явная.ПозицияСС);
		Было = НужноПоКлючу.Получить(Ключ);
		НужноПоКлючу.Вставить(Ключ, ?(Было = Неопределено, 0, Было) + Явная.КоличествоЕдСтроки * СтрокаИсточника.Количество);
	КонецЦикла;
	Индексы = Новый Массив;
	ВсеСовпали = Истина;
	Для Каждого Стр Из ЦельОбъект.Состав Цикл
		Если Стр.ВидСтроки = ВидУзел Или Не ЗначениеЗаполнено(Стр.ПозицияСС) Тогда
			Продолжить;
		КонецЕсли;
		Нужно = НужноПоКлючу.Получить(А_СтруктураСебестоимостиСервер.КлючПозиции(Стр.ЭтапСС, Стр.ПозицияСС));
		Если Нужно = Неопределено Тогда
			Продолжить;
		КонецЕсли;
		Индексы.Добавить(ЦельОбъект.Состав.Индекс(Стр));
		Если Макс(Стр.Количество - Нужно, Нужно - Стр.Количество) > 0.000001 Тогда
			ВсеСовпали = Ложь;
		КонецЕсли;
	КонецЦикла;
	Если Индексы.Количество() = 0 Тогда
		Протокол.Отличия.Добавить(СтрШаблон(НСтр("ru = 'Узел «%1»: его позиций в СС получателя нет.'; uk = 'Вузол «%1»: його позицій у СС одержувача немає.'"),
			РеквизитыУзла.Наименование));
		Возврат;
	КонецЕсли;
	Если ВсеСовпали И Индексы.Количество() = НужноПоКлючу.Количество() Тогда
		ЗаменитьСтрокиУзлом(ЦельОбъект, Индексы, УзелИсточника, СтрокаИсточника.Количество, РеквизитыУзла.ЕдиницаВыпуска);
		Протокол.УзловПереиспользовано = Протокол.УзловПереиспользовано + 1;
		Возврат;
	КонецЕсли;
	КопияОбъект = НовыйУзел(ИмяКопииУзла(РеквизитыУзла.Наименование, СуффиксИсточника, СуффиксТипа), РеквизитыУзла.КоличествоВыпуска,
		РеквизитыУзла.ЕдиницаВыпуска, СсылкаСС, РеквизитыУзла.Родитель);
	КопияОбъект.НоменклатураВузла = РеквизитыУзла.НоменклатураВузла;
	КопияОбъект.Основание = УзелИсточника;
	Масштаб = РеквизитыУзла.КоличествоВыпуска / ?(СтрокаИсточника.Количество = 0, 1, СтрокаИсточника.Количество);
	Для Каждого Индекс Из Индексы Цикл
		НоваяСтрока = КопияОбъект.Состав.Добавить();
		ЗаполнитьЗначенияСвойств(НоваяСтрока, ЦельОбъект.Состав[Индекс], , "НомерСтроки");
		НоваяСтрока.Количество = ЦельОбъект.Состав[Индекс].Количество * Масштаб;
	КонецЦикла;
	ПрименитьУзлыИсточника(КопияОбъект, УзелИсточника, СуффиксИсточника, СуффиксТипа, СсылкаСС, Протокол);
	КопияОбъект.Записать();
	ЗаменитьСтрокиУзлом(ЦельОбъект, Индексы, КопияОбъект.Ссылка, СтрокаИсточника.Количество, РеквизитыУзла.ЕдиницаВыпуска);
	Протокол.УзловСкопировано = Протокол.УзловСкопировано + 1;

КонецПроцедуры

Функция ФактКарточекПоОН(СсылкаСС)

	Результат = Новый Соответствие;
	Для Каждого СтрокаФакта Из ФактСписанийДома(ДомаТипа(СсылкаСС)) Цикл
		Список = Результат.Получить(СтрокаФакта.ОбщееНазвание);
		Если Список = Неопределено Тогда
			Список = Новый Массив;
			Результат.Вставить(СтрокаФакта.ОбщееНазвание, Список);
		КонецЕсли;
		Если Список.Найти(СтрокаФакта.Номенклатура) = Неопределено Тогда
			Список.Добавить(СтрокаФакта.Номенклатура);
		КонецЕсли;
	КонецЦикла;
	Возврат Результат;

КонецФункции

Функция ПредставленияЗначений(Значения)

	Представления = Новый Массив;
	Для Каждого Значение Из Значения Цикл
		Представления.Добавить(Строка(Значение));
	КонецЦикла;
	Возврат Представления;

КонецФункции

// «МД IRС 2026 30 м (МД IRС 2026)» → «30 м»: число перед отдельным словом «м».
Функция СуффиксТипаСС(СсылкаСС)

	Если Не ЗначениеЗаполнено(СсылкаСС) Тогда
		Возврат "";
	КонецЕсли;
	Слова = СтрРазделить(ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СсылкаСС, "Наименование"), " ", Ложь);
	Для Номер = 0 По Слова.ВГраница() - 1 Цикл
		Если Слова[Номер + 1] = "м" И СтрДлина(Слова[Номер]) <= 3 И ЭтоЦифра(Лев(Слова[Номер], 1)) Тогда
			Возврат Слова[Номер] + " м";
		КонецЕсли;
	КонецЦикла;
	Возврат "";

КонецФункции

Функция ИмяКопииУзла(Знач Имя, СуффиксИсточника, СуффиксТипа)

	Имя = СокрП(Имя);
	Если Не ПустаяСтрока(СуффиксИсточника) И СтрЗаканчиваетсяНа(Имя, СуффиксИсточника) Тогда
		Возврат Лев(Имя, СтрДлина(Имя) - СтрДлина(СуффиксИсточника)) + СуффиксТипа;
	КонецЕсли;
	Возврат Имя + " " + СуффиксТипа;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика сейчас, COM — после загрузки**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/srv_B0_kopiya.bsl`
Expected: кроме «нет BOM», «не CRLF», «неразвёрнутый //@ЗАПРОС» (исходник пишется инструментом Write) ошибок нет.
COM-тест задачи проходит в Task 40 после загрузки партии.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/srv_B0_kopiya.bsl _Rarzrabotki/Python/test/spec_domov/test_e2c_kopiya.py
git commit -m "feat(spec_domov): копия карточек и дерева узлов из другой спецификации (Р32, Т11) — жёлтые расхождения с фактом

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 40: Загрузка 2C и прогон тестов партии

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/load_lists/e2c.txt` (одна строка `CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl`)

- [ ] **Step 1–2:** тесты `test_e2c_perepriv.py`, `test_e2c_kopiya.py` падают (`AttributeError`).

- [ ] **Step 3: Сборка и загрузка**

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/build_bsl.py --partiya 2c
python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2c.txt --tests
python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2c.txt
powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e2c.txt -Tag e2c
```
Expected: `частей 15`; `ошибок: 0`; `OK load e2c`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2c_perepriv.py && python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2c_kopiya.py`
Expected: `OK test_e2c_perepriv: 1 сирота, зелёное …`, `OK копия: …; Σ 704 612,01, И4 0,00`, `OK «Не постачається» не мешает «Діє»`, `OK test_e2c_kopiya`.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/load_lists/e2c.txt _Rarzrabotki/Python/test/spec_domov/sync_manifest.json \
  _Rarzrabotki/Python/test/spec_domov/out/e2c_load_*.txt \
  "_Rarzrabotki/BASEBuh/CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"
git commit -m "feat(BASEBuh): партия 2C загружена — перепривязка и копия из другой спецификации; приёмка 15 м пройдена

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 41: Форма элемента этапа 2 — команды бухгалтера (§5, Р22, Р24, Р25, Р31, Р32, Р35)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_2.bsl`
- Test: `_Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py`

Команды и кнопки этапа 2 уже описаны в `gen_forms.py` (`CMD2_SOSTAV` — кнопки таблицы «Склад», `CMD2_FORM` — командная панель формы, «Порівняти з СС» — кнопка страницы «Порівняння»); модуль формы этапа 2 — вторая часть, которую `gen_forms.py --etap 2` дописывает после `forma_elementa_1.bsl`. Вся логика — в общем модуле: форма берёт объект (`РеквизитФормыВЗначение`), вызывает серверную функцию и возвращает объект (`ЗначениеВРеквизитФормы`). Диалоги — только немодальные (`ПоказатьВопрос`, `ПоказатьВводСтроки/Числа/Значения`, `ОписаниеОповещения`), режим совместимости 8.3.12 без `Асинх`.

- [ ] **Step 1: Написать падающий тест**

`_Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py`:

```python
# -*- coding: utf-8 -*-
"""Формы этапа 2. --files (до загрузки): у каждой команды формы — <Command> с <Action> и обработчик в модуле; у каждого
ОписаниеОповещения("X", …) — экспортная процедура X; имена процедур модуля не повторяются; id элементов уникальны;
формы подбора и перепривязки — параметры, таблицы, события. Без ключа (после загрузки 2D) — пять форм справочника в базе."""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, connect, read_text  # noqa: E402

FORMS = "Catalogs/А_СпецификацииДомов/Forms"
COMMANDS = {
    "ФормаЭлемента": ["СделатьЭталономСС", "ПоказатьВходимость", "ОбновитьДерево", "РазвернутьДоДеталей", "ЗаполнитьИзСС",
                      "ОтметитьЗеленыеПроверенными", "ПодобратьКарточку", "ВзятьИзОригиналаExcel", "ОбъединитьВУзел",
                      "РасформироватьУзел", "СоставИзДокументаКомплектации", "РассчитатьЦеныНормы", "СравнитьСоСС",
                      "НулевыеПозицииСФактом", "СкопироватьИзДругойСпецификации"],
    "ФормаСписка": ["ПроверитьЦиклы", "ПерепривязатьПозицииСС"],
    "ФормаПодбораКарточки": ["ВыбратьКарточку", "ОбновитьКандидатов"],
    "ФормаПерепривязки": ["НайтиПредложения", "ПрименитьОтмеченные"],
}
EVENTS = {
    "ФормаЭлемента": ["ПриСозданииНаСервере", "ВидСпецификацииПриИзменении", "ТолькоПроблемныеПриИзменении",
                      "СоставНоменклатураПриИзменении"],
    "ФормаПодбораКарточки": ["ПриСозданииНаСервере", "ТолькоОНПриИзменении", "СтрокаПоискаПриИзменении", "КандидатыВыбор"],
    "ФормаПерепривязки": ["ПриСозданииНаСервере", "ТипБудинкуССПриИзменении"],
}
TABLES = {"ФормаПодбораКарточки": "Кандидаты", "ФормаПерепривязки": "Предложения"}
PARAMS = {"ФормаПодбораКарточки": ["ТипБудинкуСС", "ОбщееНазвание", "ПозицияСС"], "ФормаПерепривязки": ["ТипБудинкуСС"]}


def files(form):
    x = read_text(os.path.join(BB_WT, FORMS, form, "Ext", "Form.xml"))[0]
    m = read_text(os.path.join(BB_WT, FORMS, form, "Ext", "Form", "Module.bsl"))[0]
    return x, m


def proc_names(module):
    return re.findall(r"(?m)^(?:Процедура|Функция) (\w+)\(", module)


def check_files():
    fails = []
    for form, cmds in COMMANDS.items():
        try:
            x, m = files(form)
        except FileNotFoundError as e:
            fails.append(f"{form}: нет файла {e.filename}")
            continue
        names = proc_names(m)
        dup = sorted({n for n in names if names.count(n) > 1})
        if dup:
            fails.append(f"{form}: повтор процедур {dup}")
        for c in cmds:
            if f'<Command name="{c}"' not in x or f"<Action>{c}</Action>" not in x:
                fails.append(f"{form}: нет команды {c}")
            if not re.search(r"(?m)^Процедура " + c + r"\(Команда\)", m):
                fails.append(f"{form}: нет обработчика команды {c}")
        for ev in EVENTS.get(form, []):
            if not re.search(r"(?m)^Процедура " + ev + r"\(", m):
                fails.append(f"{form}: нет обработчика события {ev}")
        for cb in sorted(set(re.findall(r'ОписаниеОповещения\("(\w+)"', m))):
            if not re.search(r"(?m)^Процедура " + cb + r"\([^)]*\) Экспорт", m):
                fails.append(f"{form}: оповещение {cb} без экспортной процедуры")
        ids = re.findall(r' id="(\d+)"', x.split("<Attributes>")[0])
        if len(ids) != len(set(ids)):
            fails.append(f"{form}: повтор id элементов")
        if form in TABLES and f'<Table name="{TABLES[form]}"' not in x:
            fails.append(f"{form}: нет таблицы {TABLES[form]}")
        for p in PARAMS.get(form, []):
            if f'<Parameter name="{p}"' not in x:
                fails.append(f"{form}: нет параметра {p}")
    return fails


def check_com():
    B = connect()
    cat = B.Metadata.Catalogs.Find("А_СпецификацииДомов")
    have = [cat.Forms.Get(i).Name for i in range(cat.Forms.Count())]
    fails = [f"нет формы {f} в базе" for f in ("ФормаЭлемента", "ФормаСписка", "ФормаГруппы", "ФормаПодбораКарточки",
                                                 "ФормаПерепривязки") if f not in have]
    for prop, form in (("DefaultObjectForm", "ФормаЭлемента"), ("DefaultListForm", "ФормаСписка"),
                       ("DefaultFolderForm", "ФормаГруппы")):
        if getattr(cat, prop).Name != form:
            fails.append(f"{prop} = {getattr(cat, prop).Name}")
    return fails


if __name__ == "__main__":
    fails = check_files() if "--files" in sys.argv else check_com()
    for f in fails:
        print("FAIL:", f)
    print(f"ПРОВАЛОВ: {len(fails)}")
    sys.exit(1 if fails else 0)
```

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py --files`
Expected: FAIL — среди строк `FAIL: ФормаЭлемента: нет команды ЗаполнитьИзСС` (так же для остальных 10 команд этапа 2), `FAIL: ФормаСписка: нет команды ПерепривязатьПозицииСС`, `FAIL: ФормаПодбораКарточки: нет файла …`, `FAIL: ФормаПерепривязки: нет файла …`; код 1.

- [ ] **Step 3: Реализовать модуль формы этапа 2**

`_Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_2.bsl`:

```bsl
#Область ОбработчикиКомандФормыЭтап2

&НаКлиенте
Процедура ЗаполнитьИзСС(Команда)

	Если Не ЗначениеЗаполнено(Объект.СтруктураСебестоимости) Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Укажите тип дома (СС).'; uk = 'Вкажіть тип будинку (СС).'"));
		Возврат;
	КонецЕсли;
	Если Объект.Состав.Количество() = 0 Тогда
		ЗаполнитьИзССЗавершение(КодВозвратаДиалога.Да, Неопределено);
		Возврат;
	КонецЕсли;
	ПоказатьВопрос(Новый ОписаниеОповещения("ЗаполнитьИзССЗавершение", ЭтотОбъект),
		НСтр("ru = 'Состав будет очищен и заполнен заново из СС. Продолжить?'; uk = 'Склад буде очищено й заповнено заново з СС. Продовжити?'"),
		РежимДиалогаВопрос.ДаНет);

КонецПроцедуры

&НаКлиенте
Процедура ЗаполнитьИзССЗавершение(Ответ, ДополнительныеПараметры) Экспорт

	Если Ответ <> КодВозвратаДиалога.Да Тогда
		Возврат;
	КонецЕсли;
	ПоказатьПредупреждение(, ЗаполнитьИзССНаСервере());

КонецПроцедуры

&НаКлиенте
Процедура ОтметитьЗеленыеПроверенными(Команда)

	Отмечено = ОтметитьЗеленыеПроверенныеНаСервере();
	ПоказатьПредупреждение(, СтрШаблон(НСтр("ru = 'Отмечено строк: %1.'; uk = 'Позначено рядків: %1.'"), Отмечено));

КонецПроцедуры

&НаКлиенте
Процедура ПодобратьКарточку(Команда)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Если ТекущиеДанные.ВидСтроки <> ПредопределенноеЗначение("Перечисление.А_ВидыСтрокСпецификацииДомов.Деталь") Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Карточку подбирают для строки «Деталь».'; uk = 'Картку підбирають для рядка «Деталь».'"));
		Возврат;
	КонецЕсли;
	ПараметрыФормы = Новый Структура("ТипБудинкуСС, ОбщееНазвание, ПозицияСС",
		Объект.СтруктураСебестоимости, ТекущиеДанные.ОбщееНазвание, ТекущиеДанные.ПозицияСС);
	ОткрытьФорму("Справочник.А_СпецификацииДомов.Форма.ФормаПодбораКарточки", ПараметрыФормы, ЭтотОбъект, , , ,
		Новый ОписаниеОповещения("ПодобратьКарточкуЗавершение", ЭтотОбъект, ТекущиеДанные.ПолучитьИдентификатор()),
		РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);

КонецПроцедуры

&НаКлиенте
Процедура ПодобратьКарточкуЗавершение(Карточка, ИдентификаторСтроки) Экспорт

	Если Не ЗначениеЗаполнено(Карточка) Тогда
		Возврат;
	КонецЕсли;
	Подсказка = ПрименитьКарточкуНаСервере(ИдентификаторСтроки, Карточка);
	Если Не ПустаяСтрока(Подсказка) Тогда
		ПоказатьОповещениеПользователя(НСтр("ru = 'Коэффициент'; uk = 'Коефіцієнт'"), , Подсказка);
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура ВзятьИзОригиналаExcel(Команда)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Если Не ВзятьИзОригиналаExcelНаСервере(ТекущиеДанные.ПолучитьИдентификатор()) Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'У позиции нет количества и единицы оригинала Excel.'; uk = 'У позиції немає кількості й одиниці оригіналу Excel.'"));
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура ОбъединитьВУзел(Команда)

	Выделенные = Элементы.Состав.ВыделенныеСтроки;
	Если Выделенные.Количество() = 0 Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Выделите строки будущего узла.'; uk = 'Виділіть рядки майбутнього вузла.'"));
		Возврат;
	КонецЕсли;
	Контекст = Новый Структура("Строки, ИмяУзла", Новый ФиксированныйМассив(Выделенные), "");
	Список = Новый СписокЗначений;
	Для Каждого ИмяИзСписка Из ИменаУзловНаСервере() Цикл
		Список.Добавить(ИмяИзСписка);
	КонецЦикла;
	Список.Добавить("", НСтр("ru = 'Другое название…'; uk = 'Інша назва…'"));
	Список.ПоказатьВыборЭлемента(Новый ОписаниеОповещения("ОбъединитьВУзелИмяВыбрано", ЭтотОбъект, Контекст),
		НСтр("ru = 'Название узла'; uk = 'Назва вузла'"));

КонецПроцедуры

&НаКлиенте
Процедура ОбъединитьВУзелИмяВыбрано(ВыбранныйЭлемент, Контекст) Экспорт

	Если ВыбранныйЭлемент = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Если ПустаяСтрока(ВыбранныйЭлемент.Значение) Тогда
		ПоказатьВводСтроки(Новый ОписаниеОповещения("ОбъединитьВУзелИмяВведено", ЭтотОбъект, Контекст), "",
			НСтр("ru = 'Название узла'; uk = 'Назва вузла'"), 150);
		Возврат;
	КонецЕсли;
	ОбъединитьВУзелИмяВведено(ВыбранныйЭлемент.Значение, Контекст);

КонецПроцедуры

&НаКлиенте
Процедура ОбъединитьВУзелИмяВведено(ИмяУзла, Контекст) Экспорт

	Если ИмяУзла = Неопределено Или ПустаяСтрока(ИмяУзла) Тогда
		Возврат;
	КонецЕсли;
	Контекст.ИмяУзла = ИмяУзла;
	ПоказатьВводЧисла(Новый ОписаниеОповещения("ОбъединитьВУзелКоличествоВведено", ЭтотОбъект, Контекст), 1,
		НСтр("ru = 'Узлов в выделенных строках (количество выпуска)'; uk = 'Вузлів у виділених рядках (кількість випуску)'"), 15, 3);

КонецПроцедуры

&НаКлиенте
Процедура ОбъединитьВУзелКоличествоВведено(КоличествоВыпуска, Контекст) Экспорт

	Если КоличествоВыпуска = Неопределено Или КоличествоВыпуска <= 0 Тогда
		Возврат;
	КонецЕсли;
	ПоказатьПредупреждение(, ОбъединитьВУзелНаСервере(Контекст.Строки, Контекст.ИмяУзла, КоличествоВыпуска));

КонецПроцедуры

&НаКлиенте
Процедура РасформироватьУзел(Команда)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	ПоказатьПредупреждение(, РасформироватьУзелНаСервере(ТекущиеДанные.ПолучитьИдентификатор()));

КонецПроцедуры

&НаКлиенте
Процедура СоставИзДокументаКомплектации(Команда)

	ТекущиеДанные = Элементы.Состав.ТекущиеДанные;
	Контекст = Новый Структура("ИдентификаторСтроки, Документ, ИмяУзла",
		?(ТекущиеДанные = Неопределено, Неопределено, ТекущиеДанные.ПолучитьИдентификатор()), Неопределено, "");
	ПоказатьВводЗначения(Новый ОписаниеОповещения("СоставИзДокументаКомплектацииДокументВыбран", ЭтотОбъект, Контекст),
		ПредопределенноеЗначение("Документ.КомплектацияНоменклатуры.ПустаяСсылка"),
		НСтр("ru = 'Документ комплектации'; uk = 'Документ комплектації'"), Тип("ДокументСсылка.КомплектацияНоменклатуры"));

КонецПроцедуры

&НаКлиенте
Процедура СоставИзДокументаКомплектацииДокументВыбран(ДокументКомплектации, Контекст) Экспорт

	Если Не ЗначениеЗаполнено(ДокументКомплектации) Тогда
		Возврат;
	КонецЕсли;
	Контекст.Документ = ДокументКомплектации;
	Данные = ДанныеДокументаКомплектации(ДокументКомплектации);
	Контекст.Вставить("Выпуск", Данные.Выпуск);
	ПоказатьВводСтроки(Новый ОписаниеОповещения("СоставИзДокументаКомплектацииИмяВведено", ЭтотОбъект, Контекст), Данные.Имя,
		НСтр("ru = 'Название узла'; uk = 'Назва вузла'"), 150);

КонецПроцедуры

&НаКлиенте
Процедура СоставИзДокументаКомплектацииИмяВведено(ИмяУзла, Контекст) Экспорт

	Если ИмяУзла = Неопределено Или ПустаяСтрока(ИмяУзла) Тогда
		Возврат;
	КонецЕсли;
	Контекст.ИмяУзла = ИмяУзла;
	ПоказатьВводЧисла(Новый ОписаниеОповещения("СоставИзДокументаКомплектацииКоличествоВведено", ЭтотОбъект, Контекст), 1,
		СтрШаблон(НСтр("ru = 'Состав документа = N узлов (выпуск документа: %1)'; uk = 'Склад документа = N вузлів (випуск документа: %1)'"),
			Контекст.Выпуск), 15, 3);

КонецПроцедуры

&НаКлиенте
Процедура СоставИзДокументаКомплектацииКоличествоВведено(УзловВДокументе, Контекст) Экспорт

	Если УзловВДокументе = Неопределено Или УзловВДокументе <= 0 Тогда
		Возврат;
	КонецЕсли;
	ПоказатьПредупреждение(, СоставИзДокументаКомплектацииНаСервере(Контекст.ИдентификаторСтроки, Контекст.Документ,
		Контекст.ИмяУзла, УзловВДокументе));

КонецПроцедуры

&НаКлиенте
Процедура РассчитатьЦеныНормы(Команда)

	ПоказатьПредупреждение(, РассчитатьЦеныНормыНаСервере());

КонецПроцедуры

&НаКлиенте
Процедура СравнитьСоСС(Команда)

	СравнитьСоССНаСервере();
	Элементы.Страницы.ТекущаяСтраница = Элементы.СтраницаСравнение;

КонецПроцедуры

&НаКлиенте
Процедура НулевыеПозицииСФактом(Команда)

	Если Не ЗначениеЗаполнено(Объект.СтруктураСебестоимости) Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Укажите тип дома (СС).'; uk = 'Вкажіть тип будинку (СС).'"));
		Возврат;
	КонецЕсли;
	ТабДок = НулевыеПозицииНаСервере();
	ТабДок.Показать(НСтр("ru = 'Нулевые позиции с фактом'; uk = 'Нульові позиції з фактом'"));

КонецПроцедуры

&НаКлиенте
Процедура СкопироватьИзДругойСпецификации(Команда)

	ПоказатьВводЗначения(Новый ОписаниеОповещения("СкопироватьИзДругойСпецификацииИсточникВыбран", ЭтотОбъект),
		ПредопределенноеЗначение("Справочник.А_СпецификацииДомов.ПустаяСсылка"),
		НСтр("ru = 'Спецификация-источник (эталон другого типа)'; uk = 'Специфікація-джерело (еталон іншого типу)'"),
		Тип("СправочникСсылка.А_СпецификацииДомов"));

КонецПроцедуры

&НаКлиенте
Процедура СкопироватьИзДругойСпецификацииИсточникВыбран(Источник, ДополнительныеПараметры) Экспорт

	Если Не ЗначениеЗаполнено(Источник) Тогда
		Возврат;
	КонецЕсли;
	ПоказатьПредупреждение(, СкопироватьИзДругойСпецификацииНаСервере(Источник));

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункцииЭтап2

&НаСервере
Функция ИндексСтрокиПоИдентификатору(ИдентификаторСтроки)

	СтрокаСостава = Объект.Состав.НайтиПоИдентификатору(ИдентификаторСтроки);
	Возврат ?(СтрокаСостава = Неопределено, Неопределено, Объект.Состав.Индекс(СтрокаСостава));

КонецФункции

&НаСервере
Функция ЗаполнитьИзССНаСервере()

	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Попытка
		Протокол = А_СпецификацииДомовСервер.ЗаполнитьИзСС(СпецОбъект);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Текст = СтрШаблон(НСтр("ru = 'Строк: %1 (зелёных %2, жёлтых %3, без карточки %4, групп %5); нулевых позиций пропущено: %6; сумма СС: %7.'; uk = 'Рядків: %1 (зелених %2, жовтих %3, без картки %4, груп %5); нульових позицій пропущено: %6; сума СС: %7.'"),
		Протокол.Строк, Протокол.Зеленых, Протокол.Желтых, Протокол.Пустых, Протокол.Групп, Протокол.ПропущеноНулевых,
		Формат(Протокол.Сумма, "ЧДЦ=2; ЧРГ=' '"));
	Если Протокол.ПропущеноНулевых > 0 Тогда
		Текст = Текст + Символы.ПС + НСтр("ru = 'Список нулевых позиций с фактом — команда «Нулевые позиции с фактом».'; uk = 'Список нульових позицій з фактом — команда «Нульові позиції з фактом».'");
	КонецЕсли;
	Если Протокол.Комплекты.Количество() > 0 Тогда
		Текст = Текст + Символы.ПС + НСтр("ru = 'Позиции-комплекты (состав — из документа комплектации):'; uk = 'Позиції-комплекти (склад — з документа комплектації):'")
			+ Символы.ПС + СтрСоединить(Протокол.Комплекты, Символы.ПС);
	КонецЕсли;
	Возврат Текст;

КонецФункции

&НаСервере
Функция ОтметитьЗеленыеПроверенныеНаСервере()

	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Отмечено = А_СпецификацииДомовСервер.ОтметитьЗеленыеПроверенными(СпецОбъект);
	Если Отмечено > 0 Тогда
		ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
		ЭтотОбъект.Модифицированность = Истина;
	КонецЕсли;
	Возврат Отмечено;

КонецФункции

// Карточка из подбора (Р35): «Вручну», «Перевірено» снят, ОН и k — как при ручной смене карточки; если единицы не совпали —
// подсказка k из названия карточки (Р23), её бухгалтер подтверждает отметкой «Перевірено».
&НаСервере
Функция ПрименитьКарточкуНаСервере(ИдентификаторСтроки, Карточка)

	СтрокаСостава = Объект.Состав.НайтиПоИдентификатору(ИдентификаторСтроки);
	Если СтрокаСостава = Неопределено Тогда
		Возврат "";
	КонецЕсли;
	СтрокаСостава.Номенклатура = Карточка;
	СтрокаСостава.ИсточникКарточки = Перечисления.А_ИсточникиКарточкиСпецификации.Вручную;
	СтрокаСостава.Проверено = Ложь;
	ОбновитьСтрокуПослеСменыКарточки(ИдентификаторСтроки);
	ЭтотОбъект.Модифицированность = Истина;
	Если СтрокаСостава.ИсточникКоэффициента <> Перечисления.А_ИсточникиКоэффициентаПересчетаСС.НеНайден Тогда
		Возврат "";
	КонецЕсли;
	Подсказка = А_СпецификацииДомовСервер.ПодсказкаКоэффициента(Карточка, СтрокаСостава.ЕдиницаСтроки);
	Если Подсказка.Коэффициент > 0 Тогда
		СтрокаСостава.Коэффициент = Подсказка.Коэффициент;
		СтрокаСостава.ИсточникКоэффициента = Перечисления.А_ИсточникиКоэффициентаПересчетаСС.Вручную;
		Возврат СтрШаблон(НСтр("ru = 'Подсказка из названия: %1 — проверьте и отметьте «Проверено».'; uk = 'Підказка з назви: %1 — перевірте й позначте «Перевірено».'"),
			Подсказка.Текст);
	КонецЕсли;
	Возврат НСтр("ru = 'Единицы не совпали — укажите k вручную.'; uk = 'Одиниці не збіглися — вкажіть k вручну.'");

КонецФункции

&НаСервере
Функция ВзятьИзОригиналаExcelНаСервере(ИдентификаторСтроки)

	Индекс = ИндексСтрокиПоИдентификатору(ИдентификаторСтроки);
	Если Индекс = Неопределено Тогда
		Возврат Ложь;
	КонецЕсли;
	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Если Не А_СпецификацииДомовСервер.ВзятьИзОригиналаExcel(СпецОбъект, Индекс) Тогда
		Возврат Ложь;
	КонецЕсли;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Возврат Истина;

КонецФункции

&НаСервереБезКонтекста
Функция ИменаУзловНаСервере()

	Возврат А_СпецификацииДомовСервер.ИменаУзловСтроителя();

КонецФункции

&НаСервере
Функция ОбъединитьВУзелНаСервере(ИдентификаторыСтрок, ИмяУзла, КоличествоВыпуска)

	Индексы = Новый Массив;
	Для Каждого ИдентификаторСтроки Из ИдентификаторыСтрок Цикл
		Индекс = ИндексСтрокиПоИдентификатору(ИдентификаторСтроки);
		Если Индекс <> Неопределено Тогда
			Индексы.Добавить(Индекс);
		КонецЕсли;
	КонецЦикла;
	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Попытка
		Узел = А_СпецификацииДомовСервер.ОбъединитьВУзел(СпецОбъект, Индексы, ИмяУзла, КоличествоВыпуска, Неопределено);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Возврат СтрШаблон(НСтр("ru = 'Создан узел «%1», строк в нём: %2. Норма спецификации не изменилась — запишите спецификацию.'; uk = 'Створено вузол «%1», рядків у ньому: %2. Норма специфікації не змінилась — запишіть специфікацію.'"),
		Узел, Индексы.Количество());

КонецФункции

&НаСервере
Функция РасформироватьУзелНаСервере(ИдентификаторСтроки)

	Индекс = ИндексСтрокиПоИдентификатору(ИдентификаторСтроки);
	Если Индекс = Неопределено Тогда
		Возврат "";
	КонецЕсли;
	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Попытка
		А_СпецификацииДомовСервер.РасформироватьУзел(СпецОбъект, Индекс);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Возврат НСтр("ru = 'Строки узла возвращены в спецификацию. Сам узел не удалён.'; uk = 'Рядки вузла повернуто в специфікацію. Сам вузол не видалено.'");

КонецФункции

&НаСервереБезКонтекста
Функция ДанныеДокументаКомплектации(ДокументКомплектации)

	Реквизиты = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(ДокументКомплектации, "Номенклатура, Количество");
	Возврат Новый Структура("Имя, Выпуск", Строка(Реквизиты.Номенклатура), Реквизиты.Количество);

КонецФункции

// «Склад з документа комплектації» (Р31): узел из документа ставится в текущую строку (её позиция и количество сохраняются)
// или — без текущей строки — новой строкой × 1.
&НаСервере
Функция СоставИзДокументаКомплектацииНаСервере(ИдентификаторСтроки, ДокументКомплектации, ИмяУзла, УзловВДокументе)

	Индекс = ?(ИдентификаторСтроки = Неопределено, Неопределено, ИндексСтрокиПоИдентификатору(ИдентификаторСтроки));
	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Попытка
		Состав = А_СпецификацииДомовСервер.СоставИзДокументаКомплектации(ДокументКомплектации, УзловВДокументе);
		Узел = А_СпецификацииДомовСервер.УзелИзДокументаКомплектации(ДокументКомплектации, ИмяУзла, УзловВДокументе,
			Неопределено, СпецОбъект.СтруктураСебестоимости);
		А_СпецификацииДомовСервер.ПоставитьУзелВСтроку(СпецОбъект, Индекс, Узел);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Возврат СтрШаблон(НСтр("ru = 'Создан узел «%1»: деталей %2 (выпуск документа %3, узлов в документе %4).'; uk = 'Створено вузол «%1»: деталей %2 (випуск документа %3, вузлів у документі %4).'"),
		Узел, Состав.Строки.Количество(), Состав.КоличествоВыпускаДокумента, УзловВДокументе);

КонецФункции

&НаСервере
Функция РассчитатьЦеныНормыНаСервере()

	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Попытка
		Протокол = А_СпецификацииДомовСервер.РассчитатьЦеныНормы(СпецОбъект);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Текст = СтрШаблон(НСтр("ru = 'Позиций: %1; Σ СС: %2; Σ нормы: %3; разница (И4): %4. Цены узлов записаны; запишите спецификацию.'; uk = 'Позицій: %1; Σ СС: %2; Σ норми: %3; різниця (И4): %4. Ціни вузлів записано; запишіть специфікацію.'"),
		Протокол.Позиций, Формат(Протокол.СуммаСС, "ЧДЦ=2; ЧРГ=' '"), Формат(Протокол.СуммаНормы, "ЧДЦ=2; ЧРГ=' '"),
		Формат(Протокол.И4, "ЧДЦ=2; ЧН=0"));
	Если Протокол.БезЛистьев.Количество() > 0 Тогда
		Текст = Текст + Символы.ПС + НСтр("ru = 'Позиции СС без строк:'; uk = 'Позиції СС без рядків:'") + Символы.ПС
			+ СтрСоединить(Протокол.БезЛистьев, Символы.ПС);
	КонецЕсли;
	Если Протокол.ПропущеноУзлов.Количество() > 0 Тогда
		Текст = Текст + Символы.ПС + НСтр("ru = 'Узлы другой СС — цены записаны только в пустые строки:'; uk = 'Вузли іншої СС — ціни записано лише в порожні рядки:'")
			+ Символы.ПС + СтрСоединить(Протокол.ПропущеноУзлов, Символы.ПС);
	КонецЕсли;
	Возврат Текст;

КонецФункции

&НаСервере
Процедура СравнитьСоССНаСервере()

	СравнениеСС.Очистить();
	Попытка
		Итог = А_СпецификацииДомовСервер.СравнитьСоСтруктуройСебестоимости(РеквизитФормыВЗначение("Объект"));
	Исключение
		ИтогСравнения = КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
		Возврат;
	КонецПопытки;
	Для Каждого Позиция Из Итог.Позиции Цикл
		ЗаполнитьЗначенияСвойств(СравнениеСС.Добавить(), Позиция);
	КонецЦикла;
	Для Каждого Сирота Из Итог.Сироты Цикл
		НоваяСтрока = СравнениеСС.Добавить();
		НоваяСтрока.ПозицияСС = Сирота.ПозицияСС;
		НоваяСтрока.КоличествоСпец = Сирота.КоличествоЕдСтроки;
		НоваяСтрока.Статус = НСтр("ru = 'Сирота — позиции нет в СС (перепривяжите)'; uk = 'Сирота — позиції немає в СС (перепривʼяжіть)'");
	КонецЦикла;
	Для Каждого Лист Из Итог.ДобавленоВДоме Цикл
		НоваяСтрока = СравнениеСС.Добавить();
		НоваяСтрока.ПозицияСС = Строка(Лист.Номенклатура);
		НоваяСтрока.КоличествоСпец = Лист.Количество;
		НоваяСтрока.СуммаНормы = Лист.СуммаНормы;
		НоваяСтрока.Статус = НСтр("ru = 'Добавлено в доме'; uk = 'Додано в будинку'");
	КонецЦикла;
	ИтогСравнения = СтрШаблон(НСтр("ru = 'Σ нормы %1 / Σ СС %2 (разница %3, в т. ч. не поставляется %4); позиций без строк: %5; сирот: %6'; uk = 'Σ норми %1 / Σ СС %2 (різниця %3, зокрема не постачається %4); позицій без рядків: %5; сиріт: %6'"),
		Формат(Итог.СуммаНормы, "ЧДЦ=2; ЧРГ=' '"), Формат(Итог.СуммаСС, "ЧДЦ=2; ЧРГ=' '"), Формат(Итог.Разница, "ЧДЦ=2; ЧН=0"),
		Формат(Итог.СуммаНеПоставляется, "ЧДЦ=2; ЧН=0"), Итог.ПозицийБезСтрок, Итог.Сироты.Количество());

КонецПроцедуры

&НаСервере
Функция НулевыеПозицииНаСервере()

	Возврат А_СпецификацииДомовСервер.НулевыеПозицииВТабличныйДокумент(РеквизитФормыВЗначение("Объект"));

КонецФункции

&НаСервере
Функция СкопироватьИзДругойСпецификацииНаСервере(Источник)

	Если Источник = Объект.Ссылка Тогда
		Возврат НСтр("ru = 'Источник совпадает с этой спецификацией.'; uk = 'Джерело збігається з цією специфікацією.'");
	КонецЕсли;
	СпецОбъект = РеквизитФормыВЗначение("Объект");
	Попытка
		Протокол = А_СпецификацииДомовСервер.СкопироватьИзДругойСпецификации(СпецОбъект, Источник);
	Исключение
		Возврат КраткоеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки;
	ЗначениеВРеквизитФормы(СпецОбъект, "Объект");
	ЭтотОбъект.Модифицированность = Истина;
	Текст = СтрШаблон(НСтр("ru = 'Карточек без расхождения с фактом: %1; жёлтых (расходятся с фактом типа): %2; не закрыто: %3; «Не поставляется» скопировано: %4; узлов переиспользовано: %5, скопировано: %6.'; uk = 'Карток без розбіжності з фактом: %1; жовтих (розходяться з фактом типу): %2; не закрито: %3; «Не постачається» скопійовано: %4; вузлів перевикористано: %5, скопійовано: %6.'"),
		Протокол.Совпало, Протокол.Отличий, Протокол.НеЗакрыто, Протокол.НеПоставляетсяСкопировано,
		Протокол.УзловПереиспользовано, Протокол.УзловСкопировано);
	Первые = Новый Массив;
	Для Номер = 0 По Мин(29, Протокол.Отличия.ВГраница()) Цикл
		Первые.Добавить(Протокол.Отличия[Номер]);
	КонецЦикла;
	Если Первые.Количество() > 0 Тогда
		Текст = Текст + Символы.ПС + СтрСоединить(Первые, Символы.ПС);
	КонецЕсли;
	Возврат Текст;

КонецФункции

#КонецОбласти
```

- [ ] **Step 4: PASS — статика исходника сейчас, файлы формы — в Task 42, COM — в Task 43**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py _Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_2.bsl`
Expected: кроме «нет BOM», «не CRLF» ошибок нет (в частности, нет «присваивание зарезервированному имени» — `Модифицированность` меняется только как `ЭтотОбъект.Модифицированность`).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/forma_elementa_2.bsl _Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py
git commit -m "feat(spec_domov): модуль формы спецификации этапа 2 — заполнение, подбор, узлы, цены, сравнение, копия

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 42: Формы подбора карточки, перепривязки и список этапа 2; генерация форм

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/bsl/forma_podbora.bsl`, `_Rarzrabotki/Python/test/spec_domov/bsl/forma_perepriv.bsl`, `_Rarzrabotki/Python/test/spec_domov/bsl/forma_spiska_2.bsl`
- Create (генерируются): `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаПодбораКарточки.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl`; `…/Forms/ФормаПерепривязки.xml` + `/Ext/Form.xml` + `/Ext/Form/Module.bsl`
- Modify (перегенерируются): `…/Forms/ФормаЭлемента/Ext/Form.xml` + `/Ext/Form/Module.bsl`, `…/Forms/ФормаСписка/Ext/Form.xml` + `/Ext/Form/Module.bsl`, `_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml` (`<Form>` двух новых форм)
- Test: `test_e2d_formy.py --files` (Task 41)

- [ ] **Step 1: Тест** — `test_e2d_formy.py --files` (Task 41).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py --files`
Expected: FAIL — `ФормаПодбораКарточки: нет файла …`, `ФормаЭлемента: нет команды ЗаполнитьИзСС`.

- [ ] **Step 3: Реализовать модули и сгенерировать формы**

`_Rarzrabotki/Python/test/spec_domov/bsl/forma_podbora.bsl`:

```bsl
#Область ОбработчикиСобытийФормы

// Подбор карточки с контекстом (Р35): кандидаты ОН позиции или — без «Лише загальна назва позиції» — поиск по названию во
// всей номенклатуре; колонки факта по домам типа, склада проекта, цены последней закупки, совпадения размера. Выбор
// закрывает форму и возвращает карточку владельцу.
&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)

	ТипБудинкуСС = Параметры.ТипБудинкуСС;
	ОбщееНазвание = Параметры.ОбщееНазвание;
	ПозицияСС = Параметры.ПозицияСС;
	ТолькоОН = ЗначениеЗаполнено(ОбщееНазвание);
	ЗаполнитьКандидатов();

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ТолькоОНПриИзменении(Элемент)

	ЗаполнитьКандидатов();

КонецПроцедуры

&НаКлиенте
Процедура СтрокаПоискаПриИзменении(Элемент)

	ЗаполнитьКандидатов();

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыКандидаты

&НаКлиенте
Процедура КандидатыВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)

	СтандартнаяОбработка = Ложь;
	ВыбратьТекущуюКарточку();

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ВыбратьКарточку(Команда)

	ВыбратьТекущуюКарточку();

КонецПроцедуры

&НаКлиенте
Процедура ОбновитьКандидатов(Команда)

	ЗаполнитьКандидатов();

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура ВыбратьТекущуюКарточку()

	ТекущиеДанные = Элементы.Кандидаты.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Закрыть(ТекущиеДанные.Номенклатура);

КонецПроцедуры

&НаСервере
Процедура ЗаполнитьКандидатов()

	Кандидаты.Очистить();
	Если Не ЗначениеЗаполнено(ТипБудинкуСС) Тогда
		Возврат;
	КонецЕсли;
	Если Не ТолькоОН И ПустаяСтрока(СтрокаПоиска) Тогда
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(НСтр("ru = 'Без общего названия позиции введите строку поиска.'; uk = 'Без загальної назви позиції введіть рядок пошуку.'"));
		Возврат;
	КонецЕсли;
	Для Каждого Карточка Из А_СпецификацииДомовСервер.ПредложитьКарточки(ТипБудинкуСС, ОбщееНазвание, ПозицияСС, ТолькоОН, СтрокаПоиска) Цикл
		ЗаполнитьЗначенияСвойств(Кандидаты.Добавить(), Карточка);
	КонецЦикла;

КонецПроцедуры

#КонецОбласти
```

`_Rarzrabotki/Python/test/spec_domov/bsl/forma_perepriv.bsl`:

```bsl
#Область ОбработчикиСобытийФормы

// «Перепривʼязати позиції СС» ([M5] §3.10, Т10): сироты всех спецификаций типа и их узлов с предложениями; зелёные отмечены
// к применению, жёлтые и узлы другой СС бухгалтер отмечает сам.
&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)

	ТипБудинкуСС = Параметры.ТипБудинкуСС;
	ДобавитьОформлениеПредложения("Зелений", Новый Цвет(226, 239, 218));
	ДобавитьОформлениеПредложения("Жовтий", Новый Цвет(255, 242, 204));
	Если ЗначениеЗаполнено(ТипБудинкуСС) Тогда
		НайтиПредложенияНаСервере();
	КонецЕсли;

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ТипБудинкуССПриИзменении(Элемент)

	НайтиПредложенияНаСервере();

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура НайтиПредложения(Команда)

	НайтиПредложенияНаСервере();

КонецПроцедуры

&НаКлиенте
Процедура ПрименитьОтмеченные(Команда)

	Изменено = ПрименитьОтмеченныеНаСервере();
	ПоказатьПредупреждение(, СтрШаблон(НСтр("ru = 'Перепривязано строк: %1.'; uk = 'Перепривʼязано рядків: %1.'"), Изменено));

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура НайтиПредложенияНаСервере()

	Предложения.Очистить();
	Если Не ЗначениеЗаполнено(ТипБудинкуСС) Тогда
		Возврат;
	КонецЕсли;
	Для Каждого Предложение Из А_СпецификацииДомовСервер.ПредложенияПерепривязки(ТипБудинкуСС) Цикл
		ЗаполнитьЗначенияСвойств(Предложения.Добавить(), Предложение);
	КонецЦикла;
	Если Предложения.Количество() = 0 Тогда
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(НСтр("ru = 'Сирот нет: все позиции строк есть в СС.'; uk = 'Сиріт немає: усі позиції рядків є в СС.'"));
	КонецЕсли;

КонецПроцедуры

&НаСервере
Функция ПрименитьОтмеченныеНаСервере()

	Изменено = А_СпецификацииДомовСервер.ПерепривязатьПозицииСС(РеквизитФормыВЗначение("Предложения"));
	НайтиПредложенияНаСервере();
	Возврат Изменено;

КонецФункции

&НаСервере
Процедура ДобавитьОформлениеПредложения(ЗначениеЦвета, ЦветФона)

	ЭлементОформления = УсловноеОформление.Элементы.Добавить();
	ОформляемоеПоле = ЭлементОформления.Поля.Элементы.Добавить();
	ОформляемоеПоле.Поле = Новый ПолеКомпоновкиДанных("Предложения");
	ЭлементОтбора = ЭлементОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("Предложения.Цвет");
	ЭлементОтбора.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = ЗначениеЦвета;
	ЭлементОформления.Оформление.УстановитьЗначениеПараметра("ЦветФона", ЦветФона);

КонецПроцедуры

#КонецОбласти
```

`_Rarzrabotki/Python/test/spec_domov/bsl/forma_spiska_2.bsl`:

```bsl
#Область ОбработчикиКомандФормыЭтап2

&НаКлиенте
Процедура ПерепривязатьПозицииСС(Команда)

	ТипБудинкуСС = Неопределено;
	ТекущиеДанные = Элементы.Список.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено И ЗначениеЗаполнено(ТекущиеДанные.СтруктураСебестоимости) Тогда
		ТипБудинкуСС = ТекущиеДанные.СтруктураСебестоимости;
	КонецЕсли;
	ОткрытьФорму("Справочник.А_СпецификацииДомов.Форма.ФормаПерепривязки", Новый Структура("ТипБудинкуСС", ТипБудинкуСС), ЭтотОбъект);

КонецПроцедуры

#КонецОбласти
```

Сгенерировать формы этапа 2:

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/gen_forms.py --etap 2`
Expected: `форма ФормаЭлемента`, `форма ФормаСписка`, `форма ФормаПодбораКарточки`, `форма ФормаПерепривязки`, `OK` (`ФормаГруппы` на этапе 2 не перегенерируется: она не в списке e2d, и её worktree-копия должна остаться равной основному каталогу для `to_main.py --verify` в Task 45).

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py --files`
Expected: `ПРОВАЛОВ: 0`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e1_files.py --only forms`
Expected: `forms OK` (формы этапа 1 не сломаны перегенерацией).
Run (для каждой из пяти форм): `powershell.exe -NoProfile -File C:/Configuration_downloads/BASERP25/.claude/skills/form-validate/scripts/form-validate.ps1 -FormPath "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms/ФормаПодбораКарточки"`
Expected: ошибок нет; допустимы предупреждение о версии 2.13 и ложные `~Список.X`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py "Catalogs/А_СпецификацииДомов.xml" "Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml" "Catalogs/А_СпецификацииДомов/Forms/ФормаСписка.xml" "Catalogs/А_СпецификацииДомов/Forms/ФормаПодбораКарточки.xml" "Catalogs/А_СпецификацииДомов/Forms/ФормаПерепривязки.xml"`
Expected: `проверено файлов: 13, ошибок: 0` (корень формы раскрывается в `Ext/Form.xml` и модуль; в модулях форм — директива компиляции перед каждой процедурой, нет присваиваний `Модифицированность`/`Объект`/`Элементы`).

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/bsl/forma_podbora.bsl _Rarzrabotki/Python/test/spec_domov/bsl/forma_perepriv.bsl \
  _Rarzrabotki/Python/test/spec_domov/bsl/forma_spiska_2.bsl \
  "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов.xml" "_Rarzrabotki/BASEBuh/Catalogs/А_СпецификацииДомов/Forms"
git commit -m "feat(BASEBuh): формы этапа 2 — подбор карточки с фактом и складом (Р35), перепривязка позиций СС, команды бухгалтера

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 43: Загрузка 2D (формы) и проверка форм в базе

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/load_lists/e2d.txt` (5 строк из раздела «Списки Partial-загрузки»)
- Create (результат): `_Rarzrabotki/Python/test/spec_domov/out/e2d_load_<stamp>.txt`

- [ ] **Step 1: Тест** — `test_e2d_formy.py` без ключа (COM, Task 41).

- [ ] **Step 2: Запустить — FAIL**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py`
Expected: FAIL — `нет формы ФормаПодбораКарточки в базе`, `нет формы ФормаПерепривязки в базе`, `ПРОВАЛОВ: 2`.

- [ ] **Step 3: Статика, копия в основной каталог, загрузка**

`_Rarzrabotki/Python/test/spec_domov/load_lists/e2d.txt`:

```
Catalogs/А_СпецификацииДомов.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаСписка.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаПодбораКарточки.xml
Catalogs/А_СпецификацииДомов/Forms/ФормаПерепривязки.xml
```

Run:
```bash
python -X utf8 _Rarzrabotki/Python/test/spec_domov/check_static.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2d.txt --tests
python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --list _Rarzrabotki/Python/test/spec_domov/load_lists/e2d.txt
powershell.exe -NoProfile -File _Rarzrabotki/Python/test/spec_domov/load_stage.ps1 -List _Rarzrabotki/Python/test/spec_domov/load_lists/e2d.txt -Tag e2d
```
Expected: `ошибок: 0`; 13 строк `скопирован …` (5 файлов списка и 8 файлов `Ext` четырёх форм — `to_main.py` раскрывает корень формы) без `ОТКАЗ`; `OK load e2d`. Ошибка `CheckModules` вида `{Справочник.А_СпецификацииДомов.Форма.ФормаЭлемента.Форма(строка,колонка)}` — исправить исходник в `bsl/forma_*.bsl`, повторить `gen_forms.py --etap 2` → `to_main.py --list …e2d.txt` → `load_stage.ps1`.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2d_formy.py`
Expected: `ПРОВАЛОВ: 0`.
Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/dump_check.py --tag e2d --canon --objects "Справочник.А_СпецификацииДомов,Справочник.А_СпецификацииДомов.Форма.ФормаЭлемента,Справочник.А_СпецификацииДомов.Форма.ФормаСписка,Справочник.А_СпецификацииДомов.Форма.ФормаПодбораКарточки,Справочник.А_СпецификацииДомов.Форма.ФормаПерепривязки"`
Expected: `различие по содержимому 0`.
Ручная проверка в клиенте (5 минут, ничего не записывать — закрыть форму с «Ні» на вопрос о сохранении): новый элемент вида «Будинок», СС 30 м → «Заповнити з СС» (протокол: 170 рядків, 17 зелених); на строке поз. 42 «Підібрати картку» — форма подбора открылась с колонками факта и склада; «Нульові позиції з фактом» открывает табличный документ; в списке «Перепривʼязати позиції СС…» открывает форму перепривязки. **Не нажимать** при ручной проверке «Обʼєднати у вузол», «Склад з документа комплектації», «Розрахувати ціни норми», «Скопіювати картки й дерево вузлів…», «Застосувати позначені» и «Зробити еталоном для СС»: они пишут узлы, спецификации или СС в рабочую базу сразу, без записи формы и вне транзакции теста — эти команды проверены COM-тестами 2B/2C с откатом.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/load_lists/e2d.txt _Rarzrabotki/Python/test/spec_domov/out/e2d_load_*.txt \
  _Rarzrabotki/Python/test/spec_domov/sync_manifest.json
git commit -m "feat(BASEBuh): партия 2D загружена — формы этапа 2 в bas_industrialbud

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 44: Приёмка этапа 2 целиком (§10.1, этап 2)

**Files:**
- Create: `_Rarzrabotki/Python/test/spec_domov/test_e2_priemka.py`
- Create (результат): `_Rarzrabotki/Python/test/spec_domov/out/e2_priemka_<stamp>.json`

- [ ] **Step 1: Написать приёмочный прогон**

`_Rarzrabotki/Python/test/spec_domov/test_e2_priemka.py`:

```python
# -*- coding: utf-8 -*-
"""Приёмка этапа 2 (§10.1) одним прогоном. Каждый тест — отдельным процессом (своё COM-соединение, свои транзакции с откатом):
30 м — 170 позиций, 17 зелёных, жёлтые только на ОН с одной позицией (test_e2a_zapolnit); мембрана, плінтус, труба, кутник,
септик 1/2/1/1 и Σ 12 000,00, Σ норм 1 009 864,74 ± 1,00 (test_e2b_ceny); узлы и узел в узле (test_e2b_uzly); 18 из 22 нулевых
с фактом, 4 алиаса (test_e2b_nulevye); 1 сирота — зелёное (test_e2c_perepriv); 15 м копией — жёлтые расхождения с фактом,
k = 1 у поз. 27/44/45, Σ 704 612,01, И4 0,00, «Не постачається» не мешает «Діє» (test_e2c_kopiya); формы (test_e2d_formy);
регресс этапа 1 (test_e1_priemka). Затем: регистрация СС к обмену в конце = в начале прогона, элементов справочника не
прибавилось."""
import datetime
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, OUT, connect, q  # noqa: E402
from test_e1_priemka import START, exchange_snapshot, k10_exchange_untouched  # noqa: E402

TESTS = ["test_e2_zaprosy.py", "test_e2a_edinicy.py", "test_e2a_zapolnit.py", "test_e2b_uzly.py", "test_e2b_nulevye.py",
         "test_e2b_ceny.py", "test_e2c_perepriv.py", "test_e2c_kopiya.py", "test_e2d_formy.py", "test_e1_priemka.py"]


def spec_refs(B):
    rows = q(B, """ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец""")
    return sorted(B.String(r["Ссылка"].УникальныйИдентификатор()) for r in rows)


def run(name):
    p = subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, name)], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=3600)
    tail = (p.stdout + p.stderr).strip().splitlines()[-4:]
    return p.returncode == 0, tail


def main():
    B = connect()
    before = spec_refs(B)
    START["exchange"] = exchange_snapshot(B)
    ok, fails = {}, {}
    for name in TESTS:
        passed, tail = run(name)
        (ok if passed else fails)[name] = tail
        print("OK  " if passed else "FAIL", name, "|", " / ".join(tail))
    try:
        k10_exchange_untouched(B)
        ok["exchange_untouched"] = "регистрация СС к обмену в конце прогона = в начале"
    except AssertionError as e:
        fails["exchange_untouched"] = repr(e)[:500]
    after = spec_refs(B)
    if after != before:
        fails["no_data_written"] = f"элементов справочника было {len(before)}, стало {len(after)} — тест записал данные мимо транзакции"
    else:
        ok["no_data_written"] = f"элементов справочника {len(after)}, как до прогона"
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e2_priemka_{datetime.datetime.now():%Y%m%d_%H%M}.json")
    json.dump({"ok": ok, "fail": fails}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("->", out, "ПРОВАЛОВ:", len(fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Запустить — до загрузок 2A–2D падает**

Ожидаемое поведение до загрузок уже подтверждено задачами 29–43 (`AttributeError` в каждом тесте этапа 2). Повторно запускать не нужно.

- [ ] **Step 3: Реализация** — задачи 29–43.

- [ ] **Step 4: Запустить — PASS**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/test_e2_priemka.py`
Expected: 10 строк `OK  test_…`, `ПРОВАЛОВ: 0`, файл `out/e2_priemka_<stamp>.json`. Если число на день приёмки отличается от §10.1 (база живая: новые ПТУ, списания, правки СС после 03.10.2026 23:41) — не подгонять тест: найти документы-причины запросом из `test_e2_zaprosy.py` (факт и склад сверяются со снимком этапа 0), записать объяснение в `out/e2_priemka_<stamp>.json` поле `"объяснения"` и в отчёт пользователю.

- [ ] **Step 5: Коммит**

```bash
git add _Rarzrabotki/Python/test/spec_domov/test_e2_priemka.py _Rarzrabotki/Python/test/spec_domov/out/e2_priemka_*.json
git commit -m "test(spec_domov): приёмка этапа 2 П1 пройдена — эталоны 30 м и 15 м в транзакции, обмен не тронут

30 м: 170 позиций, 17 зелёных, Σ 1 009 864,74; 15 м копией: Σ 704 612,01, И4 0,00; нулевые 18 из 22; перепривязка Т10.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 45: Завершение П1 этапов 0–2 — сверка основного каталога, память, git

**Files:**
- Create: `C:/Users/support1c/.claude/projects/C--Configuration-downloads-BASERP25/memory/p1_specifikaciya_etalon_buhbud.md`
- Modify: `C:/Users/support1c/.claude/projects/C--Configuration-downloads-BASERP25/memory/MEMORY.md` (одна строка в разделе «Производство и домики»)

- [ ] **Step 1: Проверка Rule #4 — каждый загруженный файл в основном каталоге равен worktree**

Run:
```bash
for L in e1a e1b e2a e2b e2c e2d; do python -X utf8 _Rarzrabotki/Python/test/spec_domov/to_main.py --verify _Rarzrabotki/Python/test/spec_domov/load_lists/$L.txt || echo "FAIL $L"; done
```
Expected: шесть строк `файлов N, расхождений 0`, ни одного `FAIL`.

- [ ] **Step 2: Проверка дампа против базы по всем объектам П1**

Run: `python -X utf8 _Rarzrabotki/Python/test/spec_domov/dump_check.py --tag final --canon --objects "ОбщийМодуль.А_СпецификацииДомовСервер,Справочник.А_СпецификацииДомов,Справочник.А_СпецификацииДомов.Форма.ФормаЭлемента,Справочник.А_СпецификацииДомов.Форма.ФормаСписка,Справочник.А_СпецификацииДомов.Форма.ФормаГруппы,Справочник.А_СпецификацииДомов.Форма.ФормаПодбораКарточки,Справочник.А_СпецификацииДомов.Форма.ФормаПерепривязки,Справочник.СтруктураСебестоимости,Справочник.СтруктураСебестоимости.Форма.ФормаЭлемента,Документ.РасчетКомплектаций,Перечисление.А_ВидыСпецификацийДомов,Перечисление.А_СтатусыСпецификацийДомов,Перечисление.А_ВидыСтрокСпецификацииДомов,Перечисление.А_ИсточникиКарточкиСпецификации,Перечисление.А_ГруппыЗакупочнойСпецификации,Перечисление.А_ИсточникиЦеныЗакупки"`
Expected: `различие по содержимому 0`.

- [ ] **Step 3: Записать память**

`C:/Users/support1c/.claude/projects/C--Configuration-downloads-BASERP25/memory/p1_specifikaciya_etalon_buhbud.md`:

```markdown
---
name: П1 «Специфікація і еталон» (BuhBud) — этапы 0–2 загружены
description: Справочник А_СпецификацииДомов (дом/узел/закупочная), 6 перечислений, СС.ЭталоннаяСпецификация, общий модуль А_СпецификацииДомовСервер; инструменты и приёмка — _Rarzrabotki/Python/test/spec_domov
type: project
---
Этапы 0–2 спецификации docs/superpowers/specs/2026-10-03-mnogourovnevaya-specifikaciya-domov-design.md загружены в bas_industrialbud
(6 Partial-загрузок 1A, 1B, 2A, 2B, 2C, 2D). План: docs/superpowers/plans/2026-10-04-p1-specifikaciya-i-etalon.md.

**How to apply:**
- Общий модуль собирается build_bsl.py из spec_domov/bsl/*.bsl; тексты запросов — spec_domov/zaprosy/*.sql (строка //@ЗАПРОС).
  Править исходники, не Module.bsl. Формы — gen_forms.py --etap 2 (form-compile для BASEBuh не годится).
- Загрузка: to_main.py (Rule #4 с манифестом хешей, Д1) → load_stage.ps1 -List <список> -Tag <метка>.
- Тесты пишут только в транзакции COM с откатом (fixtures.py, префикс «__ТЕСТ_П1 »); эталоны в базу не записаны —
  их делает бухгалтер (30 м ≈ 2 дня, 15 м копией из 30 м).
- Толкования: Т9 (деталь без карточки — только в «Чернетці»), Т10 (зелёное «єдина сирота»), Т11 (узел-позиция по тексту ≤ 2 правок).
- common.py указывает на дамп worktree budget-calc-oct-dec-2026-20c998: при удалении worktree поменять WT/BB_WT на новый.
- Дальше: П1 этапы 3+ (закупочная) — отдельный план; РК и регистры узлов — П2+.
```

В `MEMORY.md`, раздел «## Производство и домики», последней строкой раздела:

```markdown
- **[П1 «Специфікація і еталон»: справочник А_СпецификацииДомов, этапы 0–2 загружены; spec_domov — сборка модуля, формы, загрузка, приёмка в транзакции](p1_specifikaciya_etalon_buhbud.md)**
```

- [ ] **Step 4: Коммит в ветке worktree**

```bash
git add docs/superpowers/plans/2026-10-04-p1-specifikaciya-i-etalon.md
git status --short -- _Rarzrabotki/Python/test/spec_domov _Rarzrabotki/BASEBuh
git commit -m "docs(spec_domov): П1 этапы 0–2 завершены — план и итоговая сверка основного каталога

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```
Expected: `git status --short` перед коммитом не показывает несохранённых файлов П1 (все закоммичены в задачах 1–44); `ConfigDumpInfo.xml` и чужие `Documents/А_ПредварительныйБюджет*` не добавлены.

- [ ] **Step 5: Перенос в `claude/main` и на GitHub (память `feedback_autonomous_finish_merge_push_pr`, `git_push_large_files_cherry_pick`, `feedback_merge_prep_only_branch_files`)**

Файлы дампа уже лежат в основном каталоге (Rule #4, шаг 1). Коммиты П1 ветки:

```bash
git log --reverse --format=%H main..HEAD -- _Rarzrabotki/Python/test/spec_domov _Rarzrabotki/BASEBuh docs/superpowers/plans/2026-10-04-p1-specifikaciya-i-etalon.md > "C:/Users/SUPPOR~1/AppData/Local/Temp/claude/plan_p1/p1_commits.txt"
git diff --name-only $(git merge-base main HEAD) HEAD -- _Rarzrabotki/Python/test/spec_domov _Rarzrabotki/BASEBuh docs/superpowers/plans/2026-10-04-p1-specifikaciya-i-etalon.md > "C:/Users/SUPPOR~1/AppData/Local/Temp/claude/plan_p1/p1_files.txt"
git -C C:/Configuration_downloads/BASERP25 branch --show-current
tr '\n' '\0' < "C:/Users/SUPPOR~1/AppData/Local/Temp/claude/plan_p1/p1_files.txt" | xargs -0 git -C C:/Configuration_downloads/BASERP25 status --short --
```
Если основной каталог на `main`/`claude/main` и `status` по этим путям показывает только файлы П1, побайтно равные blob ветки (сравнивать с `git show HEAD:<путь>`, не с рабочей копией — память о CRLF), — закоммитить их в основном каталоге одним коммитом с сообщением `feat(BASEBuh): П1 «Специфікація і еталон» этапы 0–2 (из claude/budget-calc-oct-dec-2026-20c998)`; если по этим путям есть чужие правки — коммит в основном каталоге пропустить (Rule #4 выполнен копиями, историю несёт ветка worktree) и сказать об этом пользователю.
GitHub: локальный `main` непушабелен (файлы > 100 МБ в истории) — временный sparse-worktree от `origin/main`, cherry-pick коммитов из `p1_commits.txt`, ветка `claude/p1-specifikaciya-etalon`:

```bash
git fetch origin main
SHA=$(git rev-parse origin/main)
git worktree add --no-checkout C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1 $SHA
git -C C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1 sparse-checkout set _Rarzrabotki/Python/test/spec_domov _Rarzrabotki/BASEBuh docs/superpowers/plans
git -C C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1 reset --hard $SHA
git -C C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1 switch -c claude/p1-specifikaciya-etalon
git -C C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1 cherry-pick $(cat "C:/Users/SUPPOR~1/AppData/Local/Temp/claude/plan_p1/p1_commits.txt")
git -C C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1 push -u origin claude/p1-specifikaciya-etalon
git worktree remove --force C:/Configuration_downloads/BASERP25/.claude/worktrees/tmp-push-p1
```
Если cherry-pick не ложится (origin/main отстал по тем же файлам) — `git cherry-pick --abort` и вариант «снапшот» из памяти: скопировать конечные файлы из `p1_files.txt`, кроме `Configuration.xml`; `Configuration.xml` взять `git show origin/main:_Rarzrabotki/BASEBuh/Configuration.xml` и вписать точечно только 8 объектов П1 (справочник, 6 перечислений, общий модуль). PR пользователь создаёт по ссылке `https://github.com/Alex16111977/BASERP25_INDUSTRIALBUD/pull/new/claude/p1-specifikaciya-etalon`.
Итог пользователю в чате (по-русски): этапы 0–2 загружены, приёмка `out/e2_priemka_<stamp>.json`, ссылка на ветку, что дальше делает бухгалтер (эталон 30 м по памятке §5).


---

## Самопроверка

### 1. Покрытие спецификации (требование → задачи → чем проверено)

| Требование | Задачи | Проверка |
|---|---|---|
| §10.1 этап 0: снимок контрольных сумм (факт домов 30 м и 15 м, склад проекта, закуплено, прочие приходы и расходы) против [Z6] | 6, 7 | `out/e0_snimok_*.json`, флаги `ok` по каждой сумме, расхождения — списком документов после 03.10.2026 23:41 |
| §10.1 этап 0: дамп против живой конфигурации, планы обмена, роли | 2, 4, 8 | `out/e0_kontrol_*.json`: `main_eq_db`, `dump_eq_main`, `rk_not_in_exchange`, `users_by_roles`, `roles_write_ss`, `ss003_214` |
| §4.1–4.2 справочник с папками, 3 вида, шапка из 12 реквизитов | 12, 13, 15, 17 | `test_e1_metadata.py` (`HEADER`, `Hierarchical`, синоним) |
| §4.3 ТЧ `Состав`; §4.4 `Дома`, `СоставОснования`, копия с ключами | 13, 23 | `TS` в `test_e1_metadata.py`; `ПриКопированииСпецификации` |
| §4.5 ТЧ закупочной (метаданные; при записи — дом факта своего типа и без повторов, карточка уникальна) | 13, 23 | `TS["ДомаФакта" … "ФактПоДомам"]`; k9 (дом 15 м в закупочной 30 м — отказ) |
| §4.6 проверки при записи и при «Діє» (цикл, глубина, вид строки, «Перевірено», k «Не знайдено», позиция в СС, узлы в «Діє», закупочная) | 22, 23, 25 | `test_e1_priemka.py` k1–k9 |
| §4.7 `СС.ЭталоннаяСпецификация`, «Зробити еталоном для СС» без `ПередЗаписью` СС и без регистрации обмена (Р3, Б1), поле только для просмотра | 14, 24, 27 | k6 (СС 003: 214 из 214, строки `Изменения` не изменились), k7, `check_ssform` |
| §4.7 РК: `СпецификацияДома`, `ДатаНормы`, ТЧ `НормаДома`, `ЭкземплярыУзлов`, колонки `ПозицияСС`/`НеИзНормы`, синоним; РК непроводимый | 14 | `RK_TS`, `Проведение = Запретить` в `test_e1_metadata.py` |
| §4.10 шесть перечислений с uk-синонимами | 11, 12 | `UK_ENUM` |
| §4.11 общий модуль (Сервер + ВнешнееСоединение), права Д2, новые объекты вне планов обмена | 9, 15, 25, 44 | `ROLE_RIGHTS`, `ExchangePlans.Content`, `k10_exchange_untouched` |
| Р1 только BuhBud | все | `common.CONN` — только `bas_industrialbud`, ERP не подключается ни одним скриптом |
| Р5, Т6 узел в узле, дом — уровень 0, узлы 1–9 | 21, 22, 34 | k3–k5; `test_e2b_uzly.py` («Зовнішні мережі» ⊃ «Септик», входимость 1/2) |
| Р6 единица строки, k «1 од. картки = k од. рядка» | 20, 30, 41 | `test_e1_funkcii.py`, `test_e2a_edinicy.py`, `test_e2b_ceny.py` (мембрана k = 75) |
| Р8 без округлений в разузловании | 21 | `t_unfold` |
| Р22 «Нульові позиції з фактом», печать для экономиста | 35, 41 | `test_e2b_nulevye.py`: 22 / 18, +4 алиаса |
| Р23 синонимы единиц, подсказка k из названия; Д3 «паков» не сводится | 20, 30 | `test_e2a_edinicy.py`; `test_e2c_kopiya.py` (поз. 27, 44, 45: k = 1, «Одиниці збігаються») |
| Р24 «Не постачається» с причиной, не мешает «Діє», входит в И4 | 23, 36, 39 | k8; `test_e2b_ceny.py`; `t_diye_ne_postach` |
| Р25 жёлтая карточка только на ОН с одной ненулевой позицией и при непротиворечивом размере | 30, 32 | `test_e2a_zapolnit.py` (170 / 17 зелёных) |
| Р26 площадной узел; допуск сравнения 0,00001 | 21, 36 | `t_ploshchadnoy`; `СравнитьСоСтруктуройСебестоимости` |
| Р27 «Додано в будинку» в варианте | 23, 36 | проверка варианта при записи; список `ДобавленоВДоме` |
| Р31 «Склад з документа комплектації» без деления на выпуск | 31, 34, 41 | `t_kn`; `test_e2b_uzly.py` и `test_e2b_ceny.py`: ІБ00-000174 → 1/2/1/1, Σ 12 000,00 |
| Р32 + Т11 копия карточек и дерева узлов | 39, 41 | `test_e2c_kopiya.py`: жёлтые сверены с фактом 15 м независимо, Септик в поз. 168, «Водостік 15 м», Σ 704 612,01, И4 0,00 |
| Р35 «Позначити зелені перевіреними», подбор с фактом, складом, ценой, размером | 32, 41, 42 | `t_fill`, `t_podbor`, `test_e2d_formy.py` |
| Т9 деталь без карточки — только в «Чернетці» | 23 | k8 |
| Т10 зелёное «єдина сирота й єдина нова позиція етапу» | 38 | `test_e2c_perepriv.py` |
| §5 шаги 0–2 и 7 (работа бухгалтера) | 32, 34–36, 38, 39, 41–43 | команды формы (Task 41–42), ручная проверка в Task 43 |
| §6 разузлование до карточек; ціни норми (п. 6, И4 до копейки) | 21, 36 | `t_unfold`; `test_e2b_ceny.py` (|И4| ≤ 1,00), `test_e2c_kopiya.py` (И4 = 0,00) |
| §9 узлы строителя (21 имя) | 34 | `ИменаУзловСтроителя` в `test_e2b_uzly.py` |
| §10.1 приёмка этапа 1 | 19, 28 | `out/e1_priemka_*.json`, `ПРОВАЛОВ: 0` |
| §10.1 приёмка этапа 2 (все числа строки «2») | 31, 32, 34–36, 38, 39, 44 | `out/e2_priemka_*.json`, `ПРОВАЛОВ: 0` |
| §10.1 общие правила: Rule #-1 (запрос сперва в Python COM) | 18, 29, 20 | `zaprosy/*.sql` исполняются тестом, затем `build_bsl.py` вставляет тот же файл (`//@ЗАПРОС`) |
| §10.1 общие правила: только Partial, аудит `Configuration.xml`, `/CheckModules`, Rule #4 | 2–5, 17, 28, 33, 37, 40, 43, 45 | `load_stage.ps1` (гейты, лог, CheckModules, db-update), `to_main.py --verify` |
| Данные (эталоны, закупочные) в рабочую базу не пишутся | 1, 16, 19, 44 | `tx()` всегда откатывает; `check_static.py --tests` запрещает `CommitTransaction` и `.Удалить()`; `no_data_written` в Task 44 |

**Вне этого плана (по постановке — этапы 0–2):** §4.8–4.9 (документ `А_ФормированиеУзловДома`, регистры узлов) — этап 6; логика закупочной §7 — этап 3; РК в режиме спецификации — этап 5; варианты (Р4, Р19, Р20, Р33) — этап 4; отчёты §8.3–8.4 — этапы 7–8. Нужные этапам 0–2 метаданные этих частей (ТЧ закупочной, ТЧ РК) созданы в Task 13–14.

**Сознательное отличие от буквы §4.6:** «другая действующая закупочная той же СС переводится в «Закрита» с подтверждением» — на этапе 1 проверка «Діє» отказывает с текстом «сначала закройте её» (Task 23); закрытие с подтверждением — команда формы закупочной на этапе 3, где появляется её логика. Приёмка этапа 1 этого не требует.

**Отложено с явной пометкой (не входит в приёмку §10.1 этапов 0–2):**
- Р25, вторая половина: подсказка жёлтой строки «ЗН стоїть на N позиціях СС; ця картка вже на поз. …». По правилу Р25 жёлтая карточка ставится только на ЗН с одной ненулевой позицией, поэтому сразу после «Заповнити з СС» подсказка всегда «1 позиція, ніде більше»; смысл она получает после ручных замен. Под неё нет колонки в §4.3 (новые реквизиты — правка метаданных), поэтому в П1 её роль играет подбор Р35 (колонки факта по домам и «розмір з назви»). Сделать в П3 вместе с формой варианта или отдельным решением пользователя.
- Т2 «Дт 20/22/28»: `ceny_posledney_zakupki` берёт строки `Товары` ПТУ на склады подразделения закупки без отбора по счёту учёта. В П1 цена нужна только как вес ціни норми и как колонка подбора; для П2 (докупка по Р39, защита «×5») запрос уточняется отбором по `СчетУчета` 20/22/28 с новым Python COM тестом.

### 2. Скан плейсхолдеров

- Поиск по тексту плана: `TBD`, `TODO`, `implement later`, `fill in`, «аналогично Task», «как в Task N», «по аналогии», «добавить обработку ошибок» — совпадений нет. Шаги 4, которые ссылались на «статику как в Task 21» или только на будущую загрузку, получили явные команды `check_static.py` с ожидаемым выводом (задачи 22–24, 30–32, 34–36, 38, 39, 41).
- Многоточие «…» встречается только в путях таблиц («`…/Forms/ФормаСписка.xml`»), в ожидаемом выводе («`скопирован …`») и в пользовательских строках («Інша назва…», «Перепривʼязати позиції СС…») — не в коде.
- Тексты трёх запросов «дословно из заметок» (`kandidaty_kartochek`, `dokumenty_komplektacii`, `sostav_dokumenta_komplektacii`) не перепечатаны, а копируются скриптом Task 29 из `_p1_notes/E3_tochki_vstraivaniya.md` (закоммичен в Task 1) по точным диапазонам строк с `assert` на начало и конец; `nulevye_pozicii` перепечатан в план целиком (источник в Temp мог исчезнуть). Все четыре затем исполняются Python COM тестом до вставки в BSL.
- Каждая задача кода и метаданных содержит: Files, падающий тест (полный код или тест своей фазы), команду FAIL с ожидаемым текстом, реализацию полным кодом, команду PASS, коммит с перечнем файлов и `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`; задачи загрузки (17, 28, 33, 37, 40, 43), приёмки (44) и завершения (45) — проверочные команды до и после с ожидаемым выводом. Всего 45 коммитов на 45 задач, `git add -A` нигде нет.


### 3. Согласованность имён, типов и сигнатур (проверено скриптом по тексту плана)

- 22 исходника BSL: повторов процедур и функций нет ни в общем модуле (15 частей), ни в модулях форм (элемент = `forma_elementa_1` + `forma_elementa_2`, список = `forma_spiska_1` + `forma_spiska_2`).
- Каждый вызов `А_СпецификацииДомовСервер.X(` в модулях форм и объекта и каждый `s.X(` в Python-тестах ссылается на экспортную функцию частей `srv_*.bsl`; внутренние вызовы частей — на функции, определённые в модуле (`ВидыСубконтоНоменклатураСклад`, `НаименованиеЕдиницы`, `ЭтоЦифра`, `ЭтоБуква`, `ТекстЗапроса…`).
- Python: 39 файлов разбираются `ast`; все `from common|fixtures|test_e1_priemka|test_e2b_uzly|to_main|build_bsl import …` указывают на существующие имена.
- Колонки: `PODBOR_COLS` (Task 26) = колонки `ПредложитьКарточки` (Task 32); `PEREPRIV_COLS` = `НоваяТаблицаПредложений` (Task 38); `SR_COLS` ⊂ `СравнитьСоСтруктуройСебестоимости().Позиции` (Task 36); `RAZV_COLS` ⊂ `НоваяТаблицаЛистьев` (Task 21); `TREE_COLS` — поля `ДобавитьВетку`.
- Перечисления: значения в BSL и тестах ⊂ `UK_ENUM` (Task 9); существующее `А_ИсточникиКоэффициентаПересчетаСС` — `ЕдиницыСовпадают`, `Вручную`, `НеНайден` (сверено с основным дампом).
- Реквизиты: используемые реквизиты справочника и колонки `Состав` ⊂ `HEADER`/`TS` (Task 9); у СС — `ПодразделениеОбъекта`, `ПодразделениеЗакупки`, `Комплектующие.ЕдиницаСС/Единица/КоличествоСС_Оригинал/ЕдиницаСС_Оригинал` (сверено с основным дампом); `А_СтруктураСебестоимостиСервер.КлючПозиции/КлючОригинала/КлючЕдиницы` существуют, модуль — Сервер + ВнешнееСоединение.
- Ключевые сигнатуры, общие для нескольких задач: `РазвернутьДоКарточек(Корень, ВключатьНеПоставляется = Ложь)` → `Структура(Листья, Экземпляры, ЯвныеПозиции)`; `ЗагрузитьДеревоСпецификаций(Корень)` → `Структура(Кэш, КлючКорня)`; `ПозицииСтруктурыСебестоимости(СсылкаСС)` → `Соответствие` ключ → `Структура(НомерСтроки, Этап, НоменклатураСС, ОбщееНазвание, Количество, ЕдиницаСС, КоличествоОригинал, ЕдиницаОригинал, Сумма, Единица)`; `Входимость(Узел)` → ТЗ с `СтруктураСебестоимости`; `СоставИзДокументаКомплектации(Док, УзловВДокументе = 1)` → `Строки` с колонкой `Количество` (уже на узел); `ОбъединитьВУзел(СпецОбъект, ИндексыСтрок, ИмяУзла, КоличествоВыпуска, ЕдиницаВыпуска)` → ссылка; `ПоставитьУзелВСтроку(СпецОбъект, ИндексСтроки | Неопределено, Узел)`; `РассчитатьЦеныНормы(ЭталонОбъект, НаДату = Неопределено)` → `Структура(СуммаСС, СуммаНормы, И4, Позиций, БезЛистьев, ПропущеноУзлов)`; `СкопироватьИзДругойСпецификации(СпецОбъект, Источник, СуффиксТипа = "")` → `Структура(Совпало, Отличий, НеЗакрыто, НеПоставляетсяСкопировано, УзловПереиспользовано, УзловСкопировано, Отличия)` — формы и тесты используют ровно эти поля.
- Партии `build_bsl.py`: 6 / 9 / 13 / 15 частей = ожидаемые «частей N» в задачах 25, 33, 37, 40. Списки загрузки задач 17, 28, 33, 37, 40, 43 совпадают с разделом «Списки Partial-загрузки».

### 4. Исправлено при самопроверке

1. `to_main.py` не копировал `Ext/Form.xml` и модули форм (в списках — только корень формы): основной каталог получил бы формы без содержимого в загрузках 1B и 2D. Добавлено `expand()` — корень формы тянет `Ext`, но только новые формы П1 и синхронизированные файлы (устаревший модуль формы СС из worktree не попадёт); тест `test_expand_forms`.
2. `e0_kontrol.py`: код СС — параметром `&Код`; роли пользователя — перебором (у `КоллекцияРолейПользователя` нет `Получить()`).
3. `test_common.py`: проверка сообщений без тавтологии.
4. Task 17: в коммит — шесть файлов перечислений поимённо, не каталог `Enums`.
5. `НайтиЦикл` — выход при отсутствии корня в кэше; `ПроверитьДомаВариантов` — выход при пустой СС.
6. `ПредложенияПерепривязки`: нулевые позиции не считаются «свободными» — иначе нулевая позиция в этапе септика превращала зелёное правило Т10 в жёлтое.
7. Копия карточек сохраняет ОН строки (ОН позиции), сверка с фактом идёт по ней; тест сверяет тем же правилом.
8. Статусы «Порівняти з СС» в тестах сравниваются с обоими языковыми вариантами (`{"Збіг", "Совпадает"}`; `B.НСтр` через COM недоступен), источники и вид строки — через представления значений перечислений. Тексты ошибок в негативных кейсах (`входить сам у себе`, `не перевірено`, …) — украинские: язык сеанса `cfo` в `bas_industrialbud` — «Українська» (память `uk_lang_sinonim_pustoy`).
9. `test_e2b_nulevye.py`: «закрыто алиасами» — разность до и после четырёх алиасов (карточки, уже стоящие после «Заповнити з СС», не ломают проверку).
10. Узел из нескольких позиций при копии сравнивается по количествам × количество строки-узла источника, масштаб копии = выпуск / количество строки; «Обʼєднати у вузол» и копия используют одну `ЗаменитьСтрокиУзлом`.

### 5. Исправлено при проверке плана (ревью 04.10.2026)

Проверено: покрытие этапов 0–2 и Р*, плейсхолдеры, имена против §4, XML против образцов основного дампа (реквизит шапки СС, реквизит и ТЧ РК со стандартным `LineNumber`, `ChoiceParameters` по `Контрагенты.ГоловнойКонтрагент`, `Rights.xml`, подсистема, метаданные формы, порядок `ChildObjects`), BSL — статическим гейтом плана по всем 22 исходникам и отдельной сверкой неявных локальных (присваивания, переменные циклов, параметры) с 755 именами общих модулей и свойствами глобального контекста, входные данные тестов — запросами `ВЫБРАТЬ` к `bas_industrialbud` (единицы 2938/2009/0101/0123/0306/2110, 14 карточек, СС 009 — 170 строк, Σ 1 009 864,74, 9 строк «Гвинтові кріплення»; СС 008 — 190 строк, 22 нулевые, Σ 704 612,01, 8 строк группы; ключи позиций водостока 30 м 67–74 = 15 м 66–73, септик 30 м 170 и 15 м 168 — один этап; ІБ00-000174 — выпуск 0,5, не проведён; дома 7 и 6). Исправлено:

1. **BSL, рантайм-ошибка:** в `ЗаполнитьИзСС` (Task 32) неявная локальная `Документы` перекрывала менеджер `Документы` глобального контекста → «Поле объекта недоступно для записи». Переименована в `ДокументыКН`. `check_static.py` (Task 16) теперь ловит этот класс: свойства глобального контекста в `RESERVED` и переменные циклов `Для Каждого X Из` / `Для X =`; `test_check_static.py` проверяет это на примере.
2. **Имена против §4.11:** `ФактСписанийДомов` → `ФактСписанийДома`, `ПрименитьПерепривязку` → `ПерепривязатьПозицииСС` (BSL, формы, тесты).
3. **§4.5 не было проверки** «дом факта — дочерний к `СС.ПодразделениеОбъекта`» и пустого дома в `ДомаФакта`: добавлена в `ПроверитьЗакупочнуюПриЗаписи` (Task 23); k9 проверяет отказ на доме 15 м в закупочной 30 м.
4. **`to_main.py` (Task 4):** файлы списка, равные основному каталогу (корни ролей в e1a), давали «ОТКАЗ … не в манифесте» и роняли Task 17 — теперь «без изменений»; тест `test_refuse_unknown_existing` ждал «нет в worktree», хотя `Номенклатура.xml` в worktree есть, — ждёт «не в манифесте»; добавлен `test_skip_unchanged`.
5. **`dump_check.py` (Task 4):** XML сравнивается по C14N (порядок `xmlns` и атрибутов не различие); новые файлы П1, которые платформа сериализовала по-своему, при `--canon` принимаются каноном с выводом начала различия, строгая проверка осталась для существующих объектов. `gen_forms.py` пишет `xmlns` в каноническом порядке (по алфавиту, как `А_ТехкартаКМД/Forms`) — раньше каждая форма давала бы «различие по содержимому» в Task 28/43/45.
6. **Rule #4 в конце (Task 45):** `gen_forms.py --etap 2` больше не перегенерирует `ФормаГруппы`, `build_bsl.py` пишет модуль объекта только в партии 1 — иначе их worktree-копии расходились с каноном основного каталога и `to_main.py --verify e1b` падал.
7. **Task 29:** диапазон текста `kandidaty_kartochek` в E3 — строки 150–340 (149 и 341 — ограждения блока, `assert startswith("ВЫБРАТЬ")` падал); `nulevye_pozicii` перепечатан в план (был путь в Temp); заметки E1–E3 коммитятся в Task 1.
8. **COM:** `B.НСтр(...)` (глобальная функция языка) через COM недоступна — статусы сравниваются с обоими вариантами; даты «на сейчас» в `test_e2_zaprosy.py` — дальней датой, без переноса `ТекущаяДатаСеанса()` через COM (−3 ч, память `com_date_shift_3_hours`).
9. **Инварианты вместо снимка:** k10 и Task 44 сравнивают регистрацию СС к обмену в начале и в конце своего прогона, а не со снимком этапа 0 (между днями её меняют обмен и пользователи); `test_e2c_kopiya.py` ищет «Водостік 15 м» по `Основание` теста — боевые узлы бухгалтера проверку не ломают.
10. **Сводки загрузок:** `load_stage.ps1` сам пишет `out/<Tag>_load_<stamp>.txt` (раньше Task 28 и 43 коммитили файлы, которые ни один шаг не создавал); сводки 2A–2C добавлены в коммиты.
11. **Безопасность ручной проверки (Task 43):** перечислены команды формы, которые пишут узлы/СС сразу (их нельзя нажимать при ручной проверке — они проверены COM-тестами с откатом).

### 6. Итог исполнения (04.10.2026)

- Загрузки 1A, 1B, 2A, 2B, 2C, 2D выполнены (`out/e1a_load_*` … `out/e2d_load_*`).
- Приёмка этапа 2 (Task 44): `out/e2_priemka_20261004_1031.json` — 10 тестов OK, `ПРОВАЛОВ: 0`, регистрация СС к обмену не изменилась, элементов справочника не прибавилось; регресс этапа 1 — `out/e1_priemka_20261004_1031.json`, `ПРОВАЛОВ: 0`.
- Rule #4 (Task 45 шаг 1): `to_main.py --verify` по e1a/e1b/e2a/e2b/e2c/e2d — файлов 18/15/1/1/1/13, расхождений 0.
- Дамп против базы (Task 45 шаг 2): `dump_check.py --tag final --canon` по 16 объектам П1 — совпало 39, только форматирование 0, различие по содержимому 0.
- Основной каталог: 33 файла дампа П1 равны blob ветки (сравнение через `git hash-object --path`, autocrlf), но `Configuration.xml`, `РасчетКомплектаций.xml` и роли в основном каталоге несут чужие незакоммиченные правки (отчёт дебиторки, обработка ПДВ, удалённый РС `А_КарткаЗаЗагальноюНазвою`) — коммит в основном каталоге пропущен, историю несёт ветка worktree.
