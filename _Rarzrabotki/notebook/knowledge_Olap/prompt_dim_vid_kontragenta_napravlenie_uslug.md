# PROMPT: Виды контрагентов + Направление оказания услуг → OLAP (новые Dim_) и срезы баланса PL.pbix

> ## ⚡ СТАТУС 2026-06-12: Phases 0–4 ВЫПОЛНЕНЫ. Осталось Phase 5 (PL.pbix, нужен открытый Power BI Desktop) + Phase 6 (частично).
> - Phase 0: строка 267 удалена пользователем (конфигуратор+F7); разметка APPLY: Внутригрупповые 331, Внутренние 1106, Собственники 10 (+ пользователь сам: Внешние 7162, Кредиторы 17; НаправлениеУслуг 1104).
> - Phase 1: WHITELIST + `refresh_mapping` OK (аддитивно). Маппинг: справочник `_Reference56330`; договор `_Fld56332RRef` (вид), `_Fld56331RRef` (направление услуг).
> - Phase 2: `scripts/ddl_dim_vidy_kontragentov.sql` применён (Dim_VidyKontragentov char(32) PK + 2 колонки Dim_Contracts).
> - Phase 3: `dim_catalogs.json` — новый step `dim_vidy_kontragentov` + расширен raw_sql `dim_contracts`; полный `python main.py` Success (fact_balance 45653, run_id 446).
> - Phase 4: `scripts/verify_olap_dim_vid_kontragenta.py` **PASS** — раскладка SQL==1С, FK orphans 0, кросс-сверка «Внутригрупповые» 2026-01 РС 0,12 == SQL 0,12, полный баланс Σ Sum_Close 2026-01 = 0,00 (Актив теперь 338 857 601,40 — живая база дрейфует от канона 288M, Σ=0 держится); pytest Глобино-2 PASS.
> - Phase 5: НЕ выполнена — Power BI Desktop не был запущен (ListLocalInstances = 0).
>
> **Самодостаточный промт для свежей сессии.** Дата создания: 2026-06-12.
> Дизайн согласован с финансистом: НЕ создавать новые статьи ПВХ.СтатьиАктивовПассивов,
> НЕ менять 1С-код и структуру РС `А_ОтчетБаланс_Свод` (OD-9). Расшифровка баланса
> по видам контрагентов и направлению услуг идёт через измерение `Договор`,
> которое свод уже хранит, → snowflake от `Dim_Contracts` в OLAP.

---

## Контекст

| | |
|---|---|
| ERP база | `Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"` (COM `V83.COMConnector`) |
| SQL OLAP | `localhost`, БД `OlapBASERP`, login `sa` (пароль — memory `sql_olap_baserp_credentials`) |
| ETL | `_Rarzrabotki/Olap/Ai_Olap/` — `python main.py` (СТРОГО venv-python; дефолт = все Dim + fact_balance) |
| Mapping | `_Rarzrabotki/Olap/Ai_Olap/mapping/refresh_mapping.py` — строгий WHITELIST |
| PBIX | PL.pbix через MCP `powerbi-modeling-mcp`; user-gated: Desktop Refresh + Ctrl+S |
| Базы знаний | `_Rarzrabotki/notebook/knowledge_Olap/` (ETL/SQL/PBIX), `knowledge_Balanse/` (свод), `knowledge_Balanse_klient/` (расчёты с клиентами) |

**Исходные реквизиты (уже в конфигурации, метаданные НЕ менять):**
- `Справочник.ДоговорыКонтрагентов.А_ВидКонтрагента` → `Справочник.А_ВидыКонтрагентовДляБаланса`
- `Справочник.ДоговорыКонтрагентов.А_НаправлениеОказаниеУслуг` → `Справочник.НаправленияДеятельности`

**Справочник `А_ВидыКонтрагентовДляБаланса` — 5 предопределённых элементов (созданы пользователем):**

| Имя (predefined) | Код | Наименование |
|---|---|---|
| Внутригрупповые | 000000001 | Внутригрупповые |
| Внутренние | 000000002 | Внутренние подразделения |
| Собственники | 000000003 | Собственники |
| Внешние | 000000004 | Внешние |
| Кредиторы | 000000005 | Кредиторы |

**Факты по данным (discovery 2026-06-12):**
- Папка `_КОРПОРАТИВНІ` — в `Справочник.Партнеры` (НЕ Контрагенты; Контрагенты неиерархический, связь `Контрагент.Партнер`). 123 контрагента, **1447 договоров** (`Договор.Партнер В ИЕРАРХИИ`).
- Раскладка разметки: ветка «Група наших компаній» → Внутригрупповые (**331**), ветка «Підрозділи» → Внутренние (**1106**), партнёр «Воронцов Олександр Володимирович» (элемент под корнем) → Собственники (**10**). Вне трёх веток — 0 договоров.
- `А_НаправлениеОказаниеУслуг` уже заполнен пользователями на **~930** договорах (Производство 402, Спецтехника 223, Строительство 163, Логистика 69, ЦО 60, Девелопмент 11, И.П.С. 1) — данные пользователя, НЕ перезаполнять.
- `Dim_Directions` (НаправленияДеятельности, 9 строк) в OLAP уже существует. `Dim_Contracts` (8152+) — raw_sql, прецедент расширения: snowflake `Dim_TipyDogovorov`/`Dim_FinAgents` (2026-05-19, см. `olap_powerbi_pl_pbix.md` §13.10).

---

## Phase 0 — Разблокировка записи договоров + разметка (1С, данные)

**Блокер:** в конфигурации БД BaseERP в `ОбщийМодуль.А_СобытияОбъектов`, процедура
`ДоговорыКонтрагентовПередЗаписью` (конец процедуры), есть незакоммиченная битая строка:

```bsl
А_ПодразделениеОказаниеУслугПодразделению =Источник.А_ПодразделениеОказаниеУслугПодразделению.НаправлениеДеятельности;
```

У `СтруктураПредприятия` нет реквизита `НаправлениеДеятельности` (есть `А_НаправлениеДеятельности`) →
**любая** запись договора с `А_НеОбновлятьГруппаФинансовогоУчета = Ложь` падает:
«Поле объекта не обнаружено (НаправлениеДеятельности)». В файлах главного каталога
(`C:\Configuration_downloads\BASERP25`) строка уже удалена (файл == HEAD git).

1. Если ошибка ещё воспроизводится — донести фикс до БД: частичная загрузка
   `CommonModules/А_СобытияОбъектов/Ext/Module.bsl` (skill `/db-load-xml -Mode Partial` + `-UpdateDB`)
   ИЛИ пользователь удаляет строку в конфигураторе + F7.
2. Разметка договоров (скрипт готов, идемпотентен):
   `python _Rarzrabotki/Python/test/vg_mark_contracts.py` (DRY RUN) → проверить раскладку 331/1106/10, вне веток 0 →
   `python _Rarzrabotki/Python/test/vg_mark_contracts.py APPLY`.
   Скрипт пишет через `ОбменДанными.Загрузка = Истина`. Контрольный запрос в конце должен показать:
   Внутригрупповые 331, Внутренние подразделения 1106, Собственники 10.
3. «Внешние» (004) и «Кредиторы» (005) массово НЕ проставлять — ручная классификация пользователя.

## Phase 1 — Mapping (refresh_mapping)

1. В `mapping/refresh_mapping.py` добавить в WHITELIST `Справочник.А_ВидыКонтрагентовДляБаланса`
   (memory `olap_refresh_mapping_strict_whitelist`; `НаправленияДеятельности` уже в WHITELIST).
   Это НЕ enum — FROZEN_ENUMS не трогать.
2. Прогнать `refresh_mapping.py` (venv) → получить `_FldNNN` для `А_ВидКонтрагента` и
   `А_НаправлениеОказаниеУслуг` в `_Reference<ДоговорыКонтрагентов>` + `_ReferenceNNN`
   таблицу `А_ВидыКонтрагентовДляБаланса`.
3. Регресс: `python main.py` полностью зелёный ДО изменений пайплайнов (дрейф _FldNNN мог
   сломать существующие шаги — ловить здесь, не в Phase 3).

## Phase 2 — SQL DDL (OlapBASERP)

1. `CREATE TABLE Dim_VidyKontragentov` — колонки по образцу `Dim_TipyDogovorov`/`Dim_FinAgents`
   (ключ `VidKontragenta_ID` — тип как `Contract_ID` в `Dim_Contracts`; `Name`, `PredefinedName`, `Code`).
2. `ALTER TABLE Dim_Contracts ADD VidKontragenta_ID <тип ключа> NULL, NapravlenieUslug_ID <тип ключа> NULL;`
   (вторая — FK на существующий `Dim_Directions.Direction_ID`).
3. DDL-скрипт положить в `Ai_Olap/scripts/` (образец: `ddl_dim_tax_types.sql`).

## Phase 3 — ETL (pipelines/dim_catalogs.json)

1. Новый step `dim_vidy_kontragentov`: raw_sql по `_ReferenceNNN` (uuid из `_IDRRef`,
   `_Description`, `_Code`, `_PredefinedID`→имя при наличии; образец — шаги `dim_tipy_dogovorov`/`dim_fin_agents`),
   loader `Dim_VidyKontragentov`, `full_reload`.
2. Расширить raw_sql шага `dim_contracts` двумя колонками: `_Fld<ВидКонтрагента>RRef → VidKontragenta_ID`,
   `_Fld<НаправлениеОказаниеУслуг>RRef → NapravlenieUslug_ID` (через `varbinary_to_uuid`;
   пустые ссылки 16×0x00 → NULL — штатно).
3. `python main.py` (venv): полный прогон. Row counts: Dim_VidyKontragentov = 5;
   Dim_Contracts без изменения числа строк.

## Phase 4 — Verify (новый `verify_olap_dim_vid_kontragenta.py` по образцу `verify_olap_balance_raschety_kontragent.py`)

Гейты (все — assert, до копейки/до штуки):
1. `Dim_VidyKontragentov` = 5 строк, наименования из таблицы выше.
2. `Dim_Contracts`: `VidKontragenta_ID` NOT NULL = 1447, раскладка 331/1106/10 по видам; FK orphans = 0.
3. `Dim_Contracts`: `NapravlenieUslug_ID` NOT NULL ≈ 930; FK → `Dim_Directions` orphans = 0;
   раскладка по направлениям == 1С-эталону (Производство 402, Спецтехника 223, Строительство 163, …).
4. **Кросс-сверка баланса**: Σ `Sum_Close` `Fact_Balance` (период 2026-01) JOIN `Dim_Contracts`
   WHERE вид = «Внутригрупповые» == Python COM запрос к РС `А_ОтчетБаланс_Свод`
   (детали `Свод_РасчетыСПартнерами`, `Договор.А_ВидКонтрагента = Внутригрупповые`) — до копейки.
5. **Регресс**: полный баланс `Fact_Balance` 2026-01 Σ Sum_Close (все Source) = 0,00,
   Актив=|Пассив| = 288 787 750,11; PnL Глобино-2 2026-02 = 38 432 968,66 (`test_etl_acceptance_globyno2.py`).

⚠️ Строки без договора (плуги `Расхождение=Истина`, неторговые статьи) в разрез не попадают —
в отчёте это «(пусто)». НЕ дефект, by design.

## Phase 5 — PL.pbix (MCP powerbi-modeling-mcp)

1. Таблица «ВидыКонтрагентов» (партиция на `Dim_VidyKontragentov`) — **ОБЯЗАТЕЛЬНО native query**
   `Sql.Database(..., [Query="SELECT ..."])`: навигатор кеширует схему, новые таблицы/колонки
   иначе не видны (memory `pbix_sql_database_schema_cache`, прецедент §13.1/§13.10).
2. Партицию «ДоговорыКонтрагентов» переписать/обновить native query с двумя новыми колонками.
3. Связи (обе Many→One, single direction, active):
   - `ДоговорыКонтрагентов[VidKontragenta_ID]` → `ВидыКонтрагентов[VidKontragenta_ID]`
   - `ДоговорыКонтрагентов[NapravlenieUslug_ID]` → `НаправленияДеятельности[Direction_ID]`
   Неоднозначности путей нет: `ДоговорыКонтрагентов` не связана с `Fact_PnL`
   (у Fact_PnL своя прямая связь с НаправленияДеятельности), к `Fact_Balance` путь один —
   через `[Contract_ID]`. После создания — model validate/Calculate.
4. 1С-нотация: видимые колонки кириллицей, тех. ID скрыть (как в §13.10).
5. Гейт drill-down (DAX через MCP): Fact_Balance по `ВидыКонтрагентов[Наименование]` за 2026-01 —
   «Внутригрупповые» == сумме из Phase 4.4; срез по `НаправленияДеятельности` через договор работает.
6. **User-gated**: Desktop Refresh (Fact_Balance, ДоговорыКонтрагентов, ВидыКонтрагентов) + Ctrl+S.
   Строки/визуал в отчёте «Управлінський Баланс» (расшифровка «з клієнтами» по видам) — пользователь
   сам или отдельной задачей.

## Phase 6 — Актуализация знаний

1. `KNOWLEDGE_MAP_OLAP.md` — новая строка в таблицу состояния (дата, row counts, verify PASS).
2. `olap_sql_schema.md` (+Dim_VidyKontragentov, +2 колонки Dim_Contracts), `olap_etl_pipeline.md`
   (+step), `olap_powerbi_pl_pbix.md` (+таблица/связи).
3. Memory: обновить `vg_zadolzhennost_vidy_kontragentov.md` (этап OLAP DONE, фактические цифры).
4. Commit (сообщение на русском/украинском), ветка `claude/*`.

## Правила выполнения (обязательные)

- **Rule #-1**: каждый 1С-запрос — Python COM тест ДО использования. Алиасы не из ключевых слов (`ПО`, `ИЗ`, `В`).
- ETL только venv-python; запись в OlapBASERP — после зелёного DRY/регресса.
- Типовые регистры и структуру РС `А_ОтчетБаланс_Свод` НЕ менять (OD-9).
- Python COM: `str()` от перечислений/ссылок не работает — `erp.String()` / `XMLСтрока`; пустые ссылки — `erp.ЗначениеЗаполнено()`.
- Финал — `superpowers:verification-before-completion`: все гейты Phase 4/5 с выводом команд.
