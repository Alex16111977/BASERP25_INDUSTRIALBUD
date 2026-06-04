# SCRIPTS_CATALOG — Python COM скрипты knowledge_Balanse_klient

> **Назначение:** каталог рабочих Python COM-скриптов для будущих ИИ-сессий.
> **Правило:** прежде чем писать новый скрипт — проверить этот каталог.
> Если задача похожа — использовать существующий образец (адаптировать параметры).
> **НЕ ПЛОДИТЬ дубли.**

## Структура

Все скрипты в `Python/test/`, подключаются к BAS ERP через `V83.COMConnector`.
Подключение и утилиты — в `_common.py` (НЕ ТРОГАТЬ).

**Соглашение `_common.py`:**
```python
from _common import connect_erp, get_refs, money, save_csv, save_json, load_json, get_uuid, get_type_name, ARTIFACTS_DIR
erp = connect_erp()
refs = get_refs(erp)  # dict: Орг, Подр (Глобино-2), Статья (ЗадКл), Источник (РСКПС), НашеПредприятие
ORG = refs["Орг"]
S = erp.String  # сериализатор: ref → строка
```

## ⭐ Канонические паттерны (запоминать наизусть)

| Паттерн | Канонический скрипт | Что делает |
|---|---|---|
| **Сверка ПАП↔РСКПС/РСППС** | `11_match_via_dokumentregistrator.py` | КЛЮЧ: ПАП.Регистратор vs **РСКПС.ДокументРегистратор** (не Регистратор!) |
| **Entry-point Расхождение=Истина** | `14_find_raskhozhdenie_istina.py` | Сводка по (Месяц × Source × Подр × Статья) из РС А_ОтчётБаланс_Свод |
| **Verify плуг (BASELINE/VERIFY)** | `19_verify_raskhozhdenie.py` | Snapshot + diff vs baseline; `python 19... BASELINE\|VERIFY` |
| **Σ-инвариант баланса** | `32_verify_uprbalance_report.py` | gate: Σ signed=0 + \|Актив\|=288 М + ПАП=ПереносАванса |
| **Discovery массовая Σ Δ** | `20_discovery_all_podr_2025.py` + `20v2_discovery_dec25_apr26.py` | ПАП vs РСКПС+РСППС по всем подразделениям |
| **Фильтр ХозОп=ПереносАванса** | `15v3_perenosavansa_docs.py` | строки РСКПС/РСППС где ХозОп=ПереносАванса |
| **Массовый репост** | `17v3_repost_perenosavansa.py` + `18_repost_finrez_balans.py` | Get-or-Create + Записать(Проведение); A_ФинРез_Баланс с argv month\|ALL |
| **Cross-DB diagnostics** | `33_compare_baserazr.py` | сравнение BaseERP vs BaseERPRazr |
| **360° drill документа** | `31_deep_diag_7576.py` | все регистры (ПАП/РСКПС/РСППС/РасчСПост/Выручка/РегРасчетов) |
| **Drill ТЧ через obj.Метаданные()** | `phase6_dec25_co_vvod_rows.py` | обход всех ТЧ + дамп полей |

## Полный каталог (38 рабочих скриптов)

### Phase 0 detective Глобино-2 (Δ7,19М, инцидент 286/287)

| Скрипт | Назначение | Главный регистр | Артефакт |
|---|---|---|---|
| `01_balance_pap_klient_globino2.py` | НМ/Приход/Расход/КМ ПАП.ЗадКл (Глобино-2) ноя+дек 2025 — образец signed-суммы | `РН.ПрочиеАктивыПассивы` | `01_pap_balances.json` |
| `02_balance_rsk_klient_globino2.py` | НМ/КМ РСКПС нетто-долг (Глобино-2) — образец разделения долг/аванс | `РН.РСКПС` + `РС.АналитикаУчётаПоПартнёрам` | `02_rsk_balances.json` |
| `03_reconcile_balances.py` | Sanity-check на 4 контрольных точках | csv 01+02 | `03_balance_reconciliation.csv` |
| ⭐ `11_match_via_dokumentregistrator.py` | **Канон сверки ПАП↔РСКПС** — ПАП.Регистратор vs РСКПС.ДокументРегистратор + категоризация Оба_ОК/Только_ПАП/Только_РСК | `РН.ПАП` ↔ `РН.РСКПС` | `11_match_via_dokreg.csv` |
| `12_drilldown_spis_286_287.py` | Drill инцидента СписаниеБезнал 000Ц-000286/287 (образец per-документ диагностики) | `Док.СписаниеБезналДС` + `РН.ПАП/РСКПС` | print |

### Phase 0 discovery+analysis (массовая Σ Δ для DvAktPas FIX-2026-05-23)

| Скрипт | Назначение | Главный регистр | Артефакт |
|---|---|---|---|
| ⭐ `20_discovery_all_podr_2025.py` | Σ Δ ПАП vs РСКПС+РСППС по всем подразделениям ТОВ за 2025 | `РН.ПАП` + `РН.РСКПС` + `РН.РСППС` | `20_full_discovery.csv` |
| `20v2_discovery_dec25_apr26.py` | То же на расширенный период 01.12.2025—31.12.2026 (residual verify после FIX) | те же | `20v2_full_discovery.csv` |
| `21_discovery_perenosavansa_rows.py` | Все строки РСКПС/РСППС с ХозОп=ПереносАванса (показывает что затронет правка DvAktPas) | `РН.РСКПС/РСППС` (фильтр ХозОп) | `21_perenosavansa_rows.csv` |
| `22_discovery_typed_breakdown.py` | Pivot Σ Δ по (ТипДок × ХозОп) — классификатор | csv 20 + COM `ХозОперация` | `22_typed_breakdown.csv` |
| `23_discovery_etalon_uprbalance.py` | Сверка ПАП vs штатный `Отчёт.УправленческийБаланс` per-Подр на 31.12.2025 — **gate Phase 0** | `РН.ПАП` (агрегат) | `23_etalon_check.csv` |
| `24_discovery_report.py` | Сборка `DISCOVERY_REPORT.md` из 20-23 | csv 20+21+22 | `DISCOVERY_REPORT.md` |
| `25_analysis_root_cause_matrix.py` | Классификация документов C1/C2/C3 на массиве 20 | csv 20 + COM | `25_root_cause_matrix.csv` |
| `25v2_classify_residuals.py` | То же на 20v2 (что осталось после фикса) | csv 20v2 | `25v2_residuals_classified.csv` |
| `26_analysis_double_count_risk.py` | Кандидаты на двойной учёт после правки DvAktPas — **gate перед фиксом** | `РН.ПАП` + `РН.РСКПС/РСППС` | `26_double_count_candidates.csv` |
| `27_analysis_local_branches_coverage.py` | Regex-проверка наличия ветки ВозвратОплатыКлиенту в 3 ManagerModule.bsl (статанализ) | файлы `.bsl` | `27_local_branches.csv` |

### Phase 0 repost + cross-DB diagnostics

| Скрипт | Назначение | Главный регистр | Артефакт |
|---|---|---|---|
| `29_repost_test_docs.py` | Smoke-репост 26_double_count_candidates (паттерн Get-or-Create + Записать) | `Док.*.Записать(Проведение)` | `29_repost_results.csv` |
| `31_deep_diag_7576.py` | Глубокая диагностика проблемного документа: ВСЕ регистры (ПАП/РСППС/РасчСПост/Выручка/РегРасчётов) — **образец 360°** | 7+ регистров | print |
| ⭐ `32_verify_uprbalance_report.py` | Σ-инвариант баланса (Σ signed=0, \|Актив\|=288 М) — **gate** | `РН.ПАП` (агрегат) | print |
| `33_compare_baserazr.py` | Сравнение документа в BaseERP vs BaseERPRazr (cross-DB) | 2 COM connections | print |
| `38_run_otlozhennoe_raspredelenie.py` | Вызов `ОперативныеВзаиморасчётыСервер.ВыполнитьОтложенноеРаспределение()` через COM — **образец server-side** | server-side вызов | print |
| `39_check_options_both.py` | Компактное сравнение констант+ФО двух баз | `Константы.*` + `ФункциональныеОпции` | print |
| `40_check_all_22_after_manual.py` | Verify документов после ручного UI-перепроведения | `РН.ПАП` + `РН.РСППС` | print |

### Phase 5 (декабрь 2025 → апрель 2026, ПереносАванса репост)

| Скрипт | Назначение | Главный регистр | Артефакт |
|---|---|---|---|
| ⭐ `14_find_raskhozhdenie_istina.py` | Сводка Расхождение=Истина в А_ОтчётБаланс_Свод (entry-point Phase 5/6) | `РС.А_ОтчётБаланс_Свод` | `14_raskhozhdenie_summary.csv` + `_records.csv` + `_recorders.csv` |
| ⭐ `15v3_perenosavansa_docs.py` | Документы с ХозОп=ПереносАванса (**правильный фильтр** под DvAktPas FIX) | `РН.РСКПС/РСППС` | `15v3_perenosavansa_rows.csv` + `15v3_docs_to_repost.csv` |
| ⭐ `17v3_repost_perenosavansa.py` | Массовый репост 204 ПереносАванса-документов из 15v3 | `Док.*.Записать(Проведение)` | `17v3_progress.log` + `_errors.csv` + `_success.csv` |
| ⭐ `18_repost_finrez_balans.py` | Перепроведение `А_ФинРез_Баланс` (argv `<YYYY-MM>\|ALL`) | `Док.А_ФинРез_Баланс` | `18_finrez_progress.log` |
| ⭐ `19_verify_raskhozhdenie.py` | **Главный verify** — BASELINE/VERIFY snapshot + diff | `РС.А_ОтчётБаланс_Свод` | `19_baseline.json` + `19_current.json` + `19_diff.md` |

### Phase 6 (ЦО ВводОстатков ЮЕйДрім)

| Скрипт | Назначение | Главный регистр | Артефакт |
|---|---|---|---|
| `phase6_dec25_recheck.py` | Re-check одного месяца на Расхождение=Истина (после правок) | `РС.А_ОтчётБаланс_Свод` | print |
| `phase6_dec25_co_drill.py` | Детектив пары ±X (per-Подр × per-Статья) ПАП vs РСКПС | `РН.ПАП` + `РН.РСКПС` | print |
| `phase6_dec25_co_exact.py` | Точное воспроизведение `Свод_РасчетыСПартнерами` (логика втРасч из ObjectModule.bsl) | `РН.РСКПС.ОстаткиИОбороты` | print |
| `phase6_dec25_co_history.py` | Поиск виновника по ВСЕЙ истории (вне периода месяца) | `РН.ПАП` + `РН.РСКПС` | print |
| `phase6_dec25_co_vvod_rows.py` | Обход всех ТЧ документа через `obj.Метаданные().ТабличныеЧасти` — **образец** | COM объект | print |
| `phase6_co_fix_step1_diag.py` | Диагностика парной строки ТЧ + поиск парного документа (Долги ↔ Авансы) | COM ТЧ | print |
| `phase6_co_fix_step2_move.py` | Backup ТЧ в JSON + перенос строки между документами (Step 2) | COM ТЧ | JSON backup + print |
| `phase6_co_fix_step2b_diag_yuei.py` | Per-партнёр drill РСКПС/ПАП движений (паттерн фильтра по партнёру) | `РН.РСКПС` + `РН.ПАП` | print |
| ⭐ `phase6_co_fix_step3_adjust.py` | **Финальный применённый фикс** — корректировка `Сумма` строки ТЧ + перепроведение | COM ТЧ + Записать | print |
| `phase6_co_fix_rollback.py` | Страховочный откат значений | COM ТЧ + Записать | print |

### Общий модуль

| Файл | Назначение |
|---|---|
| `_common.py` | **НЕ ТРОГАТЬ** — `connect_erp`, `get_refs`, `money`, `save_csv`, `save_json`, `load_json`, `get_uuid`, `get_type_name`, `ARTIFACTS_DIR` |

## Удалённые скрипты (25, дубли/устаревшие)

| Скрипт | Причина |
|---|---|
| `04_movements_pap_by_recorder.py` | Архитектурно неверный паттерн (Регистратор в РСКПС служебный) → вытеснен 11 |
| `05_movements_rsk_by_recorder.py` | Same → вытеснен 11 |
| `06_reconcile_movements.py` | Сборка по неверным 04+05 → false picture; канон 11 |
| `07_drilldown_top_diff.py` | Работает над ошибочным 06; реальный drill — 12 + phase6_dec25_co_drill |
| `08_dump_payment_docs.py` | Артефакт ошибочной диагностики 06 |
| `09_recheck_doc007567.py` | One-shot инцидент 00DL-007567 (закрыт) |
| `10_deep_diag_doc007567.py` | Same → паттерн в 31 |
| `13_recheck_after_delete_286.py` | One-shot после удаления 286 |
| `15_primary_docs_via_dokreg.py` | Дубль v1 — даёт 9 674 false-positive → канон 15v3 |
| `15v2_primary_docs_via_problem_groups.py` | Дубль v2 — мис-классифицирует → канон 15v3 |
| `16_build_repost_registry.py` | Опирается на устаревший 15; 17v3 читает прямо 15v3 |
| `17_repost_primary_docs.py` | Дубль (читал 16) → вытеснен 17v3 |
| `17_smoke_test.py` | Smoke над 16 (закрыт) |
| `17_repost_by_month.py` | Дубль v2 (читал 15v2) → вытеснен 17v3 |
| `17_dec25_smoke.py` | Smoke над 15v2 декабря (закрыт) |
| `34_compare_constants.py` | Дубль — компактнее в 39 |
| `phase6_dec25_repost_balans.py` | Дубль `18_repost_finrez_balans.py 2025-12` |
| `30_drill_priobret_7576.py` | Поверхностный drill → поглощён 31 + 40 |
| `32_repost_22_priobret.py` | One-shot, инцидент закрыт (правило в MEMORY) |
| `35_compare_dogovor_ap.py` | One-shot для 007576 |
| `36_check_registrator_raschetov.py` | One-shot для 007576 |
| `37_check_registrator_000000971.py` | One-shot конкретный РегРасчётов |
| `phase6_dec25_co_doc_diff.py` | Per-месяц не нашёл; канон `_co_history.py` |
| `phase6_dec25_co_vvod.py` | Тривиальная проба; есть в `_vvod_rows` |
| `19_baseline_from_14.py` | One-shot конверсия 14→19; baseline уже в _artifacts/ |

## Рабочие процессы

### Процесс 1: Расследование расхождения в А_ОтчётБаланс_Свод
```
1. python 14_find_raskhozhdenie_istina.py  → группы Расхождение=Истина
2. python 20v2_discovery_dec25_apr26.py    → массовая Σ Δ по подразделениям
3. python 25v2_classify_residuals.py       → классификация C1/C2/C3
4. python phase6_dec25_co_drill.py         → per-Подр drill для конкретной группы
5. python phase6_dec25_co_history.py       → расширенный поиск (вся история)
6. python phase6_dec25_co_vvod_rows.py     → drill ТЧ виновного документа
```

### Процесс 2: Massive repost фиксации
```
1. python 19_verify_raskhozhdenie.py BASELINE   → snapshot ДО
2. python 15v3_perenosavansa_docs.py             → список документов
3. python 17v3_repost_perenosavansa.py           → массовый репост (UI! не COM для РСКПС/РСППС)
4. python 18_repost_finrez_balans.py ALL         → пересборка свода
5. python 19_verify_raskhozhdenie.py VERIFY      → snapshot ПОСЛЕ vs baseline
6. python 32_verify_uprbalance_report.py         → Σ-инвариант не нарушен
```

### Процесс 3: Cross-DB regression check
```
1. python 33_compare_baserazr.py    → сравнить документ в BaseERP vs BaseERPRazr
2. python 39_check_options_both.py  → константы/ФО двух баз совпадают?
3. python 38_run_otlozhennoe_raspredelenie.py  → пересчёт оперативных взаиморасчётов
```

### Процесс 4: Точечная правка ТЧ ВводОстатков
```
1. python phase6_co_fix_step1_diag.py        → ТЧ + парный документ
2. python phase6_co_fix_step2b_diag_yuei.py  → текущее состояние РСКПС
3. python phase6_co_fix_step3_adjust.py      → применить правку
4. python 18_repost_finrez_balans.py 2025-12 → пересборка свода декабря
5. python phase6_dec25_recheck.py            → плуг исчез?
6. ЕСЛИ ЧТО → python phase6_co_fix_rollback.py
```

## Правила для новых скриптов

1. **Сначала проверить каталог.** Если задача похожа — адаптировать существующий, а не плодить новый.
2. **Импортировать `_common.py`** — не дублировать connect/refs/money/save_csv.
3. **Даты СЕРВЕРНО** (`ДАТАВРЕМЯ(2025,12,1)` в тексте запроса), не через `SetParameter(datetime)` (memory `feedback_balans_etalon_period_serverside`).
4. **Сверка ПАП↔РСКПС** — только по `ПАП.Регистратор ↔ РСКПС.ДокументРегистратор` (memory FINDINGS п.8).
5. **Перепроведение** — `obj.Записать(erp.РежимЗаписиДокумента.Проведение)`, не число 1.
6. **Не удалять документы** в тестах (memory `feedback_no_doc_delete_in_tests`).
7. **Не перепроводить через COM** документы пишущие в РСКПС/РСППС (memory `feedback_com_repost_skips_registrator_raschetov`) — только UI.

## Связанные файлы

- `KNOWLEDGE_MAP.md` — главная навигация
- `LESSONS.md` — паттерны и антипаттерны
- `FINDINGS.md` — исторический детектив Глобино-2
- `docs/PHASE6_REPORT_dec25.md` — Phase 6 финал
- `docs/DISCOVERY_REPORT.md` / `ANALYSIS_REPORT.md` / `STRATEGY_DECISION.md` / `VERIFY_REPORT.md` — Phase 0-4 отчёты
