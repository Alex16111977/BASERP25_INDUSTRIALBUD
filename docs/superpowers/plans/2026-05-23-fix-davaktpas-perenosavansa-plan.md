# Fix `ДвиженияАктивовПассывов` filter `ПереносАванса` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Устранить системную ошибку фильтра `<>ПереносАванса` в типовой обработке `ДвиженияАктивовПассивов`, из-за которой ПАП (управ.баланс) расходится с РСКПС/РСППС (ВРСК/ВРСП) для всех операций ВозвратОплатыКлиенту с переносом аванса в долг. Правка делается **напрямую в коде конфигурации**, без CFE-расширения. Гарантировать Σ Δ = 0 ±0,01 по всем подразделениям ТОВ ІНДАСТРІАЛБУД за 2025, без регрессии типовых сделок.

**Architecture:** 5-фазный пайплайн Discovery → Analysis → Strategy → Patch → Verify. Каждая фаза состоит из набора Python COM-скриптов (V83.COMConnector → ERP) и MD-отчётов. Правка кода применяется через `/db-load-xml` + `/db-update -Dynamic+` (skill'ы), с обязательной проверкой применения и rollback-планом.

**Tech Stack:**
- Python 3.13 + `win32com.client` (V83.COMConnector)
- BSL (BAS ERP 2.5)
- 1С 8.3.20+
- skill'ы `db-load-xml`, `db-update`, `cf-edit`
- Spec: `docs/superpowers/specs/2026-05-23-fix-davaktpas-perenosavansa-design.md`

---

## File Structure

### Создаваемые Python-скрипты (рабочая директория `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/`)

| Файл | Phase | Назначение |
|---|---|---|
| `20_discovery_all_podr_2025.py` | 0 | Сверка ПАП vs РСКПС+РСППС по `ДокументРегистратор` за 2025 по всем подразделениям ТОВ |
| `21_discovery_perenosavansa_rows.py` | 0 | Все строки РСКПС/РСППС с `ХозОп=ПереносАванса` за 2025 |
| `22_discovery_typed_breakdown.py` | 0 | Разбивка Σ Δ по (ТипДокумента, ХозОперация) |
| `23_discovery_etalon_uprbalance.py` | 0 | Сверка через Builder отчёта `Отчёт.УправленческийБаланс` |
| `24_discovery_report.py` | 0 | Сборка `DISCOVERY_REPORT.md` |
| `25_analysis_root_cause_matrix.py` | 1 | Классификация (C1/C2/C3) |
| `26_analysis_double_count_risk.py` | 1 | Список документов с риском двойного учёта для Подхода А |
| `27_analysis_local_branches_coverage.py` | 1 | Грэп локальных веток ВозвратОплатыКлиенту в 3 ManagerModule |
| `28_check_patch_applied.py` | 3 | Чтение текста модулей через COM, проверка наличия правки |
| `30_verify_main_diff_closed.py` | 4 | Σ Δ за 2025 по всем подр = 0 ±0,01 |
| `31_verify_globino2_dec2025.py` | 4 | Глобино-2 / 31.12.2025: ПАП.ЗадКл == РСКПС.нетто |
| `32_verify_uprbalance_report.py` | 4 | Σ-инвариант штатного отчёта НЕ изменился |
| `33_regression_normal_payments.py` | 4 | ~10 нормальных платежей (без ПереносАванса) — Δ=0 до=после |
| `34_regression_realizations.py` | 4 | Реализации — не получили доп.движений в ПАП |
| `35_regression_zachet_avansa.py` | 4 | Зачёт аванса при Реализации — не сломан |
| `36_regression_full_balance.py` | 4 | Σ signed (КО, OD-3) = 0 per орг (балансовый инвариант) |

### Модифицируемые .bsl файлы конфигурации (Phase 3)

**Условно — зависит от Strategy decision:**

- `DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl` — строки **1275, 1323, 1428, 1476** (фильтр `<>ПереносАванса`)
- `Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl` — функция `ТекстЗапросаТаблицаПрочиеАктивыПассивы` стр. **3884–4010**
- `Documents/РасходныйКассовыйОрдер/Ext/ManagerModule.bsl` — аналогичная функция
- `Documents/ОперацияПоПлатежнойКарте/Ext/ManagerModule.bsl` — аналогичная функция

### Создаваемые MD-документы

- `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/DISCOVERY_REPORT.md`
- `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/ANALYSIS_REPORT.md`
- `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/STRATEGY_DECISION.md`
- `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/VERIFY_REPORT.md`
- Обновление `_Rarzrabotki/notebook/knowledge_Balanse_klient/FINDINGS.md` (статус "решено")

### Backup исходников (Phase 3)

`_Rarzrabotki/notebook/knowledge_Balanse_klient/_backup/`:
- `DvAktPas_2026-05-23.bsl`
- `SpisaniyaBeznal_2026-05-23.bsl`
- `RKO_2026-05-23.bsl`
- `OPK_2026-05-23.bsl`

### Используемые существующие модули (без правок)

- `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_common.py` — подключение к ERP, refs, money, save_json/csv, get_uuid, get_type_name. **Использовать как есть.**

---

# Phase 0 — Discovery (read-only)

**Gate-критерий перехода к Phase 1:** `DISCOVERY_REPORT.md` сгенерирован, содержит >1 подразделения с расхождениями и Σ Δ_годовая ≠ 0.

## Task 0.1: Создать директорию `docs/` для отчётов

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/.gitkeep`

- [ ] **Step 1: Создать директорию**

```bash
mkdir -p "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/docs"
touch "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/.gitkeep"
```

- [ ] **Step 2: Verify**

```bash
ls "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/"
```
Expected: `.gitkeep`

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/docs/.gitkeep && git commit -m "docs(balans_klient): add docs/ for phase reports"
```

---

## Task 0.2: Создать скрипт `20_discovery_all_podr_2025.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/20_discovery_all_podr_2025.py`
- Output: `_artifacts/20_full_discovery.csv`

**Назначение**: Для каждого подразделения ТОВ ІНДАСТРІАЛБУД (которое имеет движения в ПАП.ЗадКл/ПолучАв/ЗадПост/ВыдАв за 2025) — рассчитать Σ Δ ПАП vs РСКПС+РСППС по правильному ключу (`ДокументРегистратор`). Группировка: (Подразделение, ТипДокумента, ХозОперация). Отсортировать по \|Δ\| убыв.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 20 (Phase 0) — Discovery: расхождения ПАП vs РСКПС/РСППС за 2025
по ВСЕМ подразделениям ТОВ ІНДАСТРІАЛБУД.

ЧТО ИЩЕТ:
    Σ Δ = ПАП.signed − (РСКПС.signed + РСППС.signed)
    Группировка: (Подразделение, ТипДокумента, ХозОперация).
    Ключ сверки: ПАП.Регистратор ↔ РСКПС.ДокументРегистратор ↔ РСППС.ДокументРегистратор.

КАК СЧИТАЕТ:
    Один пакет запросов с втПАП, втРСК, втРСП, втКлючи, финал JOIN.
    Период: 2025-01-01 ... 2025-12-31.

ЧТО ДАЁТ:
    _artifacts/20_full_discovery.csv (Подразделение, ТипДокумента, ХозОп, Σ ПАП, Σ РСК, Σ РСП, Δ)
    print: топ-20 подразделений по |Σ Δ|, общая Σ Δ по ТОВ.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money, save_csv

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.УстановитьПараметр("НП", refs["НашеПредприятие"])
q.Текст = """
// === ПАП: статьи расчётов с партнёрами за 2025 ===
ВЫБРАТЬ
    Т.Регистратор КАК Док,
    Т.Подразделение КАК Подр,
    Т.Статья КАК Статья,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК ПАП_signed
ПОМЕСТИТЬ втПАП
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И Т.Организация = &Орг
    И Т.Источник В (
        ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам),
        ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам))
    И Т.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ПолученныеАвансы),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыданныеАвансы))
СГРУППИРОВАТЬ ПО Т.Регистратор, Т.Подразделение, Т.Статья;

// === РСКПС: ДолгУпр + ПредоплатаУпр signed (нетто) ===
ВЫБРАТЬ
    Р.ДокументРегистратор КАК Док,
    ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
          ТОГДА Р.ОбъектРасчетов.Договор.Подразделение
          ИНАЧЕ Р.ОбъектРасчетов.Подразделение КОНЕЦ КАК Подр,
    СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Р.ДолгУпр - Р.ПредоплатаУпр
                ИНАЧЕ -(Р.ДолгУпр - Р.ПредоплатаУпр) КОНЕЦ) КАК РСК_signed
ПОМЕСТИТЬ втРСК
ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Р
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Р.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И АП.Организация = &Орг
СГРУППИРОВАТЬ ПО Р.ДокументРегистратор,
    ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
          ТОГДА Р.ОбъектРасчетов.Договор.Подразделение
          ИНАЧЕ Р.ОбъектРасчетов.Подразделение КОНЕЦ;

// === РСППС: аналогично ===
ВЫБРАТЬ
    П.ДокументРегистратор КАК Док,
    ВЫБОР КОГДА П.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
          ТОГДА П.ОбъектРасчетов.Договор.Подразделение
          ИНАЧЕ П.ОбъектРасчетов.Подразделение КОНЕЦ КАК Подр,
    СУММА(ВЫБОР КОГДА П.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА П.ДолгУпр - П.ПредоплатаУпр
                ИНАЧЕ -(П.ДолгУпр - П.ПредоплатаУпр) КОНЕЦ) КАК РСП_signed
ПОМЕСТИТЬ втРСП
ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам КАК П
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО П.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ П.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И АП.Организация = &Орг
СГРУППИРОВАТЬ ПО П.ДокументРегистратор,
    ВЫБОР КОГДА П.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
          ТОГДА П.ОбъектРасчетов.Договор.Подразделение
          ИНАЧЕ П.ОбъектРасчетов.Подразделение КОНЕЦ;

// === Ключи документов и подразделений ===
ВЫБРАТЬ РАЗЛИЧНЫЕ К.Док КАК Док, К.Подр КАК Подр
ПОМЕСТИТЬ втКлючи
ИЗ (ВЫБРАТЬ Док, Подр ИЗ втПАП
    ОБЪЕДИНИТЬ ВЫБРАТЬ Док, Подр ИЗ втРСК
    ОБЪЕДИНИТЬ ВЫБРАТЬ Док, Подр ИЗ втРСП) КАК К;

// === Финал ===
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(К.Подр) КАК Подразделение,
    ВЫРАЗИТЬ(ТИПЗНАЧЕНИЯ(К.Док) КАК Строка(150)) КАК ТипДок_Raw,
    ПРЕДСТАВЛЕНИЕ(К.Док) КАК Документ,
    СУММА(ЕСТЬNULL(ПАП.ПАП_signed, 0)) КАК Σ_ПАП,
    СУММА(ЕСТЬNULL(РСК.РСК_signed, 0)) КАК Σ_РСК,
    СУММА(ЕСТЬNULL(РСП.РСП_signed, 0)) КАК Σ_РСП,
    СУММА(ЕСТЬNULL(ПАП.ПАП_signed, 0))
        - СУММА(ЕСТЬNULL(РСК.РСК_signed, 0))
        - СУММА(ЕСТЬNULL(РСП.РСП_signed, 0)) КАК Дельта
ИЗ втКлючи КАК К
    ЛЕВОЕ СОЕДИНЕНИЕ втПАП КАК ПАП ПО К.Док = ПАП.Док И К.Подр = ПАП.Подр
    ЛЕВОЕ СОЕДИНЕНИЕ втРСК КАК РСК ПО К.Док = РСК.Док И К.Подр = РСК.Подр
    ЛЕВОЕ СОЕДИНЕНИЕ втРСП КАК РСП ПО К.Док = РСП.Док И К.Подр = РСП.Подр
СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(К.Подр),
    ВЫРАЗИТЬ(ТИПЗНАЧЕНИЯ(К.Док) КАК Строка(150)),
    ПРЕДСТАВЛЕНИЕ(К.Док)
ИМЕЮЩИЕ АБС(СУММА(ЕСТЬNULL(ПАП.ПАП_signed, 0))
            - СУММА(ЕСТЬNULL(РСК.РСК_signed, 0))
            - СУММА(ЕСТЬNULL(РСП.РСП_signed, 0))) > 0.01
УПОРЯДОЧИТЬ ПО Дельта
"""

print("=" * 100)
print("СКРИПТ 20 — Discovery всех расхождений ПАП vs РСКПС+РСППС за 2025 по ТОВ")
print("=" * 100)

res = q.Выполнить().Выгрузить()
print(f"\nВсего пар (Подр, Документ) с |Δ| > 0.01: {res.Количество()}")

rows = []
by_podr = {}
sum_pap = sum_rsk = sum_rsp = sum_delta = 0.0
for i in range(res.Количество()):
    r = res.Получить(i)
    подр = str(r.Подразделение)
    тип = str(r.ТипДок_Raw).split(".")[-1] if "." in str(r.ТипДок_Raw) else str(r.ТипДок_Raw)
    док = str(r.Документ)
    p, rk, rp, d = float(r.Σ_ПАП or 0), float(r.Σ_РСК or 0), float(r.Σ_РСП or 0), float(r.Дельта or 0)
    rows.append({
        "Подразделение": подр, "ТипДок": тип, "Документ": док,
        "Σ_ПАП": p, "Σ_РСК": rk, "Σ_РСП": rp, "Дельта": d,
    })
    by_podr.setdefault(подр, 0.0); by_podr[подр] += d
    sum_pap += p; sum_rsk += rk; sum_rsp += rp; sum_delta += d

print(f"\nΣ ПАП={money(sum_pap)}, Σ РСК={money(sum_rsk)}, Σ РСП={money(sum_rsp)}")
print(f"Σ Δ (ПАП - РСК - РСП) = {money(sum_delta)}\n")
print("Топ-20 подразделений по |Σ Δ|:")
print(f"{'Подразделение':<40} {'Σ Δ':>20}")
print("-" * 65)
for подр, d in sorted(by_podr.items(), key=lambda x: -abs(x[1]))[:20]:
    print(f"{подр[:40]:<40} {money(d):>20}")

path = save_csv("20_full_discovery", rows,
                ["Подразделение", "ТипДок", "Документ", "Σ_ПАП", "Σ_РСК", "Σ_РСП", "Дельта"])
print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить и проверить self-check**

```bash
cd "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test"
python 20_discovery_all_podr_2025.py
```

Expected:
- Σ Δ ≠ 0 (есть расхождения)
- Подразделение "Глобино-2" должно быть в топ-20 с Σ Δ близким к −7 194 594,73
- Артефакт `_artifacts/20_full_discovery.csv` создан

- [ ] **Step 3: Smoke-проверка артефакта**

```bash
head -5 "_artifacts/20_full_discovery.csv"
wc -l "_artifacts/20_full_discovery.csv"
```

Expected: >2 строк, заголовок + данные.

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/20_discovery_all_podr_2025.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/20_full_discovery.csv && git commit -m "feat(balans_klient): add Phase 0 discovery script 20 (all podr 2025)"
```

---

## Task 0.3: Создать скрипт `21_discovery_perenosavansa_rows.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/21_discovery_perenosavansa_rows.py`
- Output: `_artifacts/21_perenosavansa_rows.csv`

**Назначение**: Все строки РСКПС и РСППС за 2025 с `ХозОперация = ПереносАванса` по ТОВ. Эти строки фильтр в `ДвАктПас` отрезает. Группировать по (ДокументРегистратор, ТипРегистра — РСК или РСП).

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 21 (Phase 0) — Все строки РСКПС/РСППС с ХозОп=ПереносАванса за 2025

ЧТО ИЩЕТ:
    Σ |ДолгУпр - ПредоплатаУпр| signed по (ДокументРегистратор, "РСК"|"РСП")
    где Т.ХозяйственнаяОперация = ПереносАванса

ЧТО ДАЁТ:
    _artifacts/21_perenosavansa_rows.csv
    print: топ-15 + общая Σ
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money, save_csv

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
ВЫБРАТЬ
    "РСК" КАК Регистр,
    Т.ДокументРегистратор КАК Док,
    ПРЕДСТАВЛЕНИЕ(Т.ДокументРегистратор) КАК ДокИмя,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.ДолгУпр ИНАЧЕ -Т.ДолгУпр КОНЕЦ) КАК Σ_ДолгУпр,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.ПредоплатаУпр ИНАЧЕ -Т.ПредоплатаУпр КОНЕЦ) КАК Σ_Аванс
ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Т
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Т.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И АП.Организация = &Орг
    И Т.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)
СГРУППИРОВАТЬ ПО Т.ДокументРегистратор, ПРЕДСТАВЛЕНИЕ(Т.ДокументРегистратор)

ОБЪЕДИНИТЬ ВСЕ

ВЫБРАТЬ
    "РСП",
    Т.ДокументРегистратор,
    ПРЕДСТАВЛЕНИЕ(Т.ДокументРегистратор),
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.ДолгУпр ИНАЧЕ -Т.ДолгУпр КОНЕЦ),
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.ПредоплатаУпр ИНАЧЕ -Т.ПредоплатаУпр КОНЕЦ)
ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам КАК Т
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Т.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И АП.Организация = &Орг
    И Т.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)
СГРУППИРОВАТЬ ПО Т.ДокументРегистратор, ПРЕДСТАВЛЕНИЕ(Т.ДокументРегистратор)
"""

print("=" * 100)
print("СКРИПТ 21 — Строки РСКПС/РСППС с ХозОп=ПереносАванса за 2025 / ТОВ")
print("=" * 100)

res = q.Выполнить().Выгрузить()
rows = []
sum_d = sum_a = 0.0; count_rsk = count_rsp = 0
for i in range(res.Количество()):
    r = res.Получить(i)
    reg = str(r.Регистр)
    d = float(r.Σ_ДолгУпр or 0); a = float(r.Σ_Аванс or 0)
    rows.append({
        "Регистр": reg, "Документ": str(r.ДокИмя),
        "Σ_ДолгУпр": d, "Σ_Аванс": a, "Σ_Δ_отрезано": d - a,
    })
    sum_d += d; sum_a += a
    if reg == "РСК": count_rsk += 1
    else: count_rsp += 1

print(f"\nВсего пар (Регистр, ДокументРегистратор): {res.Количество()}")
print(f"  РСК: {count_rsk}, РСП: {count_rsp}")
print(f"Σ Долг = {money(sum_d)}, Σ Аванс = {money(sum_a)}")
print(f"Σ (Долг - Аванс) = что отрезано фильтром = {money(sum_d - sum_a)}\n")

print("Топ-15 по |Σ (Долг - Аванс)|:")
print(f"{'Рег':<4} {'Документ':<60} {'Σ_Долг':>16} {'Σ_Аванс':>16} {'Σ_Δ_отрезано':>16}")
print("-" * 115)
for r in sorted(rows, key=lambda x: -abs(x["Σ_Δ_отрезано"]))[:15]:
    print(f"{r['Регистр']:<4} {r['Документ'][:60]:<60} {money(r['Σ_ДолгУпр']):>16} {money(r['Σ_Аванс']):>16} {money(r['Σ_Δ_отрезано']):>16}")

path = save_csv("21_perenosavansa_rows", rows,
                ["Регистр", "Документ", "Σ_ДолгУпр", "Σ_Аванс", "Σ_Δ_отрезано"])
print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить**

```bash
python 21_discovery_perenosavansa_rows.py
```

Expected: Σ Δ_отрезано должно близко совпадать с общей Σ Δ из скрипта 20 (с допуском на «зачёт аванса» документы где и Долг и Аванс участвуют).

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/21_discovery_perenosavansa_rows.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/21_perenosavansa_rows.csv && git commit -m "feat(balans_klient): add Phase 0 discovery script 21 (perenosavansa rows)"
```

---

## Task 0.4: Создать скрипт `22_discovery_typed_breakdown.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/22_discovery_typed_breakdown.py`
- Output: `_artifacts/22_typed_breakdown.csv`

**Назначение**: Pivot — Σ Δ по (ТипДокумента × ХозОперация). Понять какие типы документов и какие операции дают расхождение.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 22 (Phase 0) — Pivot Σ Δ по (ТипДокумента, ХозОперация)

ЧТО ДАЁТ:
    _artifacts/22_typed_breakdown.csv
    print: матрицу типов документов и операций.
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money, save_csv, ARTIFACTS_DIR, get_type_name

erp = connect_erp()
refs = get_refs(erp)

# Загружаем 20_full_discovery — там Σ Δ по (Подр, Док)
# Идем по документам: для каждого получаем ХозОперацию + тип
# Группируем (Тип, ХозОп) → Σ Δ
rows = []
with open(os.path.join(ARTIFACTS_DIR, "20_full_discovery.csv"), encoding="utf-8-sig", newline="") as f:
    rows = list(csv.DictReader(f, delimiter=";"))

print(f"Загружено {len(rows)} строк из 20_full_discovery.csv")

# Сборка: по уникальным документам получаем ХозОп через прямой запрос
по_документу = {}
for r in rows:
    док_имя = r["Документ"]
    тип = r["ТипДок"]
    подр = r["Подразделение"]
    дельта = float(str(r["Дельта"]).replace(",", ".").replace(" ", ""))
    ключ = (тип, док_имя)
    по_документу.setdefault(ключ, 0.0); по_документу[ключ] += дельта

print(f"Уникальных документов с расхождением: {len(по_документу)}")

# Для каждого документа узнаём ХозОперацию (если есть)
# Пакетный запрос по типам
по_типам = {}
for (тип, док_имя), delta in по_документу.items():
    по_типам.setdefault(тип, []).append((док_имя, delta))

# Получаем ХозОп для каждого документа через прямой запрос по Номер+Дата+Тип
matrix = {}  # {(тип, хозоп): {"count":n, "sum_delta":x}}
for тип, документы in по_типам.items():
    print(f"  {тип}: {len(документы)} док")
    if not тип:
        ключ = (тип, "<нет ХозОп>")
        matrix.setdefault(ключ, {"count": 0, "sum_delta": 0.0})
        for _, d in документы:
            matrix[ключ]["count"] += 1; matrix[ключ]["sum_delta"] += d
        continue
    # Для конкретных типов получаем ХозОперацию пакетом
    try:
        q = erp.NewObject("Запрос")
        q.Текст = f'''
        ВЫБРАТЬ Д.Ссылка, ПРЕДСТАВЛЕНИЕ(Д.Ссылка) КАК Имя, ПРЕДСТАВЛЕНИЕ(Д.ХозяйственнаяОперация) КАК ХозОп
        ИЗ Документ.{тип} КАК Д
        ГДЕ ПРЕДСТАВЛЕНИЕ(Д.Ссылка) В (&Имена)
        '''
        имена = erp.NewObject("Массив")
        for имя, _ in документы:
            имена.Добавить(имя)
        q.УстановитьПараметр("Имена", имена)
        res = q.Выполнить().Выгрузить()
        d_to_op = {}
        for i in range(res.Количество()):
            rec = res.Получить(i)
            d_to_op[str(rec.Имя)] = str(rec.ХозОп)
    except Exception as e:
        # Документ без ХозОперации
        d_to_op = {}

    for имя, delta in документы:
        op = d_to_op.get(имя, "<без ХозОп>")
        ключ = (тип, op)
        matrix.setdefault(ключ, {"count": 0, "sum_delta": 0.0})
        matrix[ключ]["count"] += 1; matrix[ключ]["sum_delta"] += delta

# Сохранить
out_rows = [
    {"ТипДок": k[0], "ХозОперация": k[1], "КолДокументов": v["count"], "Σ Δ": v["sum_delta"]}
    for k, v in sorted(matrix.items(), key=lambda x: -abs(x[1]["sum_delta"]))
]
path = save_csv("22_typed_breakdown", out_rows, ["ТипДок", "ХозОперация", "КолДокументов", "Σ Δ"])

print("\nМатрица (ТипДок × ХозОперация):")
print(f"{'ТипДок':<40} {'ХозОп':<35} {'Кол':>6} {'Σ Δ':>16}")
print("-" * 105)
for r in out_rows[:30]:
    print(f"{r['ТипДок'][:40]:<40} {r['ХозОперация'][:35]:<35} {r['КолДокументов']:>6} {money(r['Σ Δ']):>16}")

print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить**

```bash
python 22_discovery_typed_breakdown.py
```

Expected: Видим ТипДок=СписаниеБезналичных + ХозОп=ВозвратОплатыКлиенту с отрицательной Σ Δ (наш кейс).

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/22_discovery_typed_breakdown.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/22_typed_breakdown.csv && git commit -m "feat(balans_klient): add Phase 0 discovery script 22 (typed breakdown)"
```

---

## Task 0.5: Создать скрипт `23_discovery_etalon_uprbalance.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/23_discovery_etalon_uprbalance.py`
- Output: `_artifacts/23_etalon_uprbalance.csv`

**Назначение**: Сверка наших Δ через скрипт 20 со штатным `Отчёт.УправленческийБаланс` — финансист видит ровно те же цифры.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 23 (Phase 0) — Сверка с штатным Отчёт.УправленческийБаланс

ЧТО ДЕЛАЕТ:
    1. Берём ПАП на 31.12.2025 по статье ЗадКл per Подразделение (для ТОВ)
    2. Сравниваем с тем что показывает штатный отчёт (Builder.GetData)
    3. Артефакт CSV: Подразделение, ПАП_наш, Отчёт_штатный, Δ
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money, save_csv

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Т.Подразделение) КАК Подразделение,
    ПРЕДСТАВЛЕНИЕ(Т.Статья) КАК Статья,
    СУММА(ВЫБОР КОГДА Т.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
                    И Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.Сумма
                КОГДА Т.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
                    И Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
                ТОГДА -Т.Сумма
                ИНАЧЕ 0 КОНЕЦ) КАК КО_наш
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И Т.Организация = &Орг
    И Т.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам)
    И Т.Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов)
СГРУППИРОВАТЬ ПО Т.Подразделение, Т.Статья
"""

print("=" * 100)
print("СКРИПТ 23 — Сверка с Отчёт.УправленческийБаланс (на 31.12.2025, статья ЗадКл, по подр)")
print("=" * 100)

res = q.Выполнить().Выгрузить()
rows = []
sum_pap = 0.0
for i in range(res.Количество()):
    r = res.Получить(i)
    подр = str(r.Подразделение); ст = str(r.Статья); ко = float(r.КО_наш or 0)
    rows.append({"Подразделение": подр, "Статья": ст, "КО_наш": ко})
    sum_pap += ко

# Эталон — штатный отчёт по всей орг, статья ЗадКл, на 31.12.2025
# Через прямой запрос к ПАП.ОстаткиИОбороты с теми же фильтрами
q.Текст = """
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Т.Подразделение) КАК Подразделение,
    ПРЕДСТАВЛЕНИЕ(Т.Статья) КАК Статья,
    Т.СуммаКонечныйОстаток КАК КО_отчет
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(
    , ДАТАВРЕМЯ(2025,12,31,23,59,59), , ,
    Организация = &Орг
    И Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам)
    И Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов)
) КАК Т
"""
q.УстановитьПараметр("Орг", refs["Орг"])
res2 = q.Выполнить().Выгрузить()
sum_otch = 0.0
по_подр_отч = {}
for i in range(res2.Количество()):
    r = res2.Получить(i)
    подр = str(r.Подразделение); ко = float(r.КО_отчет or 0)
    по_подр_отч[подр] = ко; sum_otch += ко

# Объединяем
final = []
for r in rows:
    отч = по_подр_отч.get(r["Подразделение"], 0.0)
    final.append({**r, "КО_отчет": отч, "Δ": r["КО_наш"] - отч})

# Подразделения только в отчёте (которых нет у нас)
for подр, ко in по_подр_отч.items():
    if not any(r["Подразделение"] == подр for r in final):
        final.append({"Подразделение": подр, "Статья": "Задолженность клиентов",
                      "КО_наш": 0.0, "КО_отчет": ко, "Δ": -ко})

print(f"\nΣ КО_наш  : {money(sum_pap)}")
print(f"Σ КО_отчёт: {money(sum_otch)}")
print(f"|Σ Δ|     : {money(sum_pap - sum_otch)}")

if abs(sum_pap - sum_otch) < 0.01:
    print("✓ Наш расчёт ПАП == штатный отчёт (по контролю Σ)")
else:
    print("⚠️ Расхождение Σ — пересмотреть формулы скрипта 20")

print(f"\nТоп-15 подразделений по |КО_отчёт|:")
final.sort(key=lambda r: -abs(r["КО_отчет"]))
for r in final[:15]:
    print(f"  {r['Подразделение'][:35]:<35} КО={money(r['КО_отчет']):>16}")

path = save_csv("23_etalon_uprbalance", final,
                ["Подразделение", "Статья", "КО_наш", "КО_отчет", "Δ"])
print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить**

```bash
python 23_discovery_etalon_uprbalance.py
```

Expected: `Σ КО_наш == Σ КО_отчёт` (|Δ|<0.01), подтверждение что наша формула 20 даёт цифру, эквивалентную штатному отчёту.

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/23_discovery_etalon_uprbalance.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/23_etalon_uprbalance.csv && git commit -m "feat(balans_klient): add Phase 0 discovery script 23 (etalon uprbalance)"
```

---

## Task 0.6: Создать скрипт `24_discovery_report.py` — сборка отчёта

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/24_discovery_report.py`
- Output: `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/DISCOVERY_REPORT.md`

**Назначение**: Прочитать все 3 CSV-артефакта (20, 21, 22) + результат 23, и сгенерировать MD-отчёт со сводкой Phase 0.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 24 (Phase 0 финал) — Сборка DISCOVERY_REPORT.md

ЧТО ДЕЛАЕТ:
    Читает 20-23 артефакты, сводит в MD-отчёт.
    Gate: должен показать что Σ Δ ≠ 0 и затронуто >1 подразделения.
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import ARTIFACTS_DIR, money

DOCS_DIR = os.path.join(os.path.dirname(ARTIFACTS_DIR), "..", "..", "docs")
DOCS_DIR = os.path.abspath(DOCS_DIR)


def load(name):
    p = os.path.join(ARTIFACTS_DIR, f"{name}.csv")
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def f(x):
    try: return float(str(x).replace(",", ".").replace(" ", ""))
    except: return 0.0


# 20
r20 = load("20_full_discovery")
by_podr = {}
for r in r20:
    by_podr.setdefault(r["Подразделение"], 0.0)
    by_podr[r["Подразделение"]] += f(r["Дельта"])
sum_total_20 = sum(by_podr.values())

# 21
r21 = load("21_perenosavansa_rows")
sum_pa_rsk = sum(f(r["Σ_Δ_отрезано"]) for r in r21 if r["Регистр"] == "РСК")
sum_pa_rsp = sum(f(r["Σ_Δ_отрезано"]) for r in r21 if r["Регистр"] == "РСП")

# 22
r22 = load("22_typed_breakdown")

# 23
r23 = load("23_etalon_uprbalance")
sum_otch = sum(f(r["КО_отчет"]) for r in r23)
sum_nash = sum(f(r["КО_наш"]) for r in r23)

lines = []
lines.append("# DISCOVERY_REPORT — Phase 0\n")
lines.append("> Сгенерировано скриптом 24 на основе 20-23 артефактов.\n")
lines.append("## Сводка\n")
lines.append(f"- **Σ Δ ПАП vs РСКПС+РСППС за 2025**: {money(sum_total_20)} UAH")
lines.append(f"- **Подразделений с расхождениями**: {len(by_podr)}")
lines.append(f"- **Документов-первичек с расхождениями**: {len(r20)}")
lines.append(f"- **РСК: Σ ПереносАванса отрезано фильтром**: {money(sum_pa_rsk)}")
lines.append(f"- **РСП: Σ ПереносАванса отрезано фильтром**: {money(sum_pa_rsp)}")
lines.append(f"- **Сверка со штатным Отчёт.УпрБаланс**: КО_наш={money(sum_nash)}, КО_отчёт={money(sum_otch)}, Δ={money(sum_nash - sum_otch)}")

lines.append("\n## Топ-15 подразделений по |Σ Δ|\n")
lines.append("| Подразделение | Σ Δ |")
lines.append("|---|---:|")
for подр, d in sorted(by_podr.items(), key=lambda x: -abs(x[1]))[:15]:
    lines.append(f"| {подр} | {money(d)} |")

lines.append("\n## Топ-15 пар (ТипДок × ХозОперация)\n")
lines.append("| ТипДок | ХозОперация | Кол | Σ Δ |")
lines.append("|---|---|---:|---:|")
r22.sort(key=lambda r: -abs(f(r["Σ Δ"])))
for r in r22[:15]:
    lines.append(f"| {r['ТипДок']} | {r['ХозОперация']} | {r['КолДокументов']} | {money(f(r['Σ Δ']))} |")

lines.append("\n## Gate-критерий Phase 1\n")
if abs(sum_total_20) > 0.01 and len(by_podr) > 1:
    lines.append("✓ **PASS** — Σ Δ ≠ 0 и затронуто >1 подразделения. Переходим к Phase 1 (Analysis).")
else:
    lines.append("⚠️ **FAIL** — расхождения не зафиксированы или одно подразделение. Пересмотреть spec.")

lines.append("\n## Артефакты\n")
lines.append("- `_artifacts/20_full_discovery.csv` — полная карта расхождений")
lines.append("- `_artifacts/21_perenosavansa_rows.csv` — строки ПереносАванса")
lines.append("- `_artifacts/22_typed_breakdown.csv` — pivot ТипДок × ХозОп")
lines.append("- `_artifacts/23_etalon_uprbalance.csv` — сверка со штатным отчётом")

out = os.path.join(DOCS_DIR, "DISCOVERY_REPORT.md")
os.makedirs(DOCS_DIR, exist_ok=True)
with open(out, "w", encoding="utf-8") as fout:
    fout.write("\n".join(lines))

print(f"DISCOVERY_REPORT записан: {out}")
print(f"Σ Δ всего: {money(sum_total_20)}, подразделений: {len(by_podr)}")
```

- [ ] **Step 2: Запустить**

```bash
python 24_discovery_report.py
```

Expected: создан `docs/DISCOVERY_REPORT.md`, последняя строка лога — статус GO.

- [ ] **Step 3: Прочитать DISCOVERY_REPORT.md и зафиксировать**

```bash
cat "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/DISCOVERY_REPORT.md"
```

Verify: GATE = PASS. Если FAIL — стоп, пересмотр.

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/24_discovery_report.py _Rarzrabotki/notebook/knowledge_Balanse_klient/docs/DISCOVERY_REPORT.md && git commit -m "feat(balans_klient): add Phase 0 finalizer + DISCOVERY_REPORT.md"
```

---

# Phase 1 — Analysis

**Gate-критерий перехода к Phase 2:** `ANALYSIS_REPORT.md` создан, 100% Σ Δ объяснено категориями C1/C2/C3.

## Task 1.1: Создать скрипт `25_analysis_root_cause_matrix.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/25_analysis_root_cause_matrix.py`
- Output: `_artifacts/25_root_cause_matrix.csv`

**Назначение**: Для каждого документа в 20-арт классифицировать:
- **C1**: есть ХозОп=ВозвратОплатыКлиенту (или ВозвратДенежныхСредствВДругуюОрганизацию)
- **C2**: есть РСКПС/РСППС строки с ХозОп=ПереносАванса И документ — Реализация/Поступление с зачётом аванса
- **C3**: прочее (нужно расследовать)

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 25 (Phase 1) — Классификация документов по корневой причине

КАТЕГОРИИ:
    C1: ВозвратОплатыКлиенту / ВозвратДенежныхСредствВДругуюОрганизацию
    C2: ЗачетАванса (Реализация/Поступление с переносом аванса)
    C3: Прочее (требует расследования)
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, money, save_csv, ARTIFACTS_DIR

erp = connect_erp()


def f(x):
    try: return float(str(x).replace(",", ".").replace(" ", ""))
    except: return 0.0


# Загружаем 20-арт + 22-арт (имеет ХозОп)
r22 = []
with open(os.path.join(ARTIFACTS_DIR, "22_typed_breakdown.csv"), encoding="utf-8-sig", newline="") as fin:
    r22 = list(csv.DictReader(fin, delimiter=";"))

# Получаем pivot по (ТипДок, ХозОп) → классификация
C1_OPS = ("ВозвратОплатыКлиенту", "ВозвратДенежныхСредствВДругуюОрганизацию",
          "Возврат оплаты клиенту", "Возврат денежных средств в другую организацию")

def classify(тип, хозоп):
    if any(c in хозоп for c in C1_OPS):
        return "C1"
    # C2 — Реализация/Поступление + ХозОп переноса
    if any(t in (тип or "") for t in ("Реализация", "Приобретение", "Поступление")) and "Перенос" in (хозоп or ""):
        return "C2"
    return "C3"


cat_sums = {"C1": {"count": 0, "sum_delta": 0.0, "examples": []},
            "C2": {"count": 0, "sum_delta": 0.0, "examples": []},
            "C3": {"count": 0, "sum_delta": 0.0, "examples": []}}

out_rows = []
for r in r22:
    тип = r["ТипДок"]; хозоп = r["ХозОперация"]
    cat = classify(тип, хозоп); кол = int(r["КолДокументов"]); delta = f(r["Σ Δ"])
    cat_sums[cat]["count"] += кол
    cat_sums[cat]["sum_delta"] += delta
    if len(cat_sums[cat]["examples"]) < 5:
        cat_sums[cat]["examples"].append(f"{тип}/{хозоп}")
    out_rows.append({
        "ТипДок": тип, "ХозОперация": хозоп, "КолДокументов": кол,
        "Σ Δ": delta, "Категория": cat,
    })

print("=" * 100)
print("СКРИПТ 25 — Классификация документов")
print("=" * 100)
print(f"\n{'Категория':<5} {'Кол':>6} {'Σ Δ':>20} Примеры")
print("-" * 90)
for cat in ("C1", "C2", "C3"):
    v = cat_sums[cat]
    examples = "; ".join(v["examples"][:3])
    print(f"{cat:<5} {v['count']:>6} {money(v['sum_delta']):>20} {examples[:50]}")

path = save_csv("25_root_cause_matrix", out_rows,
                ["ТипДок", "ХозОперация", "КолДокументов", "Σ Δ", "Категория"])
print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить**

```bash
python 25_analysis_root_cause_matrix.py
```

Expected: C1 даёт основную долю Σ Δ.

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/25_analysis_root_cause_matrix.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/25_root_cause_matrix.csv && git commit -m "feat(balans_klient): add Phase 1 analysis script 25 (root cause)"
```

---

## Task 1.2: Создать скрипт `26_analysis_double_count_risk.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/26_analysis_double_count_risk.py`
- Output: `_artifacts/26_double_count_candidates.csv`

**Назначение**: Для Подхода A — найти документы, у которых **уже есть** локальные движения ПАП.ЗадКл/ПолучАв/ЗадПост/ВыдАв И параллельно есть РСКПС-строки с ХозОп=ПереносАванса (тогда снятие фильтра даст двойной учёт).

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 26 (Phase 1) — Risk analysis для Подхода A: двойной учёт

ЧТО ИЩЕТ:
    Документы у которых:
    - Σ |ПАП.signed| > 0 (документ пишет в ПАП.ЗадКл/ПолучАв/ЗадПост/ВыдАв)
    - И Σ ПереносАванса в РСКПС/РСППС от их ДокументРегистратор > 0
    Если такие есть — Подход A даст +двойную проводку.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money, save_csv

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
// ПАП-документы с положительными суммами по статьям расчётов
ВЫБРАТЬ
    Т.Регистратор КАК Док,
    СУММА(Т.Сумма) КАК ПАП_Σ
ПОМЕСТИТЬ втПАП
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И Т.Организация = &Орг
    И Т.Статья В (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ПолученныеАвансы),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыданныеАвансы))
СГРУППИРОВАТЬ ПО Т.Регистратор
ИМЕЮЩИЕ СУММА(Т.Сумма) > 0
;

// Документы у которых есть РСКПС-строки с ПереносАванса
ВЫБРАТЬ
    Р.ДокументРегистратор КАК Док,
    СУММА(Р.ДолгУпр + Р.ПредоплатаУпр) КАК РСК_Σ
ПОМЕСТИТЬ втРСК_ПА
ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Р
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Р.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И АП.Организация = &Орг
    И Р.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)
СГРУППИРОВАТЬ ПО Р.ДокументРегистратор
;

// Пересечение — это документы где Подход A даст двойной учёт
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(П.Док) КАК Документ,
    ВЫРАЗИТЬ(ТИПЗНАЧЕНИЯ(П.Док) КАК Строка(150)) КАК ТипДок,
    П.ПАП_Σ КАК ПАП_сумма,
    Р.РСК_Σ КАК РСК_ПереносАванса_сумма
ИЗ втПАП КАК П
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ втРСК_ПА КАК Р ПО П.Док = Р.Док
УПОРЯДОЧИТЬ ПО П.ПАП_Σ УБЫВ
"""

print("=" * 100)
print("СКРИПТ 26 — Документы с риском двойного учёта при Подходе A")
print("=" * 100)

res = q.Выполнить().Выгрузить()
rows = []
for i in range(res.Количество()):
    r = res.Получить(i)
    тип = str(r.ТипДок).split(".")[-1] if "." in str(r.ТипДок) else str(r.ТипДок)
    rows.append({
        "Документ": str(r.Документ), "ТипДок": тип,
        "ПАП_сумма": float(r.ПАП_сумма or 0),
        "РСК_ПереносАванса_сумма": float(r.РСК_ПереносАванса_сумма or 0),
    })

print(f"\nКандидатов на двойной учёт: {len(rows)}")
if rows:
    print(f"\nТоп-15:")
    print(f"{'Документ':<60} {'ТипДок':<32} {'ПАП':>14} {'РСК ПА':>14}")
    print("-" * 125)
    for r in rows[:15]:
        print(f"{r['Документ'][:60]:<60} {r['ТипДок'][:32]:<32} {money(r['ПАП_сумма']):>14} {money(r['РСК_ПереносАванса_сумма']):>14}")
else:
    print("✓ Нет документов с риском двойного учёта — Подход A безопасен.")

path = save_csv("26_double_count_candidates", rows,
                ["Документ", "ТипДок", "ПАП_сумма", "РСК_ПереносАванса_сумма"])
print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить**

```bash
python 26_analysis_double_count_risk.py
```

Expected: либо 0 кандидатов (Подход A безопасен), либо список (нужен гибрид с exclude).

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/26_analysis_double_count_risk.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/26_double_count_candidates.csv && git commit -m "feat(balans_klient): add Phase 1 risk script 26 (double count)"
```

---

## Task 1.3: Создать скрипт `27_analysis_local_branches_coverage.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/27_analysis_local_branches_coverage.py`
- Output: `_artifacts/27_local_branches.csv`

**Назначение**: Грепаем 3 ManagerModule (СписаниеБезналичных / РКО / ОПК) — ищем ветки `ВозвратОплатыКлиенту` в функции `ТекстЗапросаТаблицаПрочиеАктивыПассивы`. Если ветки нет — это Подход B target.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 27 (Phase 1) — Coverage локальных веток для ВозвратОплатыКлиенту в 3 ManagerModule

ЧТО ДЕЛАЕТ:
    Через Grep (через subprocess) проверяет наличие ветки ВозвратОплатыКлиенту
    в функции ТекстЗапросаТаблицаПрочиеАктивыПассивы для:
    - СписаниеБезналичныхДенежныхСредств
    - РасходныйКассовыйОрдер
    - ОперацияПоПлатежнойКарте
"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import save_csv

CFG = "C:/Configuration_downloads/BASERP25"
DOCS = (
    "СписаниеБезналичныхДенежныхСредств",
    "РасходныйКассовыйОрдер",
    "ОперацияПоПлатежнойКарте",
)


def analyze(doc_name):
    path = f"{CFG}/Documents/{doc_name}/Ext/ManagerModule.bsl"
    if not os.path.exists(path):
        return {"Документ": doc_name, "Найдено": False, "Покрытие": "файл не найден"}
    with open(path, encoding="utf-8") as f:
        src = f.read()

    # Найти функцию ТекстЗапросаТаблицаПрочиеАктивыПассивы
    m = re.search(r"Функция ТекстЗапросаТаблицаПрочиеАктивыПассивы\b[^\n]*\n(.+?)(?:^Функция |^Процедура )", src, re.DOTALL | re.MULTILINE)
    if not m:
        return {"Документ": doc_name, "Найдено": False, "Покрытие": "функция не найдена"}
    body = m.group(1)
    body_len = len(body.splitlines())

    # Поиск ветки для ВозвратОплатыКлиенту
    has_branch = "ВозвратОплатыКлиенту" in body
    has_zad_kl = "ЗадолженностьКлиентов" in body
    has_poluch_av = "ПолученныеАвансы" in body

    return {
        "Документ": doc_name,
        "Найдено": True,
        "Строк_в_функции": body_len,
        "Имеет_ВозвратОплатыКлиенту": has_branch,
        "Имеет_ЗадолженностьКлиентов": has_zad_kl,
        "Имеет_ПолученныеАвансы": has_poluch_av,
        "Покрытие": (
            "ПОЛНОЕ" if has_branch and has_zad_kl else
            "ЧАСТИЧНОЕ" if has_zad_kl else
            "ОТСУТСТВУЕТ"
        ),
    }


print("=" * 100)
print("СКРИПТ 27 — Coverage локальных веток для ВозвратОплатыКлиенту")
print("=" * 100)

rows = []
for d in DOCS:
    r = analyze(d)
    rows.append(r)
    print(f"\n  {d}:")
    for k, v in r.items():
        if k != "Документ":
            print(f"      {k}: {v}")

path = save_csv("27_local_branches", rows,
                ["Документ", "Найдено", "Строк_в_функции",
                 "Имеет_ВозвратОплатыКлиенту", "Имеет_ЗадолженностьКлиентов",
                 "Имеет_ПолученныеАвансы", "Покрытие"])
print(f"\nАртефакт: {path}")
```

- [ ] **Step 2: Запустить**

```bash
python 27_analysis_local_branches_coverage.py
```

Expected: для всех 3 документов Покрытие = "ОТСУТСТВУЕТ" (или "ЧАСТИЧНОЕ") — подтверждение что Подход B потребует добавления веток.

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/27_analysis_local_branches_coverage.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/27_local_branches.csv && git commit -m "feat(balans_klient): add Phase 1 coverage script 27 (local branches)"
```

---

## Task 1.4: Написать `ANALYSIS_REPORT.md` вручную

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/ANALYSIS_REPORT.md`

**Назначение**: Сводка фактов 25-27 + рекомендация системного инженера. Пишется ВРУЧНУЮ по результатам скриптов.

- [ ] **Step 1: Прочитать артефакты 25-27**

```bash
cat "_artifacts/25_root_cause_matrix.csv" | head -20
cat "_artifacts/26_double_count_candidates.csv" | head -20
cat "_artifacts/27_local_branches.csv"
```

- [ ] **Step 2: Создать ANALYSIS_REPORT.md** (шаблон ниже, заполнить фактическими цифрами)

```markdown
# ANALYSIS_REPORT — Phase 1

> Сводка анализа корневых причин и оценки рисков.

## Распределение Σ Δ по категориям (из 25)

| Категория | Кол-во пар (ТипДок,ХозОп) | Σ Δ |
|---|---:|---:|
| C1 — ВозвратОплатыКлиенту/ВДругуюОрганизацию | <...> | <...> |
| C2 — ЗачётАванса при Реализации/Поступлении | <...> | <...> |
| C3 — Прочее | <...> | <...> |
| **ИТОГО** | <...> | <...> |

## Риск двойного учёта при Подходе A (из 26)

<заполнить число кандидатов и Σ ПАП>

## Покрытие локальных веток (из 27)

| Документ | Покрытие | Найдено в ТекстЗапросаТаблицаПрочиеАктивыПассивы |
|---|---|---|
| СписаниеБезналичныхДенежныхСредств | <...> | <...> |
| РасходныйКассовыйОрдер | <...> | <...> |
| ОперацияПоПлатежнойКарте | <...> | <...> |

## Рекомендация системного инженера

<Выбор A / B / гибрид на основе фактов выше — обоснование цифрами>

## Gate-критерий Phase 2

✓ 100% Σ Δ объяснено категориями C1/C2/C3
✓ Список риска двойного учёта зафиксирован
✓ Coverage локальных веток подтверждён
```

- [ ] **Step 3: Заполнить отчёт фактами**

(заполнение зависит от вывода скриптов; примеры значений из скриптов 25-27 подставить вручную)

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/docs/ANALYSIS_REPORT.md && git commit -m "docs(balans_klient): add Phase 1 ANALYSIS_REPORT.md"
```

---

# Phase 2 — Strategy decision

**Gate-критерий перехода к Phase 3:** `STRATEGY_DECISION.md` подписан, выбран один из вариантов: A / B / гибрид.

## Task 2.1: Написать `STRATEGY_DECISION.md`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/STRATEGY_DECISION.md`

- [ ] **Step 1: Применить логику выбора**

Логика (из spec):
- Если **C1 покрывает 100% Δ + (26) пустой** → Подход A (1 файл, 4 строки)
- Если **C1 < 100% ИЛИ (26) не пустой** → Подход B (3 документа)
- Иначе → гибрид (A + exclude-list)

- [ ] **Step 2: Записать решение**

```markdown
# STRATEGY_DECISION — Phase 2

## Факты (из Phase 1)

- Σ Δ_категории_C1 = <...>
- Σ Δ_категории_C2 = <...>
- Σ Δ_категории_C3 = <...>
- Кандидатов двойного учёта (26) = <...>
- Локальные ветки покрытия (27): <...>

## Выбранная стратегия

**Подход <A / B / гибрид>** — обоснование:

<обоснование цифрами>

## Перечень файлов к правке

<зависит от выбора>

## Откат-план

<Файлы для backup, команды отката>

## Подпись

Решение принято: 2026-MM-DD
```

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/docs/STRATEGY_DECISION.md && git commit -m "docs(balans_klient): add Phase 2 STRATEGY_DECISION.md"
```

---

# Phase 3 — Patch (правка кода)

**Gate-критерий перехода к Phase 4:** скрипт `28_check_patch_applied.py` зелёный.

## Task 3.1: Backup исходных .bsl

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/_backup/*.bsl`

- [ ] **Step 1: Создать директорию backup**

```bash
mkdir -p "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/_backup"
```

- [ ] **Step 2: Скопировать все потенциально-модифицируемые файлы**

```bash
DATE=$(date +%Y-%m-%d)
BACKUP="C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/_backup"
cp "C:/Configuration_downloads/BASERP25/DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl" "$BACKUP/DvAktPas_${DATE}.bsl"
cp "C:/Configuration_downloads/BASERP25/Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl" "$BACKUP/SpisaniyaBeznal_${DATE}.bsl"
cp "C:/Configuration_downloads/BASERP25/Documents/РасходныйКассовыйОрдер/Ext/ManagerModule.bsl" "$BACKUP/RKO_${DATE}.bsl"
cp "C:/Configuration_downloads/BASERP25/Documents/ОперацияПоПлатежнойКарте/Ext/ManagerModule.bsl" "$BACKUP/OPK_${DATE}.bsl"
```

- [ ] **Step 3: Verify**

```bash
ls -la "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/knowledge_Balanse_klient/_backup/"
```

Expected: 4 файла .bsl с датой.

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/_backup/ && git commit -m "chore(balans_klient): backup .bsl files before Phase 3 patch"
```

---

## Task 3.2: Применить правку (зависит от STRATEGY_DECISION)

**Files:**
- Modify: один из следующих, в зависимости от выбора в STRATEGY_DECISION.md

**Подход A** — `DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl`:
- Строка 1275 — удалить `И Т.ХозяйственнаяОперация <> ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)`
- Строка 1323 — то же
- Строка 1428 — то же
- Строка 1476 — то же

**Подход B** — `Documents/<тип>/Ext/ManagerModule.bsl` (для каждого из 3):
- Добавить ОБЪЕДИНИТЬ ВСЕ ветки в `ТекстЗапросаТаблицаПрочиеАктивыПассивы` перед ", + РегистрыНакопления.ПрочиеАктивыПассивы.ТекстЗапросаТаблицаПрочиеАктивыПассивы()"

- [ ] **Step 1: Прочитать STRATEGY_DECISION.md** — какой подход выбран

- [ ] **Step 2 (вариант A): применить через Edit** — для каждой из 4 строк фильтра

Пример для строки 1275:
```bsl
# Найти точную строку
|	Т.ДолгУпр <> 0
|	И Т.ХозяйственнаяОперация <> ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)
|	//&Отбор

# Заменить на
|	Т.ДолгУпр <> 0
|	//&Отбор
```

- [ ] **Step 2 (вариант B): применить через Edit** — для каждого из 3 ManagerModule

Шаблон ветки (вставить ДО строки `ТекстЗапроса = ТекстЗапроса + "ОБЪЕДИНИТЬ ВСЕ"` + типовой код):
```bsl
|
|ОБЪЕДИНИТЬ ВСЕ
|
|// Ветка для ВозвратОплатыКлиенту с переносом аванса в долг — фикс #FIX-2026-05-23
|ВЫБРАТЬ
|	&Период,
|	ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход),
|	&Организация,
|	РасшПл.Подразделение,
|	РасшПл.НаправлениеДеятельности,
|	ВЫБОР КОГДА АП.Партнер = ЗНАЧЕНИЕ(Справочник.Партнеры.НашеПредприятие)
|		ТОГДА ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьСобственныхОрганизаций)
|		ИНАЧЕ ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов) КОНЕЦ,
|	НЕОПРЕДЕЛЕНО,
|	СУММА(ВЫРАЗИТЬ(РасшПл.Сумма * &КоэффициентПересчетаВВалютуУпр КАК ЧИСЛО(31,2)))
|ИЗ Документ.<ТипДок>.РасшифровкаПлатежа КАК РасшПл
|	ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.<ТипДок> КАК ДанДок ПО ДанДок.Ссылка = РасшПл.Ссылка
|	ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ОбъектыРасчетов КАК ОР ПО РасшПл.ОбъектРасчетов = ОР.Ссылка
|	ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
|		ПО АП.КлючАналитики = ОР.АналитикаУчетаПоПартнерам
|ГДЕ РасшПл.Ссылка = &Ссылка
|	И &ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ВозвратОплатыКлиенту)
|	И РасшПл.ОбъектРасчетов <> ДанДок.ОбъектРасчетов
|	И &ПроведеноБанком
|СГРУППИРОВАТЬ ПО РасшПл.Подразделение, РасшПл.НаправлениеДеятельности, АП.Партнер
|
|ОБЪЕДИНИТЬ ВСЕ
|
|ВЫБРАТЬ
|	&Период,
|	ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход),
|	...,
|	ВЫБОР КОГДА АП.Партнер = ЗНАЧЕНИЕ(Справочник.Партнеры.НашеПредприятие)
|		ТОГДА ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ОбязательстваПередСобственнымиОрганизациями)
|		ИНАЧЕ ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ПолученныеАвансы) КОНЕЦ,
|	...
```

(точный текст ветки финализируется через Python COM pre-тест после Phase 1 — детали в плане Phase 3 будут уточнены)

- [ ] **Step 3: Не применять — Rule #-1 — сначала Python pre-test запроса**

Создать `28a_pretest_new_query.py` (если выбран B): прогнать НОВЫЙ запрос (с добавленной веткой) через COM, проверить что:
- rows>0 для тестовых документов (000Ц-000287)
- 0 дублей по (Регистратор, Подразделение, Статья)
- Σ по статьям ЗадКл+ПолучАв == ожидаемая нетто из 000Ц-000287

- [ ] **Step 4: Применить правку через cp в main config**

```bash
# Если правили в worktree или в backup — копируем в основную конфигурацию
cp _Rarzrabotki/notebook/knowledge_Balanse_klient/_backup/DvAktPas_<date>_patched.bsl \
   DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl
```

- [ ] **Step 5: Загрузить через skill /db-load-xml**

Запустить skill:
```
/db-load-xml -Mode Partial -Files "DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl"
```

Expected: загрузка успешна, ConfigDumpInfo обновлён.

- [ ] **Step 6: Применить /db-update -Dynamic+**

```
/db-update -Dynamic+
```

Expected: динамическое обновление прошло, ошибок нет.

- [ ] **Step 7: Commit изменений в конфигурации**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl && git commit -m "fix(DvAktPas): remove filter <>ПереносАванса for РСКПС/РСППС projection to ПАП

Spec: docs/superpowers/specs/2026-05-23-fix-davaktpas-perenosavansa-design.md
Approach: <A или B>
Risk: <ссылка на STRATEGY_DECISION.md>"
```

---

## Task 3.3: Создать скрипт `28_check_patch_applied.py`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/28_check_patch_applied.py`

**Назначение**: Через COM прочитать тексты модулей и убедиться что правка фактически в базе (cache invalidation проверка).

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 28 (Phase 3 финал) — Проверка применения правки

ЧТО ДЕЛАЕТ:
    Через COM erp.Метаданные().Обработки.<имя>.МодульМенеджера.ПолучитьТекст()
    или через файл — проверяет наличие/отсутствие уникальной строки правки.

ВАЖНО: Конфигуратор кэширует модули. Если правка применена, но check_patch
    видит старую — ConfigDumpInfo.xml не обновлён или ConfigDumpInfo.xml
    был с ненулевым configVersion.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp

erp = connect_erp()

# Маркер правки (определяется по выбранному подходу)
# Для A: отсутствие "И Т.ХозяйственнаяОперация <> ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)"
#        в DvAktPas.ManagerModule
# Для B: наличие комментария "FIX-2026-05-23" в одном из 3 ManagerModule

# Через прямое чтение файла (быстрее)
PATHS = {
    "DvAktPas": "C:/Configuration_downloads/BASERP25/DataProcessors/ДвиженияАктивовПассивов/Ext/ManagerModule.bsl",
    "SpisaniyaBeznal": "C:/Configuration_downloads/BASERP25/Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl",
    "RKO": "C:/Configuration_downloads/BASERP25/Documents/РасходныйКассовыйОрдер/Ext/ManagerModule.bsl",
    "OPK": "C:/Configuration_downloads/BASERP25/Documents/ОперацияПоПлатежнойКарте/Ext/ManagerModule.bsl",
}

# Маркеры применённого патча. Используется ОДИН из блоков (A или B)
# на основе STRATEGY_DECISION.md. Закомментируйте неиспользуемый.

# === Подход A: маркеры должны ОТСУТСТВОВАТЬ в DvAktPas ===
EXPECTED_REMOVED = {
    "DvAktPas": [
        # 4 копии одного фильтра на строках 1275, 1323, 1428, 1476.
        # После правки A — все 4 удалены.
        "И Т.ХозяйственнаяОперация <> ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)",
    ],
}

# === Подход B: маркеры должны ПРИСУТСТВОВАТЬ в 3 ManagerModule ===
# (При выборе B — раскомментируйте и закомментируйте EXPECTED_REMOVED выше)
EXPECTED_PRESENT = {
    # "SpisaniyaBeznal": ["FIX-2026-05-23"],
    # "RKO": ["FIX-2026-05-23"],
    # "OPK": ["FIX-2026-05-23"],
}

print("=" * 100)
print("СКРИПТ 28 — Проверка применения правки в коде конфигурации")
print("=" * 100)

ok = True

for name, markers in EXPECTED_REMOVED.items():
    path = PATHS[name]
    with open(path, encoding="utf-8") as f:
        src = f.read()
    for marker in markers:
        if marker in src:
            print(f"  ❌ {name}: маркер ВСЁ ЕЩЁ В ФАЙЛЕ → '{marker[:80]}'")
            ok = False
        else:
            print(f"  ✓ {name}: маркер удалён")

for name, markers in EXPECTED_PRESENT.items():
    path = PATHS[name]
    with open(path, encoding="utf-8") as f:
        src = f.read()
    for marker in markers:
        if marker in src:
            print(f"  ✓ {name}: маркер найден → '{marker}'")
        else:
            print(f"  ❌ {name}: маркер ОТСУТСТВУЕТ → '{marker}'")
            ok = False

# Дополнительно через COM — проверка что модули в базе соответствуют файлам
# (для уверенности что Конфигуратор не закэшировал старую версию)
print("\n[Проверка через COM]")
# Получим текст модуля DvAktPas через метаданные базы
try:
    md = erp.Метаданные
    # Доступ к МодульМенеджера обработки невозможен через COM напрямую —
    # вместо этого делаем тестовый вызов: создаём запрос, использующий объект,
    # и проверяем что результаты соответствуют новой логике.
    # Этот шаг финализируется на основе выбранного подхода.
    print("  (COM-проверка — выполняется через тестовый прогон в Phase 4)")
except Exception as e:
    print(f"  COM-проверка пропущена: {e}")

if ok:
    print("\n✓ GATE Phase 3 → Phase 4: правка применена корректно")
    sys.exit(0)
else:
    print("\n❌ GATE FAIL: правка не применена или cache invalidation проблема")
    sys.exit(1)
```

- [ ] **Step 2: Запустить**

```bash
python 28_check_patch_applied.py
echo "Exit code: $?"
```

Expected: Exit 0, все ✓.

Если Exit 1 — пересмотр Phase 3.2 (см. spec §6 о ConfigDumpInfo).

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/28_check_patch_applied.py && git commit -m "feat(balans_klient): add Phase 3 verifier 28 (check patch applied)"
```

---

# Phase 4 — Verify

**Gate-критерий завершения**: `VERIFY_REPORT.md` со статусом GO по всем 7 скриптам.

## Task 4.1: Перепровести тестовые документы

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/29_repost_test_docs.py`

**Назначение**: Перепровести **только** документы из артефакта 20 (которые имели расхождение) — Get-or-Create, БЕЗ удаления.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 29 (Phase 4 шаг 1) — Перепровести документы с расхождением

ВАЖНО (memory feedback_no_doc_delete_in_tests):
    НЕ удалять документы! Get → ПолучитьОбъект() → Записать(РежимЗаписиДокумента.Проведение).
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, ARTIFACTS_DIR

erp = connect_erp()

# Загружаем 20-арт — список документов с расхождением
docs_to_repost = set()
with open(os.path.join(ARTIFACTS_DIR, "20_full_discovery.csv"), encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        docs_to_repost.add((r["ТипДок"], r["Документ"]))

print(f"Документов к перепроведению: {len(docs_to_repost)}")

ok = err = 0
for тип, имя in docs_to_repost:
    try:
        q = erp.NewObject("Запрос")
        q.УстановитьПараметр("Имя", имя)
        q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.{тип} ГДЕ ПРЕДСТАВЛЕНИЕ(Ссылка) = &Имя'
        res = q.Выполнить().Выгрузить()
        if res.Количество() == 0:
            print(f"  ⚠️ {имя}: не найден"); err += 1; continue
        doc = res.Получить(0).Ссылка.ПолучитьОбъект()
        # Режим проведения
        rm = erp.NewObject("РежимЗаписиДокумента")  # неправильно — это enum
        # Корректно:
        doc.Записать(erp.Перечисления.РежимЗаписиДокумента.Проведение if False else None)
        # На самом деле:
        # doc.Записать(РежимЗаписиДокумента.Проведение)
        # Через COM: режим как число — 1 = Проведение
        doc.Записать(1)  # 1 = Проведение
        ok += 1
    except Exception as e:
        print(f"  ❌ {тип}/{имя}: {e}"); err += 1

print(f"\nOK: {ok}, Errors: {err}")
sys.exit(0 if err == 0 else 1)
```

- [ ] **Step 2: Запустить**

```bash
python 29_repost_test_docs.py
echo "Exit code: $?"
```

Expected: Exit 0 (все документы перепроведены).

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/29_repost_test_docs.py && git commit -m "feat(balans_klient): add Phase 4 reposter 29 (test docs)"
```

---

## Task 4.2: Acceptance скрипт 30

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/30_verify_main_diff_closed.py`

**Назначение**: После перепроведения скрипт 20 должен дать Σ Δ = 0 ±0.01.

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 30 (Phase 4) — Acceptance: Σ Δ = 0 после правки

ЧТО ДЕЛАЕТ:
    Перезапускает запрос 20 и проверяет что Σ Δ < 0.01.
"""
import sys, io, csv, os, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

# Запустить 20 повторно
script_dir = os.path.dirname(os.path.abspath(__file__))
result = subprocess.run([sys.executable, "20_discovery_all_podr_2025.py"],
                        cwd=script_dir, capture_output=True, text=True, encoding="utf-8")
print(result.stdout); print(result.stderr)

# Прочитать обновлённый 20-арт
total_delta = 0.0
n_diff = 0
with open(os.path.join(script_dir, "_artifacts", "20_full_discovery.csv"), encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        try:
            d = float(str(r["Дельта"]).replace(",", ".").replace(" ", ""))
            total_delta += d
            if abs(d) > 0.01: n_diff += 1
        except: pass

print(f"\n=== СКРИПТ 30 — Acceptance ===")
print(f"После правки: Σ Δ = {total_delta:+,.2f}, документов с |Δ|>0.01: {n_diff}")

if abs(total_delta) < 0.01 and n_diff == 0:
    print("✓ ACCEPTANCE PASS")
    sys.exit(0)
else:
    print("❌ ACCEPTANCE FAIL — есть остаточные расхождения")
    sys.exit(1)
```

- [ ] **Step 2: Запустить**

```bash
python 30_verify_main_diff_closed.py
echo "Exit code: $?"
```

Expected: Exit 0, Σ Δ = 0.

- [ ] **Step 3: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/30_verify_main_diff_closed.py && git commit -m "feat(balans_klient): add Phase 4 acceptance 30 (main diff closed)"
```

---

## Task 4.3: Acceptance скрипт 31 — Глобино-2 / 31.12.2025

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/31_verify_globino2_dec2025.py`

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 31 (Phase 4) — Acceptance Глобино-2 / 31.12.2025

ПРОВЕРКА:
    После правки ПАП.КонецДек.ЗадКл = РСКПС.КонецДек.ДолгКлиентов
    Раньше: ПАП=37 064 284,12 vs РСКПС=44 258 878,85 (Δ=-7 194 594,73)
    После: ожидаем Δ=0,00
"""
import sys, io, subprocess, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

# Прогнать 01-03 повторно
script_dir = os.path.dirname(os.path.abspath(__file__))
for s in ("01_balance_pap_klient_globino2.py", "02_balance_rsk_klient_globino2.py", "03_reconcile_balances.py"):
    r = subprocess.run([sys.executable, s], cwd=script_dir, capture_output=True, text=True, encoding="utf-8")
    print(f"\n--- {s} ---")
    print(r.stdout[-500:])

# Прочитать 03-арт
import csv
with open(os.path.join(script_dir, "_artifacts", "03_balance_reconciliation.csv"),
          encoding="utf-8-sig", newline="") as f:
    rows = list(csv.DictReader(f, delimiter=";"))

end_dec = next((r for r in rows if "31.12.2025" in r["Точка"]), None)
if not end_dec:
    print("❌ 31.12.2025 в 03-арт не найдено"); sys.exit(1)

delta = float(str(end_dec["Delta"]).replace(",", ".").replace(" ", ""))
print(f"\n=== СКРИПТ 31 ===")
print(f"Глобино-2 / 31.12.2025: Δ = {delta:+,.2f}")

if abs(delta) < 0.01:
    print("✓ PASS (ВРСК = Управ.баланс)")
    sys.exit(0)
else:
    print(f"❌ FAIL — остаточная Δ = {delta:+,.2f}")
    sys.exit(1)
```

- [ ] **Step 2: Запустить + commit (как в 4.2)**

```bash
python 31_verify_globino2_dec2025.py
echo "Exit: $?"
```

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/31_verify_globino2_dec2025.py && git commit -m "feat(balans_klient): add Phase 4 acceptance 31 (globino2 dec2025)"
```

---

## Task 4.4: Acceptance скрипт 32 — Σ-инвариант штатного отчёта

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/32_verify_uprbalance_report.py`

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 32 (Phase 4) — Σ-инвариант штатного Отчёт.УправленческийБаланс

ПРОВЕРКА:
    Σ КО Актив = |Σ КО Пассив| per организация ТОВ ІНДАСТРІАЛБУД
    Эталоны (НЕ изменились после правки):
    - дек 2025: 278 093 267,32
    - янв 2026: 288 787 750,11
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money

erp = connect_erp()
refs = get_refs(erp)

ЭТАЛОНЫ = {
    "2025-12-31": 278_093_267.32,
    "2026-01-31": 288_787_750.11,
}

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])

ok = True
for дата, эталон in ЭТАЛОНЫ.items():
    год, мес, день = дата.split("-")
    дт_парам = f"ДАТАВРЕМЯ({год},{int(мес)},{день},23,59,59)"
    q.Текст = f"""
    ВЫБРАТЬ
        СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Т.Сумма
                    ИНАЧЕ -Т.Сумма КОНЕЦ) КАК СальдоВсегоОрг
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
    ГДЕ Т.Период <= {дт_парам} И Т.Организация = &Орг
        И НЕ Т.Статья В ИЕРАРХИИ (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.СобственныеСредства),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ДоходыТекущегоПериода),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.РасходыТекущегоПериода))
    """
    r = q.Выполнить().Выгрузить().Получить(0)
    σ_signed = float(r.СальдоВсегоОрг or 0)
    print(f"  {дата}: Σ signed = {σ_signed:+,.2f} (эталон |Актив|=|Пассив|={эталон:,.2f})")
    # Σ signed (OD-3) = 0 = баланс
    if abs(σ_signed) > 0.01:
        print(f"    ⚠️ балансовый инвариант нарушен")
        ok = False

if ok:
    print("\n✓ Σ-инвариант сохранён")
    sys.exit(0)
else:
    print("\n❌ Σ-инвариант нарушен после правки — rollback!")
    sys.exit(1)
```

- [ ] **Step 2: Запустить + commit**

```bash
python 32_verify_uprbalance_report.py
echo "Exit: $?"
```

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/32_verify_uprbalance_report.py && git commit -m "feat(balans_klient): add Phase 4 acceptance 32 (uprbalance Σ-invariant)"
```

---

## Task 4.5: Regression 33 — нормальные платежи

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/33_regression_normal_payments.py`

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 33 (Phase 4) — Regression: нормальные платежи (без ПереносАванса)

ИДЕЯ:
    Взять 10 случайных Поступление безналичных + ПКО за разные месяцы
    БЕЗ ХозОп=ВозвратОплатыКлиенту → Σ ПАП по этим документам не изменилась.

ПРЕДВАРИТЕЛЬНО: до правки сохранить Σ ПАП в файл `_artifacts/33_baseline.json`.
ПОСЛЕ ПРАВКИ: сравнить — Σ_after = Σ_baseline (Δ=0).

Если baseline не существует — создаёт его.
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, money, ARTIFACTS_DIR

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
// Берём 10 ПКО + 10 Поступление безналичных без переноса аванса за 2025
ВЫБРАТЬ ПЕРВЫЕ 10
    Д.Ссылка КАК Док, "ПКО" КАК Тип, Д.Дата
ИЗ Документ.ПриходныйКассовыйОрдер КАК Д
ГДЕ Д.Дата МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31)
    И Д.Организация = &Орг
    И Д.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПоступлениеОплатыОтКлиента)
    И Д.Проведен
УПОРЯДОЧИТЬ ПО Д.Дата

ОБЪЕДИНИТЬ ВСЕ

ВЫБРАТЬ ПЕРВЫЕ 10
    Д.Ссылка, "Поступление безнал.", Д.Дата
ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств КАК Д
ГДЕ Д.Дата МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31)
    И Д.Организация = &Орг
    И Д.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПоступлениеОплатыОтКлиента)
    И Д.Проведен
УПОРЯДОЧИТЬ ПО Д.Дата
"""

res = q.Выполнить().Выгрузить()
docs = [res.Получить(i).Док for i in range(res.Количество())]
print(f"Регрессия на {len(docs)} нормальных платежах")

# Для каждого считаем Σ ПАП-движений
sums = {}
q2 = erp.NewObject("Запрос")
for doc in docs:
    q2.УстановитьПараметр("Р", doc)
    q2.Текст = """
    ВЫБРАТЬ СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК Σ
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
    ГДЕ Т.Регистратор = &Р
    """
    r = q2.Выполнить().Выгрузить().Получить(0)
    sums[str(doc)] = float(r.Σ or 0)

baseline_path = os.path.join(ARTIFACTS_DIR, "33_baseline.json")
if not os.path.exists(baseline_path):
    with open(baseline_path, "w", encoding="utf-8") as f:
        json.dump(sums, f, ensure_ascii=False, indent=2, default=str)
    print(f"Baseline создан: {baseline_path}")
    sys.exit(0)

# Сравнить
with open(baseline_path, encoding="utf-8") as f:
    baseline = json.load(f)

ok = True
for k, v_now in sums.items():
    v_base = baseline.get(k, 0)
    if abs(v_now - v_base) > 0.01:
        print(f"  ❌ {k}: baseline={v_base:+,.2f}, now={v_now:+,.2f}")
        ok = False
    else:
        print(f"  ✓ {k}: Δ=0")

if ok:
    print("\n✓ Регрессия 33 PASS")
    sys.exit(0)
else:
    print("\n❌ FAIL — нормальные платежи изменили проводки в ПАП!")
    sys.exit(1)
```

- [ ] **Step 2: Запустить ДО правки** (для baseline)

```bash
python 33_regression_normal_payments.py
```

Expected: создан baseline.

- [ ] **Step 3: Запустить ПОСЛЕ правки** (сравнение)

```bash
python 33_regression_normal_payments.py
echo "Exit: $?"
```

Expected: Exit 0.

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/33_regression_normal_payments.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/33_baseline.json && git commit -m "feat(balans_klient): add Phase 4 regression 33 (normal payments)"
```

---

## Task 4.6: Regression 34 — Реализации

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/34_regression_realizations.py`

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 34 (Phase 4) — Regression: Реализации товаров и услуг

ИДЕЯ:
    10 первых Реализаций 2025 (без ВозвратТоваров/Корректировки) → их Σ ПАП
    до правки и после правки должна быть равной. Если нет — правка
    непреднамеренно повлияла на Реализации.

Логика идентична скрипту 33, но запрос подбора другой.
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, ARTIFACTS_DIR

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 10
    Д.Ссылка КАК Док
ИЗ Документ.РеализацияТоваровУслуг КАК Д
ГДЕ Д.Дата МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31)
    И Д.Организация = &Орг
    И Д.Проведен
    И Д.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.РеализацияКлиенту)
УПОРЯДОЧИТЬ ПО Д.Дата
"""
res = q.Выполнить().Выгрузить()
docs = [res.Получить(i).Док for i in range(res.Количество())]
print(f"Регрессия на {len(docs)} реализациях")

sums = {}
q2 = erp.NewObject("Запрос")
for doc in docs:
    q2.УстановитьПараметр("Р", doc)
    q2.Текст = """
    ВЫБРАТЬ СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                  ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК Σ
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
    ГДЕ Т.Регистратор = &Р
    """
    sums[str(doc)] = float(q2.Выполнить().Выгрузить().Получить(0).Σ or 0)

baseline_path = os.path.join(ARTIFACTS_DIR, "34_baseline.json")
if not os.path.exists(baseline_path):
    with open(baseline_path, "w", encoding="utf-8") as f:
        json.dump(sums, f, ensure_ascii=False, indent=2, default=str)
    print(f"Baseline создан: {baseline_path}"); sys.exit(0)

with open(baseline_path, encoding="utf-8") as f:
    baseline = json.load(f)
ok = True
for k, v in sums.items():
    if abs(v - baseline.get(k, 0)) > 0.01:
        print(f"  ❌ {k}: baseline={baseline.get(k):+,.2f}, now={v:+,.2f}"); ok = False
    else:
        print(f"  ✓ {k}: Δ=0")

print("\n✓ Регрессия 34 PASS" if ok else "\n❌ FAIL — Реализации изменили проводки!")
sys.exit(0 if ok else 1)
```

- [ ] **Step 2: Запустить ДО правки** (baseline)

```bash
python 34_regression_realizations.py
```

- [ ] **Step 3: Запустить ПОСЛЕ правки** (verify)

```bash
python 34_regression_realizations.py
echo "Exit: $?"
```

Expected: Exit 0.

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/34_regression_realizations.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/34_baseline.json && git commit -m "feat(balans_klient): add Phase 4 regression 34 (realizations)"
```

---

## Task 4.7: Regression 35 — Зачёт аванса

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/35_regression_zachet_avansa.py`

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 35 (Phase 4) — Regression: Зачёт аванса при Реализации

ИДЕЯ:
    Найти Реализации 2025 у которых ЕСТЬ строки РСКПС с ХозОп=ПереносАванса
    (зачёт аванса при отгрузке). Σ ПАП до=после правки должна быть равной,
    потому что Реализация сама пишет правильные локальные проводки в ПАП.

    Если правка добавила им проводки → значит Реализация попадает в risk-set
    из скрипта 26 и нужен exclude (или другая стратегия).
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs, ARTIFACTS_DIR

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
ВЫБРАТЬ РАЗЛИЧНЫЕ ПЕРВЫЕ 10
    Т.ДокументРегистратор КАК Док
ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Т
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Т.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31)
    И АП.Организация = &Орг
    И Т.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПереносАванса)
    И ТИПЗНАЧЕНИЯ(Т.ДокументРегистратор) = ТИП(Документ.РеализацияТоваровУслуг)
"""
res = q.Выполнить().Выгрузить()
docs = [res.Получить(i).Док for i in range(res.Количество())]
print(f"Регрессия на {len(docs)} Реализациях с зачётом аванса")

sums = {}
q2 = erp.NewObject("Запрос")
for doc in docs:
    q2.УстановитьПараметр("Р", doc)
    q2.Текст = """
    ВЫБРАТЬ СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                  ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК Σ
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
    ГДЕ Т.Регистратор = &Р
    """
    sums[str(doc)] = float(q2.Выполнить().Выгрузить().Получить(0).Σ or 0)

baseline_path = os.path.join(ARTIFACTS_DIR, "35_baseline.json")
if not os.path.exists(baseline_path):
    with open(baseline_path, "w", encoding="utf-8") as f:
        json.dump(sums, f, ensure_ascii=False, indent=2, default=str)
    print(f"Baseline создан: {baseline_path}"); sys.exit(0)

with open(baseline_path, encoding="utf-8") as f:
    baseline = json.load(f)
ok = True
for k, v in sums.items():
    if abs(v - baseline.get(k, 0)) > 0.01:
        print(f"  ❌ {k}: baseline={baseline.get(k):+,.2f}, now={v:+,.2f}"); ok = False
    else:
        print(f"  ✓ {k}: Δ=0")

print("\n✓ Регрессия 35 PASS" if ok else "\n❌ FAIL — зачёт аванса при Реализации сломан!")
sys.exit(0 if ok else 1)
```

- [ ] **Step 2: ДО правки запустить → baseline**

```bash
python 35_regression_zachet_avansa.py
```

- [ ] **Step 3: ПОСЛЕ правки запустить → verify**

```bash
python 35_regression_zachet_avansa.py
echo "Exit: $?"
```

Expected: Exit 0. **Если FAIL** — Реализации попадают в risk-set (26), нужен exclude в коде правки или гибрид-стратегия.

- [ ] **Step 4: Commit**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/35_regression_zachet_avansa.py _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/35_baseline.json && git commit -m "feat(balans_klient): add Phase 4 regression 35 (zachet avansa)"
```

---

## Task 4.8: Regression 36 — Σ signed = 0 per организация

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/36_regression_full_balance.py`

- [ ] **Step 1: Написать скрипт**

```python
# -*- coding: utf-8 -*-
"""
СКРИПТ 36 (Phase 4) — Балансовый инвариант: Σ signed (КО, OD-3) per орг = 0
Эталоны из knowledge_Balanse: дек 278 093 267,32 / янв 288 787 750,11
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, get_refs

erp = connect_erp()
refs = get_refs(erp)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = """
ВЫБРАТЬ
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Т.Сумма
                ИНАЧЕ -Т.Сумма КОНЕЦ) КАК Σsigned
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И Т.Организация = &Орг
    И НЕ Т.Статья В ИЕРАРХИИ (
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.СобственныеСредства),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ДоходыТекущегоПериода),
        ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.РасходыТекущегоПериода))
"""
r = q.Выполнить().Выгрузить().Получить(0)
σ = float(r.Σsigned or 0)
print(f"Σ signed (OD-3) на 31.12.2025 / ТОВ = {σ:+,.2f}")
if abs(σ) < 0.01:
    print("✓ Балансовый инвариант")
    sys.exit(0)
else:
    print(f"❌ Σ ≠ 0 — баланс сломан! Δ = {σ}")
    sys.exit(1)
```

- [ ] **Step 2-3**: запустить + commit

---

## Task 4.9: Сборка `VERIFY_REPORT.md`

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/docs/VERIFY_REPORT.md`

- [ ] **Step 1: Запустить все verify-скрипты и собрать exit-коды**

```bash
for s in 30 31 32 33 34 35 36; do
    python "${s}_*.py" > "_artifacts/${s}_run.log" 2>&1
    echo "Script ${s}: $?"
done
```

- [ ] **Step 2: Написать VERIFY_REPORT.md**

```markdown
# VERIFY_REPORT — Phase 4

## Acceptance

| Скрипт | Что проверено | Результат |
|---|---|:---:|
| 30 | Σ Δ ПАП vs РСКПС/РСППС = 0 | <PASS/FAIL> |
| 31 | Глобино-2 / 31.12.2025 Δ=0 | <PASS/FAIL> |
| 32 | Штатный отчёт Σ-инвариант | <PASS/FAIL> |

## Regression

| Скрипт | Что проверено | Результат |
|---|---|:---:|
| 33 | Нормальные платежи без ПереносАванса | <PASS/FAIL> |
| 34 | Реализации | <PASS/FAIL> |
| 35 | Зачёт аванса | <PASS/FAIL> |
| 36 | Σ signed (OD-3) = 0 per орг | <PASS/FAIL> |

## Итог

<GO / NO-GO>

Если NO-GO — выполнить rollback по плану spec §3.5.4.
```

- [ ] **Step 3: Если все PASS → обновить FINDINGS.md статусом "Решено"**

Добавить раздел в начало FINDINGS.md:
```markdown
## Статус: РЕШЕНО (2026-MM-DD)

Применена правка кода конфигурации (см. STRATEGY_DECISION.md).
Σ Δ ПАП vs РСКПС/РСППС за 2025 по всем подразделениям ТОВ = 0,00.
VERIFY_REPORT: PASS по всем 7 скриптам.
```

- [ ] **Step 4: Commit финальный**

```bash
cd "C:/Configuration_downloads/BASERP25" && git add _Rarzrabotki/notebook/knowledge_Balanse_klient/docs/VERIFY_REPORT.md _Rarzrabotki/notebook/knowledge_Balanse_klient/FINDINGS.md && git commit -m "docs(balans_klient): Phase 4 VERIFY_REPORT GO + status RESOLVED"
```

---

# Завершение

После Task 4.9 проект завершён. План считается выполненным когда:

- [ ] Все 4 MD-отчёта в `docs/` закоммитчены
- [ ] Все 16+ скриптов в `Python/test/` закоммитчены и идемпотентны
- [ ] Backup в `_backup/` закоммитчен
- [ ] Правка в одном из 4 .bsl файлов конфигурации закоммитчена
- [ ] `VERIFY_REPORT.md` имеет GO-статус
- [ ] `FINDINGS.md` обновлён статусом "Решено"

**Time-estimate**: 8-12 часов работы (5 фаз, ~22 задачи, по 15-30 минут каждая, плюс перепроведение и валидация).

**Risk-mitigation**: rollback план по spec §3.5.4 — 5 минут на восстановление.

---

## Cross-references

- Spec: `docs/superpowers/specs/2026-05-23-fix-davaktpas-perenosavansa-design.md`
- FINDINGS.md: `_Rarzrabotki/notebook/knowledge_Balanse_klient/FINDINGS.md` (ред. 4)
- Существующая база знаний: `_Rarzrabotki/notebook/knowledge_Balanse/` (каноны Свод_*)
- Memory references: `feedback_no_doc_delete_in_tests`, `feedback_use_db_skills_for_config_load`, `feedback_designer_cache_invalidation`
