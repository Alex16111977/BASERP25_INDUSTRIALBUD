# knowledge_А_ОтражениеЗПпоКазне

База знаний по документу `Документ.А_ОтражениеЗПпоКазне` (BAS ERP 2.5 INDUSTRIALBUD).

## Статус

🟢 **База знаний создана 2026-05-26.** 10 файлов knowledge + 4 Python COM скрипта + 4 артефакта _artifacts/.

## Точка входа

**👉 [KNOWLEDGE_MAP.md](KNOWLEDGE_MAP.md)** — главный entry point для ИИ-сессий и аналитиков.

## Краткое описание

Документ-агрегатор, тянет данные ЗП из **3-х внешних баз 1С**:
- **BuhKazn** (Казна) — РегНакопл.БДДС
- **zup_1** (BAS ЗУП регламент) — НДФЛ/ВС/ЕСВ + исп.листы
- **zup_2** (BAS ЗУП управленческий) — НачислениеЗарплатыРаботникам

Использует **cross-base UUID 1:1** для связи документов Казна↔ERP. Заполняет 7 ТЧ и создаёт 2 подчинённых документа: `НачислениеЗарплаты` (ЗУП-аналог) + `ОтражениеЗарплатыВФинансовомУчете`.

⚠️ **Поправка к [PROMPT.md](PROMPT.md):** реальный код подключается к **zup_1** + **zup_2** (BAS ЗУП), а НЕ к BuhBud (BAS Бухгалтерія). Детали в [KNOWLEDGE_MAP.md](KNOWLEDGE_MAP.md) и [07_integracii.md](07_integracii.md).

## Структура

```
knowledge_А_ОтражениеЗПпоКазне/
├── PROMPT.md                       (исходный промт-задание)
├── README.md                       (этот файл)
├── KNOWLEDGE_MAP.md                ★ entry point
├── 01_dokument_struktura.md        реквизиты + 7 ТЧ + 32 процедуры
├── 02_zagruzka_iz_kazni.md         блок 1: Казна + UUID lookup
├── 03_zagruzka_iz_buh.md           блок 2: zup_1 (НДФЛ/ЕСВ/удержания)
├── 04_zapolnenie_nachisleniy.md    блок 3: zup_2 + 5 _ОтрЗП_*
├── 05_sozdanie_dok_nachislenie.md  блок 4а: НачислениеЗарплаты ЗУП
├── 06_sozdanie_dok_otrazhenie.md   блок 4б: ОтражениеЗП в Фин.учёте
├── 07_integracii.md                сводка 3 баз + UUID 1:1
├── LESSONS.md                      15 антипаттернов
├── FINDINGS.md                     эталоны discovery (UUID, размеры ТЧ)
└── Python/
    ├── README.md
    ├── test/
    │   ├── discovery_kazna_bdds_struktura.py
    │   ├── discovery_zup_struktura.py
    │   ├── verify_uuid_lookup_kazna_erp.py
    │   └── verify_test_doc_loading.py
    └── _artifacts/
        ├── kazna_bdds_schema.md
        ├── zup_schema.md
        ├── uuid_pair_example.md
        └── test_doc_sample.md
```

## Образцы аналогичных баз

- [`knowledge_НеоборотныеАктивы/`](../knowledge_НеоборотныеАктивы/) — компактный образец (10 файлов)
- [`knowledge_Balanse/`](../knowledge_Balanse/) — расширенный (~10 файлов)
- [`knowledge_PL/`](../knowledge_PL/) — справочный

## Связанные базы знаний

- [[knowledge_Balanse]] — управленческий баланс (Оплата труда уходит в свод)
- [[knowledge_PL]] — статьи PL (ФОТ/ЕСВ)
- [[knowledge_Balanse_money]] — поток денег ЗП в свод
