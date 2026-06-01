# knowledge_А_ОтражениеЗПпоКазне

База знаний по документу `Документ.А_ОтражениеЗПпоКазне` (BAS ERP 2.5 INDUSTRIALBUD).

## Статус

🟢 **База знаний создана 2026-05-26.** 13 файлов knowledge + Python COM скрипты + артефакты _artifacts/.

🟢 **2026-05-27** — добавлены файлы 11, 12, 13 (реализация ВКассу из РаспределениеФ2 → кнопка ОтражениеЗПпоКазне → подформа редактирования А_Расшифровки → унификация ТЧ.А_Расшифровка с ТЧ.Зарплата). 8 Python COM тестов PASS.

🟢 **2026-05-29…06-01** — файлы 14–18: РКО-взаиморасчёты из ВКассу (14), гард А_Необновлять (15), `А_ПередачаНачисленийМеждуПодразделениями` (16), `НачисленияБух`+разделение `ФормаPL` Ф1/Ф2 (17), фикс −0.01 в `А_ВзаиморасчетыССотрудниками` (18). Текущий entry-point со статусами — [KNOWLEDGE_MAP.md](KNOWLEDGE_MAP.md).

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
├── 01_dokument_struktura.md        реквизиты + 7 ТЧ (теперь 8 — добавлена А_Расшифровка) + процедуры
├── 02_zagruzka_iz_kazni.md         блок 1: Казна + UUID lookup
├── 03_zagruzka_iz_buh.md           блок 2: zup_1 (НДФЛ/ЕСВ/удержания)
├── 04_zapolnenie_nachisleniy.md    блок 3: zup_2 + 5 _ОтрЗП_*
├── 05_sozdanie_dok_nachislenie.md  блок 4а: НачислениеЗарплаты ЗУП
├── 06_sozdanie_dok_otrazhenie.md   блок 4б: ОтражениеЗП в Фин.учёте
├── 07_integracii.md                сводка 3 баз + UUID 1:1
├── 08_doc_raspredelenie_zp.md      структура А_РаспределениеЗП / РаспределениеФ2
├── 09_vedomost_vyplaty.md          структура ВКассу + регистр ЗарплатаКВыплате
├── 10_payment_flow.md              полный поток выплаты Казна → ERP → Ведомости
├── 11_vedomost_iz_raspredeleniyaF2.md  🟢 РЕАЛИЗОВАНО — ВКассу из Ф2 (8 уроков, эталон №000000026, Σ=348 800)
├── 12_button_iz_otrazheniya.md         🟢 РЕАЛИЗОВАНО 2026-05-27 — кнопка ОтражениеЗПпоКазне → N×ВКассу
├── 13_vkassu_podforma_rashifrovka.md   🟢 РЕАЛИЗОВАНО 2026-05-27 — подформа А_ФормаСпискаРасшифровкиПоФЛ (13 колонок)
├── 14_rko_vzaimoraschety_iz_vkassu.md  🟢 2026-05-27 — А_ВзСС в РКО из ВКассу в разрезе подразделений
├── 15_vkassu_skip_neobnovlyat.md       🟢 2026-05-29 — гард А_Необновлять (ВКассу занятая РКО не перезаписывается)
├── 16_peredacha_nachisleniy.md         🟢 2026-05-29/30 — А_ПередачаНачисленийМеждуПодразделениями (перенос ЗП/налогов на целевые подр)
├── 17_formapl_split.md                 🟢 2026-05-31 — НачисленияБух (gross 661) + разделение ФормаPL (Ф1 бух / Ф2 управ)
├── 18_vzss_uderzhanie_ndfl_align.md    🟢 2026-06-01 — фикс −0.01 ВзСС Ф1 (выравнивание «Удержание» ↔ НачисленныйНДФЛ; перенос до копейки)
├── LESSONS.md                      21 антипаттерн
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

## Реализация ВКассу — связанные spec/plan

| Этап | Spec | Plan | Knowledge |
|---|---|---|---|
| ВКассу из РаспределениеФ2 (hotfix v2 + v3 проведение) | `docs/superpowers/specs/2026-05-26-vedomost-vkassu-iz-raspredeleniyaf2-design.md` | `docs/superpowers/plans/2026-05-26-vedomost-vkassu-iz-raspredeleniyaf2-plan.md` | [11](11_vedomost_iz_raspredeleniyaF2.md) |
| Кнопка ОтражениеЗПпоКазне → N×ВКассу | `2026-05-27-otrazhenie-zp-vkassu-design.md` | `2026-05-27-otrazhenie-zp-vkassu-plan.md` | [12](12_button_iz_otrazheniya.md) |
| Подформа редактирования А_Расшифровки | `2026-05-27-vkassu-podforma-rashifrovka-design.md` | `2026-05-27-vkassu-podforma-rashifrovka-plan.md` | [13](13_vkassu_podforma_rashifrovka.md) |
| Унификация ТЧ.А_Расшифровка с ТЧ.Зарплата (13 полей) | `2026-05-27-rashifrovka-unify-with-zarplata-design.md` | `2026-05-27-rashifrovka-unify-with-zarplata-plan.md` | [11 Урок №8](11_vedomost_iz_raspredeleniyaF2.md), [13 Урок №7](13_vkassu_podforma_rashifrovka.md) |

## Python COM тесты (`_Rarzrabotki/Python/test/`)

| Тест | Что проверяет | Эталон |
|---|---|---|
| `test_create_vedomost_vkassu_iz_f2_smoke.py` | Создание ВКассу из Ф2 → шапка + проведение | Ф2 №000000026 |
| `test_create_vedomost_vkassu_iz_f2_idempotency.py` | UUID не меняется при повторных перепроведениях | Ф2 №000000026 |
| `test_create_vedomost_vkassu_iz_f2_rashifrovka.py` | Зеркало 1:1 ТЧ.Зарплата ↔ А_Расшифровка, Σ КВыплате совпадает | Ф2 №000000026 (Σ=348 800) |
| `test_otrazhenie_vkassu_smoke.py` | Кнопка ОтражениеЗПпоКазне → 7 ВКассу | №000000006 |
| `test_otrazhenie_vkassu_idempotency.py` | Повторный вызов кнопки не плодит дубли | №000000006 |
| `test_otrazhenie_vkassu_statyaddc.py` | СтатьяДДС в А_Расшифровке = ТЧ.РаспределениеКазна | №000000006 |
| `test_vkassu_subform_metadata.py` | Подформа зарегистрирована, ТЧ.А_Расшифровка 13 реквизитов | ВКассу 000Ц-* |
| `test_vkassu_subform_returns_changes.py` | Паттерн возврата: удалить ФЛ → добавить из подформы | ВКассу 000Ц-000005 |

## Образцы аналогичных баз

- [`knowledge_НеоборотныеАктивы/`](../knowledge_НеоборотныеАктивы/) — компактный образец (10 файлов)
- [`knowledge_Balanse/`](../knowledge_Balanse/) — расширенный (~10 файлов)
- [`knowledge_PL/`](../knowledge_PL/) — справочный

## Связанные базы знаний

- [[knowledge_Balanse]] — управленческий баланс (Оплата труда уходит в свод)
- [[knowledge_PL]] — статьи PL (ФОТ/ЕСВ)
- [[knowledge_Balanse_money]] — поток денег ЗП в свод
