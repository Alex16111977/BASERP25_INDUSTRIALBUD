# _archive_pre_split/ — Backup монолітних PL-dump версій

**Дата архівації:** 2026-04-23

## Призначення

Це backup **монолітних** версій помісячних PL-виписок, **які були у NotebookLM до 2026-04-23**. У межах upgrade-плану (`skill-superpowers-brainstorming-atomic-clock.md`) ці файли замінено на **розбиті за групами PL-статей** (6 файлів/місяць) для кращого RAG retrieval.

## Вміст

- `pl_dump_2025_12_december.md` — **1 040 КБ**, повна виписка за Грудень 2025.
- `pl_dump_2026_01_january.md` — **862 КБ**, повна виписка за Січень 2026.
- `pl_dump_2026_02_february.md` — **923 КБ**, повна виписка за Лютий 2026.

## Чому розбили

Досліджено (arxiv 2402.05131, Snowflake Eng Blog):
- Оптимальний chunk-size для RAG: **~1,800 chars**.
- 14,400+ chars → −10-20% accuracy.
- Наші монолітні файли = ~900K chars → NotebookLM повертав тільки headers без даних таблиць.

## Чому не видаляємо

- Архівна копія на випадок коли потрібна суцільна виписка по місяцю (manual review, audit).
- Довідник ДЕ що було до розбиття (для історичного трейсу).

## Що замість

Див. у `knowledge_PL/`:
- `pl_dump_YYYY_MM_01_summary.md` — TL;DR + топ-N
- `pl_dump_YYYY_MM_02_income_revenue.md` — група Операционный доход
- `pl_dump_YYYY_MM_03_cost_of_goods.md` — група Себестоимость
- `pl_dump_YYYY_MM_04_opex_admin.md` — групи ОПЗ + Адмін
- `pl_dump_YYYY_MM_05_marketing_fin.md` — групи Маркетинг + Фінансова + Податки
- `pl_dump_YYYY_MM_06_cash_anomalies.md` — Каса + аномалії

## ❌ НЕ завантажувати в NotebookLM

Ці файли у папці `_archive_pre_split/` **не повинні потрапити** в блокнот INDUSTRIALBUD_PL. Актуальні — у корені `knowledge_PL/`.
