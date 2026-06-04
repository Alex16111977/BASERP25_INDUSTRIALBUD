# Знак суми за ТипСтатьи у А_ФинРез_PL → А_ОтчетPL_Свод → Fact_PnL → PL.pbix (з 2026-05-21)

Дата генерації: 2026-05-21  
Тип: ПОСТІЙНИЙ (бізнес-правило стабільне)  
Пов'язаний файл: pl_articles_catalog.md, pl_report_architecture_analyst.md, knowledge_Olap/olap_powerbi_pl_pbix.md

## Contextual metadata для AI-retrieval

**Company:** ТОВ "ІНДАСТРІАЛБУД" (INDUSTRIALBUD LLC)  
**Document type:** Семантика знака суми у фактовій таблиці PnL — mirror-логіка за реквізитом ТипСтатьи (Доход +, Расход −) для P&L одною формулою без CASE/SIGN  
**Scope:** Проведення `Документ.А_ФинРез_PL` → `РегистрСведений.А_ОтчетPL_Свод` → ETL → `Fact_PnL`.[Sum_*] колонки → DAX-міра Sum_Fact у `PL.pbix`  
**Purpose:** Фінансист бачить чистий P&L (`Σ Sum_Fact`) як готову формулу без додаткових перетворень.  
**Pattern:** mirror-знак — Расход множиться на −1 (зберігає сторно: −50 + (−1) = +50, корекційне зменшення витрат стає прибутком)  
**Related to:** А_ФинРез_DDS (той самий патерн через `А_ТипПоказателя` ДДС), А_ФинРез_Баланс (через ПАП.ВидДвижения)

---

## TL;DR

З 2026-05-21 у документі `А_ФинРез_PL.ОбработкаПроведения` фінальний SELECT у `втРезультат` обертає **6 ресурсів** (СуммаФ1_Excel, СуммаФ2_Excel, Сумма_Excel, СуммаФ1, СуммаФ2, Сумма) у CASE по `Статья.ТипСтатьи`:

```sql
ВЫБОР
    КОГДА Статья.ТипСтатьи = ЗНАЧЕНИЕ(Перечисление.А_ТипСтатьиPL.Расход)
        ТОГДА -Сумма
    ИНАЧЕ Сумма
КОНЕЦ
```

| ТипСтатьи статьи PL | Знак у регістрі |
|---|---|
| Доход (3 статті) | + (як є) |
| Расход (65 статей) | − (mirror) |
| Прочее (ОперационныйИтог, Информационный, ПустаяСсылка) | + (як є) |

## Acceptance результат (2026-05-21)

| Категорія | Кількість рядків | Σ signed | Σ |abs| |
|---|---|---|---|
| Доход (3 статті) | 3 515 | +361 767 426,83 ₽ | 361 767 426,83 ₽ |
| Расход (65 статей) | 51 631 | **−367 290 335,66 ₽** | 367 298 266,84 ₽ |
| **P&L** | 55 146 | **−5 522 908,83 ₽** (УБЫТОК) | |

**1С регістр `А_ОтчетPL_Свод`** і **OLAP `Fact_PnL`** дають **ідентичні** числа до копійки (acceptance verified через `_verify_fact_pnl_signs.py` 2026-05-21).

## Передумова: Σ-інваріант (mirror, не |abs|)

Для **Расход** правка зберігає:
- `Σ_signed_after = −1 × Σ_signed_before` (точне дзеркало)
- `Σ_|abs|_after  = Σ_|abs|_before` (модуль не змінюється)

Це **mirror**, не `|abs|`. Чому це важливо:

| Сценарій | mirror (наш) | |abs| (відкинуто) |
|---|---|---|
| Расход +100 (звичайний) | → −100 | → −100 |
| Расход −50 (сторно/корекція) | → **+50** | → −50 |
| Σ за статтю після | -50 (грошей менше витрачено) | -150 (подвійний обсяг витрат) |

Mirror математично коректний для P&L: сторно витрат збільшує прибуток (як економія).

## Каскад змін (28 проведених А_ФинРез_PL за 2024-01..2026-04 + ETL + OLAP)

```
1. Documents/А_ФинРез_PL/Ext/ObjectModule.bsl
   └→ Функция СформироватьЗапросСверткиPL() → фінальний SELECT втРезультат
      CASE по ТипСтатьи на 6 ресурсах (рядки 587-619)

2. Заливка: /db-load-xml partial + /db-update -Dynamic+

3. Массовий перепровод (28 документів):
   _Rarzrabotki/Python/scripts/_reprovesti_finrez_pl.py
   Результат: OK=28, FAIL=0
   Лог: data/json/20_reprovesti_finrez_pl_log.json

4. Acceptance:
   _Rarzrabotki/Python/test/test_finrez_pl_sign_pretest.py    (baseline)
   _Rarzrabotki/Python/test/test_finrez_pl_sign_verify.py     (mirror-інваріант)
   _Rarzrabotki/Python/test/_verify_fact_pnl_signs.py         (1С == OLAP)

5. OLAP ETL extension:
   - SQL ALTER TABLE Dim_PL_Articles ADD Type_Statya nvarchar(50) NULL
   - mapping/refresh_mapping.py: WHITELIST += "Перечисление.А_ТипСтатьиPL"
   - ai_olap/transformers/enum_resolver.py: FROZEN_ENUMS += ТипСтатьиPL (4 значення)
   - pipelines/dim_catalogs.json step dim_pl_articles: + fields["ТипСтатьи"]
                                                       + enum_resolver column_to_enum
                                                       + column_map ТипСтатьи→Type_Statya
   - python main.py --run-once dim_catalogs → 68 statей загружено з Type_Statya
   - python main.py --run-once fact_pnl     → 55146 рядків з підписаними знаками

6. PowerBI PL.pbix (live підключення):
   - partition_operations Update — M-вираз партиції Dim_PL_Articles
     перейменовано на native query (Sql.Database(..., [Query="SELECT ... Type_Statya FROM ..."]))
     щоб обійти PBI schema cache (memory pbix_sql_database_schema_cache.md)
   - При наступному PBI Refresh колонка Type_Statya → ТипСтатьи з'явиться у модельній таблиці А_Статьи_PL
   - DAX-міра Sum_Fact = SUM(Fact_PnL[Sum_Fact]) тепер дає natural P&L без CASE/SIGN
```

## Що сламається, якщо хтось додасть новий ТипСтатьи

Поточно FROZEN_ENUMS містить 4 значення: Доход (0), Расход (1), ОперационныйИтог (2), Информационный (3).

Якщо у Configurator додають 5-й тип (наприклад «Прочее»), ETL `dim_pl_articles` впаде з помилкою `TransformError: enum order 4 out of frozen list (size 4)`. Це навмисний guard — щоб бухгалтер не зломав схему OLAP без узгодження.

**Дія при додаванні нового типу**:
1. Додати в `ai_olap/transformers/enum_resolver.py` → `FROZEN_ENUMS["Перечисление.А_ТипСтатьиPL"]` нове значення на правильну позицію (`_EnumOrder`).
2. Якщо новий тип потребує специфічного знака в `А_ФинРез_PL` — додати гілку CASE в `ObjectModule.bsl`.
3. Перепровести 28 документів і прогнати ETL.

## Як DAX тепер використовує знак

**До правки** (DAX-міра у Table_Measures):
```dax
Маржинальный доход = 
    CALCULATE(SUM(Fact_PnL[Sum_Fact]), Dim_PL_Articles[Type_Statya] = "Доход") -
    CALCULATE(SUM(Fact_PnL[Sum_Fact]), Dim_PL_Articles[Type_Statya] = "Расход")
```

**Після правки** (одна формула):
```dax
Маржинальный доход = SUM(Fact_PnL[Sum_Fact])
```

Це спрощує всі похідні міри (Прибыль за период, Рентабельность, EBITDA і т.д.).

## Що **НЕ** змінилось

- Структура регістра `А_ОтчетPL_Свод` — 12 вимірів + 6 ресурсів, без змін.
- Звіт `Отчет.А_ОтчетPL` (UI 1С) — читає регістри ЕРП напряму, не Свод, тому не зачеплено.
- Документ `Документ.А_ОтчетPL` (план з Excel) — приходить з PnL pipeline незмінно, перетворення знака відбувається ТІЛЬКИ при проведенні А_ФинРез_PL.
- `|Σ abs|` для кожної (Орг, Статья) ідентичний pretest до 0,01 ₽ (mirror-семантика).

## Anti-patterns (що НЕ робити)

| Анти-патерн | Чому погано | Правильно |
|---|---|---|
| Робити `|abs|` замість mirror | Сторно витрат не дзеркалиться, P&L спотворений | mirror: `-Сумма` для Расход, без модуля |
| Міняти знак у проміжних CTE (втEРП_Расх / втPL / втРезультат секцій UNION) | Може зламати логіку CoGS/БезPL/Excel | Менять тільки у фінальному СУММА в випадку final SELECT |
| Додавати CASE для пустої ТипСтатьи (PL_ЕРП БезPL) | Спотворює рідкісні data-gap рядки | ИНАЧЕ = як є |
| Запустити ETL без оновлення `FROZEN_ENUMS` | TransformError на enum_order out of list | Спочатку enum_resolver.py + whitelist mapping, потім main.py |
| Робити ALTER TABLE без оновлення M-вираз партиції | PBI schema cache не побачить нову колонку | Перейти на native query Sql.Database(..., [Query="SELECT ..."]) |
