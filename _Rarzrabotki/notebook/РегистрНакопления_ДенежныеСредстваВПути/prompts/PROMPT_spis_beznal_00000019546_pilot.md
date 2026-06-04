# Промт — ПИЛОТ доработки СписаниеБезнал 00000019546 (Подразделение в РНДС.ВПути)

## TL;DR

Пилот №3 (после ПКО ✅ и РКО ✅). Доработать `ТекстЗапросаТаблицаДенежныеСредстваВПути` в `Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl` (стр. **2859**) — добавить `Подразделение = ДанныеДокумента.Подразделение` в **3 блока ОБЪЕДИНИТЬ ВСЕ**.

**Тестовый документ:** `Списание безналичных ДС 00000019546 от 10.12.2025 17:17:20`, БС = «ОТП_ТОВ ІНДАСТРІАЛБУД UA9730», Подразделение шапки = **Строительство**, Сумма 2 636 018,46 UAH, ХозОп **«Перечисление ДС на другой счет»** (межбанк, БЕЗ РСКПС/РСППС → COM-репост безопасен).

**Pilot-скрипты:** `pilot_spis_01_baseline.py` → `pilot_spis_02_sql_pretest.py` → `pilot_spis_03_repost.py` → `pilot_spis_04_verify.py`.

## Параметры

| Параметр | Значение |
|---|---|
| Файл | `Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl` |
| Функция | `ТекстЗапросаТаблицаДенежныеСредстваВПути` (стр. **2859**) |
| Тестовый док | `00000019546` от 10.12.2025 17:17:20, 2 636 018,46 UAH |
| Подразделение шапки | **Строительство** |
| ХозОп пилота | Перечисление ДС на другой счет (блок 2 — стр. **2897-2922**) |
| Безопасность COM | ✅ безопасно (для этой ХозОп нет РСКПС/РСППС; проверено в `pilot_spis_03_repost.py` через safety-check) |

⚠ **Для других СписаниеБезнал** с ХозОп `Оплата поставщику` / `Возврат ДС от поставщика` — COM ОПАСЕН (`feedback_com_repost_skips_registrator_raschetov`). Эти документы требуют UI-репоста. `pilot_spis_03_repost.py` имеет встроенный safety-check (выходит с ошибкой если есть РСКПС/РСППС).

## Что уже сделано

1. ✅ ПКО пилот PASS (Documents/ПриходныйКассовыйОрдер/Ext/ManagerModule.bsl, 5 блоков)
2. ✅ РКО пилот PASS (Documents/РасходныйКассовыйОрдер/Ext/ManagerModule.bsl, 3 блока)
3. ✅ pilot_spis_01 baseline протестирован: РНДС.ВПути 1 строка Приход 2 636 018,46 с Подр=(пусто)
4. ✅ pilot_spis_02 sql пре-тест протестирован: SQL с `ДанныеДокумента.Подразделение` → «Строительство»

## Структура функции (3 блока)

| # | Строки | ХозОперация | Получатель | Отправитель |
|---|---|---|---|---|
| 1 | 2867-2893 | СнятиеНаличныхДенежныхСредств | `КассаПолучатель` | `БанковскийСчет` |
| 2 | 2897-2922 | **ПеречислениеДенежныхСредствНаДругойСчет** ⭐ пилот | `БанковскийСчетПолучатель` | `БанковскийСчет` |
| 3 | 2926-2954 | КонвертацияВалюты | `БанковскийСчетПолучатель` | НЕОПРЕДЕЛЕНО |

### ⚠ ОТЛИЧИЕ от ПКО/РКО

В этой функции SELECT **НЕТ строки** `&ХозяйственнаяОперация КАК ХозяйственнаяОперация` (она вычисляется снаружи через `&ХозяйственнаяОперация` в WHERE). Только `&СтатьяДвиженияДенежныхСредств КАК ...`.

**Вставлять Подразделение перед `&СтатьяДвиженияДенежныхСредств`:**

```sql
|	ДанныеДокумента.СуммаДокумента * &КоэффициентПересчетаВВалютуРегл            КАК СуммаРегл,
|
|	ДанныеДокумента.Подразделение                                                КАК Подразделение,
|	&СтатьяДвиженияДенежныхСредств                                               КАК СтатьяДвиженияДенежныхСредств
```

## Алгоритм (8 шагов)

```bash
# 1. Бэкап
cp "C:/Configuration_downloads/BASERP25/Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl" \
   "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/РегистрНакопления_ДенежныеСредстваВПути/_backup/SpisBeznal_ManagerModule_BEFORE.bsl"

# 2-3. Baseline + SQL пре-тест (уже выполнены, можно перезапустить)
cd "C:/Configuration_downloads/BASERP25/_Rarzrabotki/notebook/РегистрНакопления_ДенежныеСредстваВПути/Python/test"
python pilot_spis_01_baseline.py
python pilot_spis_02_sql_pretest.py

# 4. Read функции (Read offset=2859 limit=100)

# 5. Edit 3 блоков (replace_all:false с уникальным контекстом по WHERE ХозОп)
#    Для каждого блока добавить:
#    |	ДанныеДокумента.Подразделение                                                КАК Подразделение,

# 6-7. /db-load-xml + /db-update (закрыть Конфигуратор!)
powershell.exe -NoProfile -File "C:\Configuration_downloads\BASERP25\.claude\skills\db-load-xml\scripts\db-load-xml.ps1" `
  -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" `
  -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" -UserName "Администратор" -Password "24043" `
  -ConfigDir "C:\Configuration_downloads\BASERP25" `
  -Mode Partial -Files "Documents/СписаниеБезналичныхДенежныхСредств/Ext/ManagerModule.bsl"

powershell.exe -NoProfile -File "C:\Configuration_downloads\BASERP25\.claude\skills\db-update\scripts\db-update.ps1" `
  -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" `
  -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" -UserName "Администратор" -Password "24043" `
  -Dynamic "+"

# 8. Repost + Verify
python pilot_spis_03_repost.py
python pilot_spis_04_verify.py
```

## Acceptance

PASS если:
1. ✅ 3 блока SELECT содержат `Подразделение`
2. ✅ `/db-update` прошёл
3. ✅ `pilot_spis_03_repost.py` показывает `[OK] Проведено` (или safety-stop если появились РСКПС движения)
4. ✅ `pilot_spis_04_verify.py` показывает PASS — Подр в РНДС.ВПути = «Строительство»
5. ✅ Σ Сумма (2 636 018,46) НЕ изменилась
6. ✅ Прочие регистры без изменений

## Антипаттерны

- ❌ Edit `replace_all: true` — задеть другие блоки
- ❌ Использовать `Касса.Подразделение` или `БанковскийСчет.Подразделение` — нужно ИЗ ШАПКИ (как РНДС.Наличные)
- ❌ Перепроводить через COM документы с РСКПС/РСППС (для этого ПКО — безопасно, для других — НЕТ)
- ❌ Запускать pilot_spis_03 ДО /db-update

## Контакты

- Memory: `balans_money_knowledge_base`, `feedback_com_repost_skips_registrator_raschetov`, `feedback_use_db_skills_for_config_load`, `feedback_balans_etalon_period_serverside`.
- Пилот ПКО `PROMPT_pko_000001_pilot.md` (закрыт PASS).
- Пилот РКО `PROMPT_rko_N0000053020_pilot.md` (закрыт PASS).
- Baseline уже создан: `_artifacts/pilot_spis_01_baseline.json`.
- SQL пре-тест пройден: SQL с `ДанныеДокумента.Подразделение` → «Строительство».
