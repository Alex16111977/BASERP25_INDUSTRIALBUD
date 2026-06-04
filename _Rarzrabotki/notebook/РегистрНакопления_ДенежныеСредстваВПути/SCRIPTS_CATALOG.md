# SCRIPTS_CATALOG — pilot-скрипты доработки РН.ДенежныеСредстваВПути

> Каталог Python COM-скриптов в `Python/test/`. Запускать из этой папки.
> Все скрипты — самодостаточны (connect_erp inline).

## Универсальные diag-скрипты

| Скрипт | Назначение |
|---|---|
| `pilot_rko_check.py` | Быстрый дамп реквизитов + движений 1 документа (РКО) — образец |
| `pilot_spis_post_check.py` | Дамп Спис + Пост + safety-check РСКПС/РСППС движений |
| `pilot_rkr_check.py` | Список всех 5 РасчетКурсовыхРазниц с движениями в РНДС.ВПути |

## Пайплайн пилота (4-shage шаблон)

Для каждого документа 4 скрипта `pilot_<doc>_NN_*.py`:

| # | Назначение | Когда запускать |
|---:|---|---|
| 01 | `baseline.py` — дамп шапки + всех регистров → `_artifacts/pilot_<doc>_01_baseline.json` | ДО правки BSL |
| 02 | `sql_pretest.py` — Python COM тест SQL с новой колонкой Подразделение (Rule #-1) | До правки BSL |
| 03 | `repost.py` — отмена + проведение через COM (со safety-check РСКПС) | После Edit BSL + /db-load-xml + /db-update |
| 04 | `verify.py` — diff baseline vs текущее, проверка Подразделения и Σ | После 03 |

## По документам

### ПКО 000Ц-000001 (✅ закрыт)
| Скрипт | Статус |
|---|---|
| `pilot_01_baseline.py` | ✅ |
| `pilot_02_sql_pretest.py` | ✅ |
| `pilot_03_repost.py` | ✅ |
| `pilot_04_verify.py` | ✅ |

### РКО N0000053020 (✅ закрыт)
| Скрипт | Статус |
|---|---|
| `pilot_rko_01_baseline.py` | ✅ |
| `pilot_rko_02_sql_pretest.py` | ✅ |
| `pilot_rko_03_repost.py` | ✅ |
| `pilot_rko_04_verify.py` | ✅ |

### СписаниеБезнал 00000019546 (✅ закрыт)
| Скрипт | Статус |
|---|---|
| `pilot_spis_01_baseline.py` | ✅ |
| `pilot_spis_02_sql_pretest.py` | ✅ |
| `pilot_spis_03_repost.py` (с safety-check РСКПС) | ✅ |
| `pilot_spis_04_verify.py` | ✅ |

### ПостБезнал 00DL-007179 (✅ закрыт)
| Скрипт | Статус |
|---|---|
| `pilot_post_01_baseline.py` | ✅ |
| `pilot_post_02_sql_pretest.py` | ✅ |
| `pilot_post_03_repost.py` (с safety-check) | ✅ |
| `pilot_post_04_verify.py` | ✅ |

### РасчетКурсовыхРазниц 000Ц-000007 (🔴 в работе)
| Скрипт | Статус |
|---|---|
| `pilot_rkr_check.py` | ✅ инвентаризация 5 документов |
| `pilot_rkr_01_baseline.py` | ⚠ требует доработки — искать по ХозОп, не по номеру (дубль номеров) |
| `pilot_rkr_02_diagnostic.py` | ✅ диагностика валютных остатков |
| `pilot_rkr_03_repost.py` | 🔴 создать по шаблону pilot_post_03 |
| `pilot_rkr_04_verify.py` | 🔴 создать по шаблону pilot_post_04 |

## Шаблон для будущих документов (15 неиспользуемых типов)

Если в БД появятся новые документы из 12 неиспользуемых (ЧекККМ, ВыемкаИзКассыККМ, ЭквайринговыеОперации и т.п.) — копировать pilot_post_*.py (как самый сложный) и менять:
1. Имя `Документ.<ТипДок>` в запросах
2. Номер документа
3. Структуру блоков SELECT (см. CODE_USAGE.md)

## Артефакты (созданы при запуске)

В `_artifacts/`:
- `pilot_*_baseline.json` — snapshot перед правкой
- (pilot_04 не создаёт артефактов — только print + exit-code)

## Правила для новых скриптов

1. **Использовать `getattr(rec, имя, None)`** для итерации полей записи регистра (НЕ `rec[имя]` — TypeError)
2. **Даты СЕРВЕРНО** через `ДАТАВРЕМЯ(y,m,d,h,m,s)` в тексте запроса
3. **Не использовать `--` комментарии** в SQL (только `//` или без комментариев)
4. **Safety-check РСКПС/РСППС** в pilot_03 для документов которые могут писать в расчёты с партнёрами
5. **JSON-snapshot** для baseline → use в verify для diff

## Связанные файлы

- [`README.md`](README.md) — статус 4/5 пилотов PASS
- [`PROGRESS.md`](PROGRESS.md) — хронология
- [`LESSONS.md`](LESSONS.md) — уроки 4 пилотов
- [`CODE_USAGE.md`](CODE_USAGE.md) — карта 17 WRITE-точек
- [`prompts/*.md`](prompts/) — 5 промтов пилотов
