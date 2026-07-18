# ПРОМТ: А_ФинРез_PL — знак суммы по ТипСтатьи (Доход +, Расход −) + сквозная передача в OLAP+PowerBI

Ты работаешь в `C:\Configuration_downloads\BASERP25` (BAS ERP 2.5 INDUSTRIALBUD, 1C:Enterprise 8.3.20+).
Базы: BaseERP (Srvr=SQLSERVER, Usr=Администратор, Pwd=24043), OlapBASERP (Srvr=localhost, Usr=sa, Pwd=Brw739182465!).
Python: `C:\Python313\python.exe` (системный) или `_Rarzrabotki\Python\venv\Scripts\python.exe` (для OLAP — там зависимости).

## Цель (как просит финансист)

В справочнике `Catalog.А_Статьи_PL` есть реквизит `ТипСтатьи` типа `ПеречислениеСсылка.А_ТипСтатьиPL` со значениями `Доход` / `Расход`.

Нужно при проведении документа `Документ.А_ФинРез_PL` (свертка PnL для OLAP) делать движения в `РегистрСведений.А_ОтчетPL_Свод` со **знаком, который соответствует семантике статьи**:

- `ТипСтатьи = Доход` → **все 6 ресурсов** регистра идут со знаком **+** (положительные).
- `ТипСтатьи = Расход` → **все 6 ресурсов** идут со знаком **−** (отрицательные).

Семантика: после этого `СУММА(Сумма)` без модуля даёт сразу **маржинальный доход / убыток / net P&L** одной формулой, без CASE/SIGN на стороне DAX-измерений или SQL.

После правки 1С:
- расширить ETL → Dim_PL_Articles + Type_Statya (новая колонка в OlapBASERP);
- обновить `PL.pbix` — добавить колонку Type_Statya в модель, обеспечить корректность Sum_Fact с учётом нового знака;
- перепровести все 28 проведённых документов `А_ФинРез_PL` (2024-01..2026-04) чтобы регистр содержал новые знаки;
- запустить acceptance-тест на финансовом invariant'е (|Σ| абсолютных сумм не меняется — меняется только знак).

## Образец-эталон: точно тот же паттерн уже работает в А_ФинРез_DDS

`Documents/А_ФинРез_DDS/Ext/ObjectModule.bsl:161-164` (и аналогично для наличных, казны, плана):

```bsl
ВЫБОР КОГДА Бн.СтатьяДвиженияДенежныхСредств.А_ТипПоказателя = &ТипПрих
    ТОГДА ВЫБОР КОГДА Бн.Сумма < 0 ТОГДА -Бн.Сумма ИНАЧЕ Бн.Сумма КОНЕЦ
    ИНАЧЕ -(ВЫБОР КОГДА Бн.Сумма < 0 ТОГДА -Бн.Сумма ИНАЧЕ Бн.Сумма КОНЕЦ)
КОНЕЦ КАК Сумма
```

Декомпозиция: `|Сумма|` (абсолют от исходной) → знак по типу показателя. То же самое делается в `А_ФинРез_Баланс` (см. `knowledge_Balanse/balanse_provedenie_logic.md` — там это часть «свертки активов/пассивов»). Для PL пользуемся реквизитом `СтатьяPL.ТипСтатьи` вместо `А_ТипПоказателя` ДДС.

**ВАЖНО про абсолют:** в исходных данных Excel (`Документ.А_ОтчетPL.ДанныеОтчета`) суммы плана уже **положительные** для расходных статей (финансист записывает фактическое значение, без знака). В ЕРП-регистрах факт расходов тоже положительный (`СуммаПриход` всегда ≥ 0). Поэтому в нашем CASE для PL можно **не делать `|Сумма|`** — достаточно `ВЫБОР ТипСтатьи = Расход ТОГДА -Сумма ИНАЧЕ Сумма`. Но **для устойчивости** возьми `|Сумма|` через CASE — если кто-то в будущем добавит отрицательные суммы в источник, поведение останется корректным.

---

## Метаданные подтверждены через MCP get_metadata_structure / Glob (2026-05-21)

### Справочник.А_Статьи_PL
- `ТипСтатьи` — `ПеречислениеСсылка.А_ТипСтатьиPL (Доход, Расход)` — есть в конфигурации.
- В текущей базе: 65 статей с типом Расход + 3 с типом Доход (см. `pl_articles_catalog.md`).

### Документ.А_ФинРез_PL.ObjectModule.bsl
- Функция `СформироватьЗапросСверткиPL()` (line 113-620) собирает 4 СТЕ + 6 секций UNION ALL.
- Финальный SELECT (line 585-619) делает `СУММА()` по 6 ресурсам: `СуммаФ1_Excel`, `СуммаФ2_Excel`, `Сумма_Excel`, `СуммаФ1`, `СуммаФ2`, `Сумма`.
- Все 6 ресурсов сейчас **только положительные**.

### РегистрСведений.А_ОтчетPL_Свод
- 12 измерений + 6 ресурсов (те же что в СУММА).
- Регистр Независимый, без периодичности (одна запись на ключ {Орг+Подр+Источник+Статья+ДДС+Контрагент+ФизЛицо+ДокДвижения+СтатьяДоходов+СтатьяРасходов+Аналитика+Комментарий}).
- В БД сейчас 28 проведённых документов А_ФинРез_PL за 2024-01..2026-04 (~50K строк регистра).

### OLAP — OlapBASERP
- Таблица `Dim_PL_Articles` имеет колонки: `PL_Article_ID, PL_Article_Code, PL_Article_Name, Parent_ID, Is_Group, Marked_For_Deletion, Group_ID, Sort_Order, Loaded_At` (9 шт).
- Поле `Type_Statya` **ОТСУТСТВУЕТ** — его надо добавить.
- Pipeline `Ai_Olap/pipelines/dim_catalogs.json` шаг `dim_pl_articles` (line 667-708) извлекает 7 fields, `ТипСтатьи` НЕ извлекается.
- Whitelist в `mapping/refresh_mapping.py` нужно расширить, иначе `sql_backend_extractor` тихо скипнет поле (см. memory `olap_refresh_mapping_strict_whitelist.md`).
- FROZEN_ENUMS в том же файле должен содержать `Перечисление.А_ТипСтатьиPL` чтобы enum_resolver преобразовал UUID в текст «Доход» / «Расход».

### PowerBI
- File: `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\PowerBi\PL.pbix`
- Подключение: Sql.Database("localhost", "OlapBASERP"), Import mode, Auth=sa.
- Power Query партиция `Dim_PL_Articles` → модельное имя `А_Статьи_PL` (после переименования 2026-05-03).
- Новая колонка `Type_Statya` должна стать видимой в модельной таблице с локализованным именем `ТипСтатьи`.

---

## Stage 1 — Python COM pretest: снимок baseline регистра

**Цель:** до правки BSL зафиксировать Σ |Сумма| по (Организация, Статья, ТипСтатьи) для всех 28 месяцев. После правки эта сумма абсолютных значений должна остаться **неизменной** — поменяется только знак.

**Файл:** `_Rarzrabotki/Python/test/test_finrez_pl_sign_pretest.py`

```python
# -*- coding: utf-8 -*-
"""Pretest baseline для А_ФинРез_PL → А_ОтчетPL_Свод.
Снимает Σ |Сумма| по (Орг, Статья, ТипСтатьи) до правки знака."""
import sys, io, json, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    Р.Организация КАК Орг,
    Р.Статья КАК Статья,
    Р.Статья.ТипСтатьи КАК ТипСтатьи,
    КОЛИЧЕСТВО(*) КАК Строк,
    СУММА(Р.СуммаФ1_Excel) КАК Σ_Ф1_Excel,
    СУММА(Р.СуммаФ2_Excel) КАК Σ_Ф2_Excel,
    СУММА(Р.Сумма_Excel) КАК Σ_Excel,
    СУММА(Р.СуммаФ1) КАК Σ_Ф1,
    СУММА(Р.СуммаФ2) КАК Σ_Ф2,
    СУММА(Р.Сумма) КАК Σ_Сумма,
    СУММА(ВЫБОР КОГДА Р.СуммаФ1_Excel < 0 ТОГДА -Р.СуммаФ1_Excel ИНАЧЕ Р.СуммаФ1_Excel КОНЕЦ) КАК Σabs_Ф1_Excel,
    СУММА(ВЫБОР КОГДА Р.Сумма < 0 ТОГДА -Р.Сумма ИНАЧЕ Р.Сумма КОНЕЦ) КАК Σabs_Сумма
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р
СГРУППИРОВАТЬ ПО Р.Организация, Р.Статья, Р.Статья.ТипСтатьи
"""
tz = q.Выполнить().Выгрузить()

print(f"Всего групп (Орг,Статья,ТипСтатьи): {tz.Количество()}")

# Сводка по типам
дох = [r for r in tz if conn.String(r.ТипСтатьи) == "Доход"]
расх = [r for r in tz if conn.String(r.ТипСтатьи) == "Расход"]
пуст = [r for r in tz if r.ТипСтатьи is None or not conn.ЗначениеЗаполнено(r.ТипСтатьи)]

def Σ(rows, fld):
    return sum(float(getattr(r, fld)) for r in rows)

print(f"\n=== Доход ({len(дох)} групп) ===")
print(f"  Σ Сумма (со знаком как сейчас): {Σ(дох,'Σ_Сумма'):>20,.2f}")
print(f"  Σ |Сумма| (абсолют):            {Σ(дох,'Σabs_Сумма'):>20,.2f}")

print(f"\n=== Расход ({len(расх)} групп) ===")
print(f"  Σ Сумма (со знаком как сейчас): {Σ(расх,'Σ_Сумма'):>20,.2f}")
print(f"  Σ |Сумма| (абсолют):            {Σ(расх,'Σabs_Сумма'):>20,.2f}")

print(f"\n=== Без ТипСтатьи / Пустая ({len(пуст)} групп) ===")
print(f"  Σ Сумма:                         {Σ(пуст,'Σ_Сумма'):>20,.2f}")

# Сохранить baseline
out = []
for r in tz:
    out.append({
        "org": conn.String(r.Орг),
        "статья": conn.String(r.Статья),
        "тип": conn.String(r.ТипСтатьи) if conn.ЗначениеЗаполнено(r.ТипСтатьи) else "",
        "строк": int(r.Строк),
        "abs_sums": {
            "Ф1_Excel": float(r.Σabs_Ф1_Excel),
            "Ф2_Excel": float(r.Σ_Ф2_Excel),  # для Excel абсолюты считать опционально
            "Сумма": float(r.Σabs_Сумма),
        },
        "signed_sums": {
            "Сумма_now": float(r.Σ_Сумма),
        }
    })

path = os.path.join(os.path.dirname(__file__), "finrez_pl_sign_pretest.json")
with open(path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print(f"\nSaved: {path}")
```

**Запуск:**
```powershell
C:\Python313\python.exe _Rarzrabotki\Python\test\test_finrez_pl_sign_pretest.py
```

**Acceptance Stage 1:**
- Создан baseline JSON `finrez_pl_sign_pretest.json` с разрезом (Орг, Статья, ТипСтатьи).
- Видно соотношение Доход/Расход/Пустых групп (статьи без ТипСтатьи — это тревожный сигнал, отдельный alert).
- Для строк PL без статьи (Источник=PL_ЕРП БезPL) — `ТипСтатьи` пустой, для них знак не меняется.

---

## Stage 2 — Правка BSL: добавить знак-CASE в финальный SELECT

**Файл:** `Documents/А_ФинРез_PL/Ext/ObjectModule.bsl`

**Точка изменения:** функция `СформироватьЗапросСверткиPL()`, финальный SELECT в `втРезультат` (текущие строки 585-619).

**Стратегия (минимально-инвазивная):**

1. **НЕ менять** 4 промежуточные СТЕ (`втPL`, `втPL_ДДС`, `втЕРП_Расх`, `втЕРП_Дох`, `втЕРП_CoGS`) и 6 секций UNION ALL формирования `втРезультат` — все эти суммы остаются положительными, как и раньше.

2. **В финальном SELECT** (line 585-619) обернуть каждую `СУММА(...)` для всех 6 ресурсов в CASE по `Статья.ТипСтатьи`:

```bsl
|	СУММА(ВЫБОР
|		КОГДА втРезультат.Статья.ТипСтатьи = ЗНАЧЕНИЕ(Перечисление.А_ТипСтатьиPL.Расход)
|			ТОГДА -(ВЫБОР КОГДА втРезультат.СуммаФ1_Excel < 0
|				ТОГДА -втРезультат.СуммаФ1_Excel ИНАЧЕ втРезультат.СуммаФ1_Excel КОНЕЦ)
|		КОГДА втРезультат.Статья.ТипСтатьи = ЗНАЧЕНИЕ(Перечисление.А_ТипСтатьиPL.Доход)
|			ТОГДА (ВЫБОР КОГДА втРезультат.СуммаФ1_Excel < 0
|				ТОГДА -втРезультат.СуммаФ1_Excel ИНАЧЕ втРезультат.СуммаФ1_Excel КОНЕЦ)
|		ИНАЧЕ втРезультат.СуммаФ1_Excel
|	КОНЕЦ) КАК СуммаФ1_Excel,
```

Аналогично для `СуммаФ2_Excel`, `Сумма_Excel`, `СуммаФ1`, `СуммаФ2`, `Сумма`. **6 одинаковых CASE-блоков**, отличаются только именем поля.

3. **Обработка пустой ТипСтатьи** — `ИНАЧЕ Сумма` (без переворота). Это покрывает строки «PL_ЕРП БезPL» (Источник = PL_ЕРП, Статья = ПустаяСсылка) — там тип неизвестен, оставляем как есть. Это редкие data-gap строки, отдельно анализируются финансистом.

**Рефакторинг для читаемости (опционально):** выделить CASE в общий блок ниже, заворачивая `втРезультат.Сумма` в `СУММА(ВЫБОР …)`. Если получится 6 одинаковых блоков подряд — можно вынести в подзапрос с pre-computed знаком, но это усложнит читаемость. **Рекомендация:** оставить 6 inline-CASE — это явно, нет magic.

**Контроль:** функция должна остаться синтаксически корректной (закрытие всех скобок, точки с запятой). После правки сделать smoke-тест: `python _Rarzrabotki/Python/test/test_finrez_pl_sign_query_validate.py` — простой скрипт который собирает текст запроса через `Метаданные.Документы.А_ФинРез_PL.МодульМенеджера` + парсит SQL для проверки.

**Скопировать в основную конфигурацию:** правка идёт сразу в `C:\Configuration_downloads\BASERP25\Documents\А_ФинРез_PL\Ext\ObjectModule.bsl` (не worktree).

---

## Stage 3 — Загрузка конфигурации + smoke 1 документа

**Skills для использования:**
- `db-load-xml` — partial load одного файла.
- `db-update` — UpdateDBCfg Dynamic+.

**Команда:**
```powershell
powershell.exe -NoProfile -File .claude/skills/db-load-xml/scripts/db-load-xml.ps1 `
    -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" `
    -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" `
    -UserName "Администратор" -Password "24043" `
    -ConfigDir "C:\Configuration_downloads\BASERP25" `
    -Mode Partial -Files "Documents/А_ФинРез_PL/Ext/ObjectModule.bsl"

powershell.exe -NoProfile -File .claude/skills/db-update/scripts/db-update.ps1 `
    -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" `
    -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" `
    -UserName "Администратор" -Password "24043" -Dynamic "+"
```

**Smoke-тест 1 документа** через MCP:
1. `find_document_ref(doc_type="А_ФинРез_PL", number="00000000005", date="2026-04-30T...")` — найти один документ за апрель 2026 (UUID).
2. `unpost_document(doc_ref=UUID)` → должен пройти.
3. `post_document(doc_ref=UUID)` → должен пройти.
4. `execute_query` — Σ Сумма за апрель 2026 разделить на Доход/Расход:

```sql
ВЫБРАТЬ
    Р.Статья.ТипСтатьи КАК Тип,
    СУММА(Р.Сумма) КАК Σ_Сумма,
    СУММА(ВЫБОР КОГДА Р.Сумма < 0 ТОГДА -Р.Сумма ИНАЧЕ Р.Сумма КОНЕЦ) КАК Σabs
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р
ГДЕ Р.ДокументДвижения = &Док  -- А_ФинРез_PL UUID
СГРУППИРОВАТЬ ПО Р.Статья.ТипСтатьи
```

**Acceptance Stage 3:**
- Доход: Σ_Сумма == Σabs (всё положительное).
- Расход: Σ_Сумма == -Σabs (всё отрицательное).
- Σ_Сумма_Расход + Σ_Сумма_Доход = чистая прибыль/убыток (читаемо одной формулой).
- |Σabs| совпадает с pretest для этого месяца (до копейки).

---

## Stage 4 — Массовый перепровод 28 А_ФинРез_PL

**Файл:** `_Rarzrabotki/Python/scripts/_reprovesti_finrez_pl.py`

```python
# -*- coding: utf-8 -*-
"""Массовый перепровод всех А_ФинРез_PL за 2024-01..2026-04."""
import sys, io, json, os, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ Ссылка, Номер, Дата, Месяц, Организация
ИЗ Документ.А_ФинРез_PL
ГДЕ Проведен И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Месяц, Организация
"""
tz = q.Выполнить().Выгрузить()
print(f"Найдено А_ФинРез_PL к перепроводу: {tz.Количество()}\n")

mode = conn.PredefinedValue("РежимЗаписиДокумента.Проведение")
results = []
ok = fail = 0

for i in range(tz.Количество()):
    row = tz.Получить(i)
    obj = row.Ссылка.ПолучитьОбъект()
    if obj is None:
        results.append({"номер": str(row.Номер), "статус": "FAIL", "ошибка": "obj is None"})
        fail += 1
        print(f"  FAIL {row.Номер} {row.Месяц}: obj is None")
        continue
    try:
        obj.Записать(mode)
        results.append({
            "номер": str(row.Номер), "месяц": str(row.Месяц),
            "орг": conn.String(row.Организация), "статус": "OK"
        })
        ok += 1
        print(f"  OK   {row.Номер} {row.Месяц}")
    except Exception as e:
        err = str(e)
        if hasattr(e, 'excepinfo') and e.excepinfo:
            err = e.excepinfo[2] or err
        results.append({"номер": str(row.Номер), "статус": "FAIL", "ошибка": err})
        fail += 1
        print(f"  FAIL {row.Номер}: {err}")

print(f"\n=== ИТОГ: OK={ok}, FAIL={fail} ===")

log_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "data", "json"))
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "20_reprovesti_finrez_pl_log.json")
with open(log_path, "w", encoding="utf-8") as f:
    json.dump({
        "ran_at": datetime.datetime.now().isoformat(),
        "total": tz.Количество(), "ok": ok, "fail": fail, "items": results,
    }, f, ensure_ascii=False, indent=2)
print(f"Log: {log_path}")
```

**Запуск:**
```powershell
C:\Python313\python.exe _Rarzrabotki\Python\scripts\_reprovesti_finrez_pl.py
```

**Acceptance Stage 4:**
- OK = 28, FAIL = 0.
- В логе `data/json/20_reprovesti_finrez_pl_log.json` все 28 OK.

**Verify test:** `_Rarzrabotki/Python/test/test_finrez_pl_sign_verify.py` — повторяет SQL pretest и сравнивает:
- |Σ Сумма| (абсолют) по каждой группе (Орг, Статья) должен совпасть с pretest до 0.01 ₽.
- Σ Сумма (со знаком) — для Доход = +Σabs, для Расход = −Σabs.
- 0 групп с ТипСтатьи=Пусто И Σabs > 0 (если есть — alert в логе, для PL_ЕРП БезPL это нормально).

---

## Stage 5 — OLAP ETL: добавить Type_Statya в Dim_PL_Articles

**Шаг 5.1 — SQL DDL** (через `mcp__sql-server-mcp__execute_query` или прямой ALTER):

```sql
USE OlapBASERP;
ALTER TABLE Dim_PL_Articles ADD Type_Statya nvarchar(50) NULL;
```

**Шаг 5.2 — расширить pipeline `Ai_Olap/pipelines/dim_catalogs.json` шаг `dim_pl_articles` (line 667-708):**

```json
{
  "step_id": "dim_pl_articles",
  "extractor": {
    "type": "sql_backend",
    "object": "Справочник.А_Статьи_PL",
    "fields": [
      "Ссылка", "Код", "Наименование", "Родитель", "ЭтоГруппа",
      "ПометкаУдаления", "Группа",
      "ТипСтатьи"
    ]
  },
  "transformer": {
    "steps": ["varbinary_to_uuid", "enum_resolver", "column_mapper"],
    "options": {
      "enum_resolver": {
        "column_to_enum": {
          "ТипСтатьи": "Перечисление.А_ТипСтатьиPL"
        }
      },
      "column_mapper": {
        "column_map": {
          "Ссылка": "PL_Article_ID",
          "Код": "PL_Article_Code",
          "Наименование": "PL_Article_Name",
          "Родитель": "Parent_ID",
          "ЭтоГруппа": "Is_Group",
          "ПометкаУдаления": "Marked_For_Deletion",
          "Группа": "Group_ID",
          "ТипСтатьи": "Type_Statya"
        },
        "defaults": {"Is_Group": false, "Marked_For_Deletion": false}
      }
    }
  },
  "loader": {"target_table": "Dim_PL_Articles", "mode": "full_reload"}
}
```

**Шаг 5.3 — расширить whitelist в `Ai_Olap/mapping/refresh_mapping.py`:**

Memory `olap_refresh_mapping_strict_whitelist.md` говорит: новый 1С-объект для ETL/enum добавить в WHITELIST ДО прогона, иначе `sql_backend_extractor` тихо скипнет поле + возможен дрейф `_FldNNN` при `python mapping/refresh_mapping.py`.

```python
# В WHITELIST_FIELDS["Справочник.А_Статьи_PL"] добавить "ТипСтатьи"
# В FROZEN_ENUMS добавить "Перечисление.А_ТипСтатьиPL"
```

**Шаг 5.4 — refresh mapping + run ETL:**

```powershell
cd C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap
C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\.venv\Scripts\python.exe mapping\refresh_mapping.py
C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\Ai_Olap\.venv\Scripts\python.exe main.py --pipeline dim_catalogs --step dim_pl_articles
```

**Acceptance Stage 5 (SQL верификация):**

```sql
SELECT Type_Statya, COUNT(*) AS Cnt
FROM OlapBASERP.dbo.Dim_PL_Articles
WHERE Marked_For_Deletion = 0
GROUP BY Type_Statya
ORDER BY Cnt DESC;
```

Ожидается: 3 группы (Доход / Расход / NULL для групп-папок). Для не-группа статей — NULL только если в 1С статья не заполнила реквизит (тогда финансисту alert на ручную правку).

---

## Stage 6 — PowerBI PL.pbix: добавить колонку Type_Statya в модель

**Файл:** `C:\Configuration_downloads\BASERP25\_Rarzrabotki\Olap\PowerBi\PL.pbix`

**Skills для использования** — `powerbi-modeling-mcp__*`:
- `connection_operations` — открыть .pbix
- `column_operations` — добавить колонку из источника
- `partition_operations` — обновить Power Query M (чтобы новая колонка попала в импорт)
- `model_operations` — сохранить

**Шаг 6.1 — открыть PL.pbix:**
```python
mcp__powerbi-modeling-mcp__connection_operations(action="open", file_path="C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap/PowerBi/PL.pbix")
```

**Шаг 6.2 — обновить M-выражение партиции Dim_PL_Articles:**

Текущее выражение (читает все колонки таблицы Sql.Database):
```m
let
    Source = Sql.Database("localhost", "OlapBASERP"),
    Dim_PL_Articles = Source{[Schema="dbo",Item="Dim_PL_Articles"]}[Data]
in
    Dim_PL_Articles
```

Это автоматически подхватит новую SQL-колонку `Type_Statya`. Но memory `pbix_sql_database_schema_cache.md` предупреждает: новая колонка после ALTER может **не появиться** из-за PBI schema cache. Workaround — переписать на native query:

```m
let
    Source = Sql.Database("localhost", "OlapBASERP", [Query="SELECT PL_Article_ID, PL_Article_Code, PL_Article_Name, Parent_ID, Is_Group, Marked_For_Deletion, Group_ID, Sort_Order, Type_Statya, Loaded_At FROM Dim_PL_Articles"])
in
    Source
```

Применить через `partition_operations(action="update", table="Dim_PL_Articles", expression=...)`.

**Шаг 6.3 — переименовать колонку для UI:** через `column_operations(action="rename", table="А_Статьи_PL", column="Type_Statya", new_name="ТипСтатьи")` (Power BI модельное имя по-русски согласно конвенции v3.8).

**Шаг 6.4 — сохранить + Full Refresh:**
```python
mcp__powerbi-modeling-mcp__model_operations(action="save")
# Затем — пользователь открывает .pbix в Power BI Desktop, Home → Refresh
```

**Acceptance Stage 6:**
- В модели `А_Статьи_PL` появилась колонка `ТипСтатьи` со значениями `Доход` / `Расход`.
- DAX-мера `Sum_Fact = SUM(Fact_PnL[Sum])` теперь даёт **знаковую** сумму (Доход +, Расход −).
- Если финансист добавит фильтр `[ТипСтатьи] = "Доход"` на визуал — фильтрует только доходные статьи (выручка, прочие доходы, фин. доход).
- Маржинальный доход в матрице = `[Sum_Fact]` без `SIGN()`/`IF()` (формула упрощается).

---

## Финальная Acceptance: сквозной P&L

После всех 6 стадий открыть в Power BI визуал «Маржинальный доход по периодам». Проверить:

1. **Декабрь 2025**: Σ доходов − Σ собівартість − Σ opex = реальный P&L (можно сверить с Отчёт.А_ОтчетPL в 1С).
2. **Все 28 месяцев** в Fact_PnL имеют корректные знаки в колонке `Sum`.
3. **% NULL в Type_Statya** в Dim_PL_Articles ≤ кол-во групп-папок (это нормально); для конкретных статей NULL ≠ допустимо — alert финансисту.
4. **Σ |Sum|** в Fact_PnL (через DAX `SUMX(Fact_PnL, ABS([Sum]))`) === Σ Сумма из 1С регистра до правки (по запросу через MCP).

---

## Anti-patterns (что НЕ делать)

| Анти-паттерн | Почему плохо | Правильно |
|---|---|---|
| Менять знак в `втРезультат` секциях UNION (СтейджаPL_Excel, PL_ЕРП, CoGS и т.д.) | Может разбить логику CoGS / БезPL (там разные ИЗ-источники) | Менять только в **финальном SELECT СУММА**, оставив 6 UNION-секций положительными |
| Делать `ИНАЧЕ -Сумма КОНЕЦ` (без CASE на Доход) | Пустые ТипСтатьи (PL_ЕРП БезPL) получат знак − что неверно | Три ветки: Доход → +, Расход → −, ИНАЧЕ → как есть |
| Изменить структуру регистра А_ОтчетPL_Свод | Ломает Fact_PnL ETL и существующие dashboards | Структура регистра остаётся, меняется только знак ресурсов |
| Не обновить whitelist `refresh_mapping.py` ДО прогона ETL | `sql_backend_extractor` тихо скипнет ТипСтатьи без ошибки | Добавить ТипСтатьи в WHITELIST_FIELDS + FROZEN_ENUMS ПЕРЕД main.py |
| Запустить main.py без `--step dim_pl_articles` | Полный ETL займёт 30+ минут на ничего не меняющую регенерацию | Узкий шаг `--pipeline dim_catalogs --step dim_pl_articles` |
| Удалить колонку SuMPL / Σ_Excel | Эти ресурсы используются в существующих DAX-мерах | Все 6 ресурсов остаются, только знак меняется |
| Не делать pretest снимок | Невозможно доказать что |Σabs| совпадает | ОБЯЗАТЕЛЬНО Stage 1 ДО Stage 2 |
| Использовать `--no-verify` для git коммита | Hooks BSL/Python могут поймать syntax errors | Запускать без флага, чинить ошибки если есть |

---

## Используемые skills (последовательно)

| Этап | Skill | Действие |
|---|---|---|
| Stage 1 | `mcp__python-runner__run_command` | Запуск pretest скрипта |
| Stage 2 | `Edit` + `Read` | Правка ObjectModule.bsl |
| Stage 3 | `db-load-xml` + `db-update` | Partial load + UpdateDBCfg |
| Stage 3 | `mcp__1c-workerp__find_document_ref`, `unpost_document`, `post_document`, `execute_query` | Smoke 1 документа |
| Stage 4 | `mcp__python-runner__run_command` | Массовый перепровод |
| Stage 5 | `Edit` (json конфигурации) | Расширить dim_catalogs.json + refresh_mapping.py |
| Stage 5 | `Bash` или `mcp__python-runner` | ALTER TABLE + python main.py |
| Stage 6 | `mcp__powerbi-modeling-mcp__*` | Открыть/обновить/сохранить .pbix |
| Финал | `mcp__1c-workerp__execute_query` + SQL Server query | Acceptance verification |

## Используемые memories (контекст для ИИ)

- `olap_refresh_mapping_strict_whitelist.md` — обязательно расширить whitelist
- `pbix_sql_database_schema_cache.md` — обход PBI schema cache через native query
- `feedback_universal_register_fill.md` — паттерн логики проведения в регистр
- `feedback_no_typical_register_changes.md` — НЕ трогать типовые регистры (А_ОтчетPL_Свод кастомный — можно)
- `feedback_controller_verifies_no_reviewer_subagents.md` — контроллер сам проверяет, без code-reviewer субагентов
- `feedback_autonomous_finish_merge_push_pr.md` — финал git автономно

## Документация для обновления (в конце)

После успешного завершения всех 6 этапов:

1. `_Rarzrabotki/notebook/knowledge_PL/pl_report_architecture_analyst.md` — добавить раздел «Семантика знака суммы в А_ОтчетPL_Свод (с 2026-05-21)».
2. `_Rarzrabotki/notebook/knowledge_Olap/olap_powerbi_pl_pbix.md` — обновить колонки `А_Статьи_PL` (10 → 11) + добавить `ТипСтатьи` к таблице полей.
3. `_Rarzrabotki/notebook/knowledge_Olap/olap_sql_schema.md` — добавить `Type_Statya` к `Dim_PL_Articles`.
4. `_Rarzrabotki/notebook/knowledge_Olap/olap_changelog_2026_05.md` — entry для этой правки.
5. `_Rarzrabotki/notebook/knowledge_PL/KNOWLEDGE_MAP_PL.md` — обновить «Last update» + запись в «v3.x changes».

---

## TL;DR для финансиста

1. **Что делаем**: добавляем знак к суммам регистра PnL по семантике статьи (доход +, расход −).
2. **Зачем**: упрощается DAX мера `Маржинальный доход = Σ Сумма` без CASE/SIGN; согласовано с балансом и DDS, которые уже работают этим паттерном.
3. **Что меняется визуально**: визуалы Power BI «Прибыль/убыток», «P&L по статьям» становятся читаемее (отрицательные = расходы).
4. **Что НЕ меняется**: Σ |абсолют| остаётся, отчёт А_ОтчетPL (UI 1С) не затронут, исторические данные перепроводятся в фоне.
5. **Время выполнения**: ~30 мин ИИ-автономно (BSL edit 5 мин, db-load+update 2 мин, перепровод 28 × 5 сек = 2.5 мин, ETL 1 мин, .pbix 5 мин, верификация 10 мин).
