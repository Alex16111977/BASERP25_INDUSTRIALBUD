# Отчёт «Сверка ВзСС ЕРП ↔ ЗУП (zup_2)» — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Внешний отчёт `А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП` — сверка `РегистрНакопления.А_ВзаиморасчетыССотрудниками` ЕРП (Ф1+Ф2, разрез ФЛ) против `РегистрНакопления.ВзаиморасчетыСРаботниками` базы zup_2 (COM), с флагом «Только сотрудники из ЗУП».

**Architecture:** Точная копия паттерна `А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсBASБухгалтерия`: вся сверка в ObjectModule (`ПриКомпоновкеРезультата`), COM-доступ к zup_2 через `V83.COMConnector`, слияние таблиц по ключу `КодПоДРФО`, результат — один внешний набор СКД `Сотрудники_Сравнение`. Спека: `docs/superpowers/specs/2026-06-11-sverka-vzss-erp-zup-design.md`.

**Tech Stack:** 1С BSL (внешний отчёт .erf), СКД, V83.COMConnector, Python COM-тесты (Rule #-1), скилы erf-validate/erf-build.

**Базы:** ERP `Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"`, ЗУП `Srvr="SQLSERVER";Ref="zup_2";Usr="cfo";Pwd="2442"`.

**Критические грабли (из памяти проекта):**
- Запросы — сначала Python COM, потом BSL (Rule #-1). В Python-тестах границы периода — литералами `ДАТАВРЕМЯ(...)` (TZ-сдвиг datetime→COM).
- Не объявлять переменные `Ссылка`, `МетаДанные`, процедуру `Выполнить`.
- `python -c` с кириллицей портится → только `.py`-файлы UTF-8.
- Исходники 1С: UTF-8 **BOM + CRLF** (после Write перекодировать).
- erf-build/erf-validate — против реальной BaseERP (стаб-база падает «формат 2.17»).
- Worktree → в конце копировать всё в `C:\Configuration_downloads\BASERP25\` (Rule #4).

**File Structure (создаваемые файлы, пути от корня worktree):**

| Файл | Ответственность |
|---|---|
| `_Rarzrabotki/Python/test/test_sverka_vzss_zup_pretest.py` | Гейт 1: 4 запроса 1:1 как BSL (ЕРП, zup_2-регистр, zup_2-справочник, join-резолв) |
| `_Rarzrabotki/Python/test/test_sverka_vzss_zup_pretest2_signs.py` | Гейт 2: эмпирический знаковый маппинг zup_2→конвенция ЕРП |
| `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.xml` | Корневой XML метаданных отчёта |
| `.../А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Ext/ObjectModule.bsl` | Вся логика сверки |
| `.../А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Ext/Help.xml` + `Ext/Help/ru.html`, `uk.html` | Справка |
| `.../А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Templates/ОсновнаяСхемаКомпоновкиДанных.xml` + `.../Ext/Template.xml` | СКД |
| `_Rarzrabotki/Python/test/fix_encoding_sverka_zup.py` | Перекодировка BOM+CRLF |
| `_Rarzrabotki/Python/test/test_sverka_vzss_zup_acceptance.py` | Гейт 3: COM-зеркало слияния + проверка флага |

---

### Task 1: Pretest запросов (Python COM, Rule #-1)

**Files:**
- Create: `_Rarzrabotki/Python/test/test_sverka_vzss_zup_pretest.py`

- [ ] **Step 1.1: Написать pretest**

```python
# -*- coding: utf-8 -*-
"""Pretest (Rule #-1): 4 запроса отчёта А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.
Тексты 1:1 как будущий BSL. Период: декабрь 2025, границы — литералами ДАТАВРЕМЯ."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
zup = v8.Connect('Srvr="SQLSERVER";Ref="zup_2";Usr="cfo";Pwd="2442"')

FAILED = False

def run(conn, name, text):
    global FAILED
    q = conn.NewObject("Запрос")
    q.Text = text
    try:
        t = q.Execute().Выгрузить()
        print(f"[OK] {name}: строк={t.Количество()}")
        return t
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        print(f"[FAIL] {name}: {msg}")
        FAILED = True
        return None

# --- 1) ЕРП: А_ВзСС БЕЗ фильтра ФормаPL (Ф1+Ф2), разрез ФЛ + КодПоДРФО ---
Q_ERP = """ВЫБРАТЬ
	Ост.ФизическоеЛицо КАК ФизическоеЛицо,
	Ост.ФизическоеЛицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаВзаиморасчетовНачальныйОстаток) КАК НачальныйОстаток,
	СУММА(Ост.СуммаВзаиморасчетовПриход) КАК Начисления,
	СУММА(Ост.СуммаВзаиморасчетовРасход) КАК Выплаты,
	СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КонечныйОстаток
ИЗ
	РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , ) КАК Ост
СГРУППИРОВАТЬ ПО
	Ост.ФизическоеЛицо,
	Ост.ФизическоеЛицо.КодПоДРФО"""
t_erp = run(erp, "ЕРП А_ВзСС (Ф1+Ф2)", Q_ERP)
if t_erp is not None:
    total = t_erp.Количество()
    pust = 0
    s_kon = 0.0
    for i in range(total):
        r = t_erp.Получить(i)
        if not (r.ДРФО or "").strip():
            pust += 1
        s_kon += float(r.КонечныйОстаток or 0)
    print(f"  ЕРП: всего {total}, пустой ДРФО у {pust}, Σ КонОст = {s_kon:,.2f}")

# --- 2) zup_2: ВзаиморасчетыСРаботниками.ОстаткиИОбороты, условие ВТ по ДРФО ---
Q_ZUP = """ВЫБРАТЬ
	Ост.Физлицо.КодПоДРФО КАК ДРФО,
	Ост.Физлицо.Наименование КАК ФЛ_Наименование,
	СУММА(Ост.СуммаУпрНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаУпрПриход) КАК Приход,
	СУММА(Ост.СуммаУпрРасход) КАК Расход,
	СУММА(Ост.СуммаУпрКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.ВзаиморасчетыСРаботниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , Физлицо.КодПоДРФО <> "") КАК Ост
СГРУППИРОВАТЬ ПО
	Ост.Физлицо.КодПоДРФО,
	Ост.Физлицо.Наименование"""
t_zup = run(zup, "zup_2 ВзСР ОстаткиИОбороты", Q_ZUP)
if t_zup is not None:
    s_kon2 = 0.0
    for i in range(t_zup.Количество()):
        s_kon2 += float(t_zup.Получить(i).КонОст or 0)
    print(f"  zup_2: Σ КонОст = {s_kon2:,.2f}")

# --- 3) zup_2: справочник ФЛ — множество членства для флага ---
Q_FL = """ВЫБРАТЬ РАЗЛИЧНЫЕ
	ФЛ.КодПоДРФО КАК ДРФО
ИЗ
	Справочник.ФизическиеЛица КАК ФЛ
ГДЕ
	ФЛ.КодПоДРФО <> ""
	И НЕ ФЛ.ПометкаУдаления"""
run(zup, "zup_2 справочник ФЛ (ДРФО)", Q_FL)

# --- 4) ЕРП: join-резолв ДРФО -> ссылка ФЛ (МИНИМУМ + ЛЕВОЕ СОЕДИНЕНИЕ) ---
Q_JOIN = """ВЫБРАТЬ "2888011632" КАК ДРФО
ПОМЕСТИТЬ ВТ_ЗУП
;
////////////////////////////////////////////////////////////////////////////////
ВЫБРАТЬ
	ВТ.ДРФО КАК ДРФО,
	МИНИМУМ(ФЛ.Ссылка) КАК ФизическоеЛицо
ИЗ
	ВТ_ЗУП КАК ВТ
		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ФизическиеЛица КАК ФЛ
		ПО ВТ.ДРФО = ФЛ.КодПоДРФО
СГРУППИРОВАТЬ ПО
	ВТ.ДРФО"""
t_join = run(erp, "ЕРП join-резолв ДРФО->ФЛ", Q_JOIN)
if t_join is not None and t_join.Количество() == 1:
    fl = t_join.Получить(0).ФизическоеЛицо
    print(f"  Буша resolve: ЗначениеЗаполнено = {erp.ЗначениеЗаполнено(fl)}")

print("PRETEST: " + ("FAIL" if FAILED else "PASS"))
sys.exit(1 if FAILED else 0)
```

- [ ] **Step 1.2: Запустить**

Run: `python "_Rarzrabotki\Python\test\test_sverka_vzss_zup_pretest.py"` (cwd = worktree)
Expected: 4 × `[OK]`, `PRETEST: PASS`, exit 0. Если `[FAIL]` (например, имя колонки `СуммаУпр*` иное или у измерения другое имя) — поправить текст запроса в тесте до зелёного, итоговые тексты считать каноном для BSL (Task 4) и СКД-независимы.

- [ ] **Step 1.3: Commit**

```bash
git add _Rarzrabotki/Python/test/test_sverka_vzss_zup_pretest.py
git commit -m "test: pretest запросов отчёта сверки ВзСС ЕРП<->zup_2 (Rule #-1)"
```

---

### Task 2: Знаковый маппинг zup_2 → конвенция ЕРП (эмпирика)

**Files:**
- Create: `_Rarzrabotki/Python/test/test_sverka_vzss_zup_pretest2_signs.py`

Конвенция ЕРП: `+` = долг работодателя, Приход = начисление. Кандидаты:
- **T0 (identity):** НачОст=НачОст, Начисления=Приход, Выплаты=Расход, КонОст=КонОст.
- **T1 (инверсия+swap):** НачОст=−НачОст, Начисления=Расход, Выплаты=Приход, КонОст=−КонОст.

- [ ] **Step 2.1: Написать тест**

```python
# -*- coding: utf-8 -*-
"""Pretest2: эмпирический выбор знакового маппинга zup_2 -> конвенция ЕРП.
Месяц: ноябрь 2025 (эталон Буша, ДРФО 2888011632: начисление-нетто 52000-16940=35060)."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BUSHA = "2888011632"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
zup = v8.Connect('Srvr="SQLSERVER";Ref="zup_2";Usr="cfo";Pwd="2442"')

def fetch(conn, text, cols):
    q = conn.NewObject("Запрос")
    q.Text = text
    t = q.Execute().Выгрузить()
    out = {}
    for i in range(t.Количество()):
        r = t.Получить(i)
        drfo = (r.ДРФО or "").strip()
        if drfo:
            out[drfo] = tuple(float(getattr(r, c) or 0) for c in cols)
    return out

ERP_ROWS = fetch(erp, """ВЫБРАТЬ
	Ост.ФизическоеЛицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаВзаиморасчетовНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаВзаиморасчетовПриход) КАК Приход,
	СУММА(Ост.СуммаВзаиморасчетовРасход) КАК Расход,
	СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 11, 1), ДАТАВРЕМЯ(2025, 11, 30, 23, 59, 59), , , ) КАК Ост
СГРУППИРОВАТЬ ПО Ост.ФизическоеЛицо.КодПоДРФО""", ("НачОст", "Приход", "Расход", "КонОст"))

ZUP_ROWS = fetch(zup, """ВЫБРАТЬ
	Ост.Физлицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаУпрНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаУпрПриход) КАК Приход,
	СУММА(Ост.СуммаУпрРасход) КАК Расход,
	СУММА(Ост.СуммаУпрКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.ВзаиморасчетыСРаботниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 11, 1), ДАТАВРЕМЯ(2025, 11, 30, 23, 59, 59), , , Физлицо.КодПоДРФО <> "") КАК Ост
СГРУППИРОВАТЬ ПО Ост.Физлицо.КодПоДРФО""", ("НачОст", "Приход", "Расход", "КонОст"))

print(f"ЕРП ФЛ: {len(ERP_ROWS)}, zup_2 ФЛ: {len(ZUP_ROWS)}, общих: {len(set(ERP_ROWS) & set(ZUP_ROWS))}")
print(f"Σ КонОст ЕРП   = {sum(v[3] for v in ERP_ROWS.values()):,.2f}")
print(f"Σ КонОст zup_2 = {sum(v[3] for v in ZUP_ROWS.values()):,.2f}")

print(f"\nБуша ЕРП  : {ERP_ROWS.get(BUSHA)}")
print(f"Буша zup_2: {ZUP_ROWS.get(BUSHA)}  # начисление-нетто ноября = 35060 -> та сторона (Приход или Расход), где ~35060, и есть «начисления»")

def transform(row, mode):
    n, p, r, k = row
    if mode == "T0":
        return (n, p, r, k)
    return (-n, r, p, -k)  # T1: инверсия остатков, swap приход/расход

for mode in ("T0", "T1"):
    diffs = []
    for drfo in set(ERP_ROWS) & set(ZUP_ROWS):
        e = ERP_ROWS[drfo]
        z = transform(ZUP_ROWS[drfo], mode)
        diffs.append(abs(e[3] - z[3]))
    diffs.sort()
    exact = sum(1 for d in diffs if d <= 0.01)
    med = diffs[len(diffs) // 2] if diffs else 0
    print(f"{mode}: совпало КонОст (<=0.01) {exact}/{len(diffs)}, медиана |Δ| = {med:,.2f}, Σ|Δ| = {sum(diffs):,.2f}")

print("\nВЕРДИКТ: выбрать маппинг с большим exact/меньшей медианой; сверить с Бушей вручную по выводу выше.")
```

- [ ] **Step 2.2: Запустить и зафиксировать вердикт**

Run: `python "_Rarzrabotki\Python\test\test_sverka_vzss_zup_pretest2_signs.py"`
Expected: вывод без исключений. Выбрать маппинг по двум независимым сигналам: (а) у Буши сторона с ~35060 = «начисления»; (б) больший `exact`-счётчик / меньшая медиана |Δ| КонОст. Дописать в конец файла комментарий `# ВЕРДИКТ: T0` или `# ВЕРДИКТ: T1` с датой. Этот вердикт используется в Task 4 (блок маппинга BSL) и Task 7 (acceptance).

- [ ] **Step 2.3: Commit**

```bash
git add _Rarzrabotki/Python/test/test_sverka_vzss_zup_pretest2_signs.py
git commit -m "test: знаковый маппинг zup_2->ЕРП для отчёта сверки ВзСС (вердикт зафиксирован)"
```

---

### Task 3: Каркас отчёта — корневой XML и Help

**Files:**
- Create: `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.xml`
- Create: `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Ext/Help.xml` (копия из эталона)
- Create: `.../Ext/Help/ru.html`, `.../Ext/Help/uk.html`

- [ ] **Step 3.1: Корневой XML** (UUID новые, ClassId — константа класса ExternalReport из эталона)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">
	<ExternalReport uuid="b7e3d2a4-91c8-4f5e-8b06-2d74a9c1e583">
		<InternalInfo>
			<xr:ContainedObject>
				<xr:ClassId>e41aff26-25cf-4bb6-b6c1-3f478a75f374</xr:ClassId>
				<xr:ObjectId>4f8c61b9-2e07-4a83-95d4-7c310fb6e8a2</xr:ObjectId>
			</xr:ContainedObject>
			<xr:GeneratedType name="ExternalReportObject.А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП" category="Object">
				<xr:TypeId>8a25c7e3-64f9-4d18-b3a0-19e5d2c84b76</xr:TypeId>
				<xr:ValueId>d19b84f6-3c52-4e07-a8c1-650f97e23ab4</xr:ValueId>
			</xr:GeneratedType>
		</InternalInfo>
		<Properties>
			<Name>А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП</Name>
			<Synonym>
				<v8:item>
					<v8:lang>uk</v8:lang>
					<v8:content>Звірка взаєморозрахунків зі співробітниками: ЕРП (Ф1+Ф2) ↔ ЗУП</v8:content>
				</v8:item>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Сверка взаиморасчётов с сотрудниками: ЕРП (Ф1+Ф2) ↔ ЗУП</v8:content>
				</v8:item>
			</Synonym>
			<Comment/>
			<DefaultForm>CommonForm.ФормаОтчета</DefaultForm>
			<AuxiliaryForm/>
			<MainDataCompositionSchema>ExternalReport.А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.Template.ОсновнаяСхемаКомпоновкиДанных</MainDataCompositionSchema>
			<DefaultSettingsForm>CommonForm.ФормаНастроекОтчета</DefaultSettingsForm>
			<AuxiliarySettingsForm/>
			<DefaultVariantForm>CommonForm.ФормаВариантаОтчета</DefaultVariantForm>
			<VariantsStorage/>
			<SettingsStorage/>
		</Properties>
		<ChildObjects>
			<Template>ОсновнаяСхемаКомпоновкиДанных</Template>
		</ChildObjects>
	</ExternalReport>
</MetaDataObject>
```

- [ ] **Step 3.2: Help**

Скопировать `Help.xml` эталона без правок (он лишь ссылается на страницы ru/uk):

```powershell
$src = "_Rarzrabotki\Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсBASБухгалтерия\Ext"
$dst = "_Rarzrabotki\Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП\Ext"
New-Item -ItemType Directory -Force "$dst\Help"
Copy-Item "$src\Help.xml" "$dst\Help.xml"
```

Записать `Ext/Help/ru.html` (uk.html — то же по-украински):

```html
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"><html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body>
<h2>Сверка взаиморасчётов с сотрудниками: ЕРП (Ф1+Ф2) ↔ ЗУП</h2>
<p>Диагностический отчёт. Сравнивает регистр ЕРП <b>А_ВзаиморасчетыССотрудниками</b>
(обе ФормыPL, т.е. «Общий» = Форма1 + Форма2) с регистром <b>ВзаиморасчетыСРаботниками</b>
базы zup_2 (подключение COM). Разрез — физическое лицо, ключ сопоставления — код ДРФО.</p>
<p>Показатели за период: начальный остаток, начисления, выплаты/удержания, конечный остаток —
по каждой стороне и разница (ЕРП − ЗУП). Знаки приведены к конвенции ЕРП: «+» = долг работодателя.</p>
<p>Флаг <b>«Только сотрудники из ЗУП»</b> (по умолчанию включён) оставляет только физлиц,
присутствующих в справочнике «Физические лица» базы zup_2 (с непустым ДРФО, без пометки удаления).
Если флаг снят — выводится полная картина, включая строки, существующие только в ЕРП.</p>
</body></html>
```

```html
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"><html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body>
<h2>Звірка взаєморозрахунків зі співробітниками: ЕРП (Ф1+Ф2) ↔ ЗУП</h2>
<p>Діагностичний звіт. Порівнює регістр ЕРП <b>А_ВзаиморасчетыССотрудниками</b>
(обидві ФормиPL, тобто «Загальний» = Форма1 + Форма2) з регістром <b>ВзаиморасчетыСРаботниками</b>
бази zup_2 (підключення COM). Розріз — фізична особа, ключ зіставлення — код ДРФО.</p>
<p>Показники за період: початковий залишок, нарахування, виплати/утримання, кінцевий залишок —
по кожній стороні та різниця (ЕРП − ЗУП). Знаки приведені до конвенції ЕРП: «+» = борг роботодавця.</p>
<p>Прапорець <b>«Тільки співробітники з ЗУП»</b> (типово увімкнений) залишає лише фізосіб,
наявних у довіднику «Фізичні особи» бази zup_2 (з непорожнім ДРФО, без позначки видалення).
Якщо прапорець знятий — виводиться повна картина, включно з рядками, що існують лише в ЕРП.</p>
</body></html>
```

- [ ] **Step 3.3: Commit**

```bash
git add "_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.xml" "_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП"
git commit -m "feat: каркас отчёта А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП (root XML + Help)"
```

---

### Task 4: СКД Template

**Files:**
- Create: `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Templates/ОсновнаяСхемаКомпоновкиДанных.xml`
- Create: `.../Templates/ОсновнаяСхемаКомпоновкиДанных/Ext/Template.xml`

- [ ] **Step 4.1: Дескриптор макета** — скопировать `Templates/ОсновнаяСхемаКомпоновкиДанных.xml` эталона 1:1 (он generic: имя+тип DataCompositionSchema), заменив в нём uuid на `c5d28f17-4b96-4e03-a8d1-72f60c9b3e45`. Если в эталонном файле uuid отсутствует — копировать без правок.

- [ ] **Step 4.2: Template.xml** — полный файл (отличия от эталона: нет ОргБух, суффикс `_ЗУП`, поля `ДРФО` и `ФЛ_НаименованиеЗУП`, параметр `ТолькоСотрудникиИзЗУП`, группировки ФЛ → ФЛ_НаименованиеЗУП):

```xml
<?xml version="1.0" encoding="UTF-8"?>
<DataCompositionSchema xmlns="http://v8.1c.ru/8.1/data-composition-system/schema" xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<dataSource>
		<name>ИсточникДанных1</name>
		<dataSourceType>Local</dataSourceType>
	</dataSource>
	<dataSet xsi:type="DataSetObject">
		<name>Сотрудники_Сравнение</name>
		<field xsi:type="DataSetFieldField">
			<dataPath>ФизическоеЛицо</dataPath>
			<field>ФизическоеЛицо</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Физическое лицо</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config">d5p1:CatalogRef.ФизическиеЛица</v8:Type>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>ДРФО</dataPath>
			<field>ДРФО</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>ДРФО (ключ)</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:string</v8:Type>
				<v8:StringQualifiers>
					<v8:Length>20</v8:Length>
					<v8:AllowedLength>Variable</v8:AllowedLength>
				</v8:StringQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>ФЛ_НаименованиеЗУП</dataPath>
			<field>ФЛ_НаименованиеЗУП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>ФЛ (ЗУП)</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:string</v8:Type>
				<v8:StringQualifiers>
					<v8:Length>150</v8:Length>
					<v8:AllowedLength>Variable</v8:AllowedLength>
				</v8:StringQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>НачальныйОстаток_ЕРП</dataPath>
			<field>НачальныйОстаток_ЕРП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Нач.остаток ЕРП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>НачальныйОстаток_ЗУП</dataPath>
			<field>НачальныйОстаток_ЗУП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Нач.остаток ЗУП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>РазницаНачальныйОстаток</dataPath>
			<field>РазницаНачальныйОстаток</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Разн. нач.остаток</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>Начисления_ЕРП</dataPath>
			<field>Начисления_ЕРП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Начисления ЕРП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>Начисления_ЗУП</dataPath>
			<field>Начисления_ЗУП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Начисления ЗУП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>РазницаНачисления</dataPath>
			<field>РазницаНачисления</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Разн. начисления</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>Выплаты_ЕРП</dataPath>
			<field>Выплаты_ЕРП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Выплаты ЕРП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>Выплаты_ЗУП</dataPath>
			<field>Выплаты_ЗУП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Выплаты ЗУП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>РазницаВыплаты</dataPath>
			<field>РазницаВыплаты</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Разн. выплаты</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>КонечныйОстаток_ЕРП</dataPath>
			<field>КонечныйОстаток_ЕРП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Кон.остаток ЕРП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>КонечныйОстаток_ЗУП</dataPath>
			<field>КонечныйОстаток_ЗУП</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Кон.остаток ЗУП</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<field xsi:type="DataSetFieldField">
			<dataPath>РазницаКонечныйОстаток</dataPath>
			<field>РазницаКонечныйОстаток</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Разн. кон.остаток</v8:content>
				</v8:item>
			</title>
			<valueType>
				<v8:Type>xs:decimal</v8:Type>
				<v8:NumberQualifiers>
					<v8:Digits>15</v8:Digits>
					<v8:FractionDigits>2</v8:FractionDigits>
					<v8:AllowedSign>Any</v8:AllowedSign>
				</v8:NumberQualifiers>
			</valueType>
		</field>
		<dataSource>ИсточникДанных1</dataSource>
		<objectName>Сотрудники_Сравнение</objectName>
	</dataSet>
	<totalField>
		<dataPath>НачальныйОстаток_ЕРП</dataPath>
		<expression>Сумма(НачальныйОстаток_ЕРП)</expression>
	</totalField>
	<totalField>
		<dataPath>НачальныйОстаток_ЗУП</dataPath>
		<expression>Сумма(НачальныйОстаток_ЗУП)</expression>
	</totalField>
	<totalField>
		<dataPath>РазницаНачальныйОстаток</dataPath>
		<expression>Сумма(РазницаНачальныйОстаток)</expression>
	</totalField>
	<totalField>
		<dataPath>Начисления_ЕРП</dataPath>
		<expression>Сумма(Начисления_ЕРП)</expression>
	</totalField>
	<totalField>
		<dataPath>Начисления_ЗУП</dataPath>
		<expression>Сумма(Начисления_ЗУП)</expression>
	</totalField>
	<totalField>
		<dataPath>РазницаНачисления</dataPath>
		<expression>Сумма(РазницаНачисления)</expression>
	</totalField>
	<totalField>
		<dataPath>Выплаты_ЕРП</dataPath>
		<expression>Сумма(Выплаты_ЕРП)</expression>
	</totalField>
	<totalField>
		<dataPath>Выплаты_ЗУП</dataPath>
		<expression>Сумма(Выплаты_ЗУП)</expression>
	</totalField>
	<totalField>
		<dataPath>РазницаВыплаты</dataPath>
		<expression>Сумма(РазницаВыплаты)</expression>
	</totalField>
	<totalField>
		<dataPath>КонечныйОстаток_ЕРП</dataPath>
		<expression>Сумма(КонечныйОстаток_ЕРП)</expression>
	</totalField>
	<totalField>
		<dataPath>КонечныйОстаток_ЗУП</dataPath>
		<expression>Сумма(КонечныйОстаток_ЗУП)</expression>
	</totalField>
	<totalField>
		<dataPath>РазницаКонечныйОстаток</dataPath>
		<expression>Сумма(РазницаКонечныйОстаток)</expression>
	</totalField>
	<parameter>
		<name>Период</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Период</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>v8:StandardPeriod</v8:Type>
		</valueType>
		<value xsi:type="v8:StandardPeriod">
			<v8:variant xsi:type="v8:StandardPeriodVariant">ThisMonth</v8:variant>
		</value>
		<useRestriction>false</useRestriction>
	</parameter>
	<parameter>
		<name>НачалоПериода</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Начало периода</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:dateTime</v8:Type>
			<v8:DateQualifiers>
				<v8:DateFractions>DateTime</v8:DateFractions>
			</v8:DateQualifiers>
		</valueType>
		<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>
		<useRestriction>true</useRestriction>
		<expression>&amp;Период.ДатаНачала</expression>
	</parameter>
	<parameter>
		<name>КонецПериода</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Конец периода</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:dateTime</v8:Type>
			<v8:DateQualifiers>
				<v8:DateFractions>DateTime</v8:DateFractions>
			</v8:DateQualifiers>
		</valueType>
		<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>
		<useRestriction>true</useRestriction>
		<expression>КОНЕЦПЕРИОДА(&amp;Период.ДатаОкончания, "ДЕНЬ")</expression>
	</parameter>
	<parameter>
		<name>ТолькоСотрудникиИзЗУП</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Только сотрудники из ЗУП</v8:content>
			</v8:item>
			<v8:item>
				<v8:lang>uk</v8:lang>
				<v8:content>Тільки співробітники з ЗУП</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:boolean</v8:Type>
		</valueType>
		<value xsi:type="xs:boolean">true</value>
		<useRestriction>false</useRestriction>
	</parameter>
	<settingsVariant>
		<dcsset:name>Основной</dcsset:name>
		<dcsset:presentation xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Сверка взаиморасчётов с сотрудниками: ЕРП (Ф1+Ф2) vs ЗУП</v8:content>
			</v8:item>
			<v8:item>
				<v8:lang>uk</v8:lang>
				<v8:content>Звірка взаєморозрахунків зі співробітниками: ЕРП (Ф1+Ф2) vs ЗУП</v8:content>
			</v8:item>
		</dcsset:presentation>
		<dcsset:settings xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows">
			<dcsset:selection>
				<dcsset:item xsi:type="dcsset:SelectedItemFolder">
					<dcsset:lwsTitle>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Нач.остаток</v8:content>
						</v8:item>
					</dcsset:lwsTitle>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>НачальныйОстаток_ЕРП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>НачальныйОстаток_ЗУП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>РазницаНачальныйОстаток</dcsset:field>
					</dcsset:item>
					<dcsset:placement>Auto</dcsset:placement>
				</dcsset:item>
				<dcsset:item xsi:type="dcsset:SelectedItemFolder">
					<dcsset:lwsTitle>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Начисления (долг растёт)</v8:content>
						</v8:item>
					</dcsset:lwsTitle>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>Начисления_ЕРП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>Начисления_ЗУП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>РазницаНачисления</dcsset:field>
					</dcsset:item>
					<dcsset:placement>Auto</dcsset:placement>
				</dcsset:item>
				<dcsset:item xsi:type="dcsset:SelectedItemFolder">
					<dcsset:lwsTitle>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Выплаты и удержания (долг уменьшается)</v8:content>
						</v8:item>
					</dcsset:lwsTitle>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>Выплаты_ЕРП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>Выплаты_ЗУП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>РазницаВыплаты</dcsset:field>
					</dcsset:item>
					<dcsset:placement>Auto</dcsset:placement>
				</dcsset:item>
				<dcsset:item xsi:type="dcsset:SelectedItemFolder">
					<dcsset:lwsTitle>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Кон.остаток</v8:content>
						</v8:item>
					</dcsset:lwsTitle>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>КонечныйОстаток_ЕРП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>КонечныйОстаток_ЗУП</dcsset:field>
					</dcsset:item>
					<dcsset:item xsi:type="dcsset:SelectedItemField">
						<dcsset:field>РазницаКонечныйОстаток</dcsset:field>
					</dcsset:item>
					<dcsset:placement>Auto</dcsset:placement>
				</dcsset:item>
			</dcsset:selection>
			<dcsset:dataParameters>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:parameter>Период</dcscor:parameter>
					<dcscor:value xsi:type="v8:StandardPeriod">
						<v8:variant xsi:type="v8:StandardPeriodVariant">ThisMonth</v8:variant>
					</dcscor:value>
					<dcsset:userSettingID>a1b2c3d4-0001-4002-8003-000000000011</dcsset:userSettingID>
				</dcscor:item>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:parameter>ТолькоСотрудникиИзЗУП</dcscor:parameter>
					<dcscor:value xsi:type="xs:boolean">true</dcscor:value>
					<dcsset:userSettingID>a1b2c3d4-0002-4002-8003-000000000012</dcsset:userSettingID>
				</dcscor:item>
			</dcsset:dataParameters>
			<dcsset:item xsi:type="dcsset:StructureItemGroup">
				<dcsset:groupItems>
					<dcsset:item xsi:type="dcsset:GroupItemField">
						<dcsset:field>ФизическоеЛицо</dcsset:field>
						<dcsset:groupType>Items</dcsset:groupType>
						<dcsset:periodAdditionType>None</dcsset:periodAdditionType>
						<dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>
						<dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>
					</dcsset:item>
				</dcsset:groupItems>
				<dcsset:order>
					<dcsset:item xsi:type="dcsset:OrderItemAuto"/>
				</dcsset:order>
				<dcsset:selection>
					<dcsset:item xsi:type="dcsset:SelectedItemAuto"/>
				</dcsset:selection>
				<dcsset:item xsi:type="dcsset:StructureItemGroup">
					<dcsset:groupItems>
						<dcsset:item xsi:type="dcsset:GroupItemField">
							<dcsset:field>ФЛ_НаименованиеЗУП</dcsset:field>
							<dcsset:groupType>Items</dcsset:groupType>
							<dcsset:periodAdditionType>None</dcsset:periodAdditionType>
							<dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>
							<dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>
						</dcsset:item>
					</dcsset:groupItems>
					<dcsset:order>
						<dcsset:item xsi:type="dcsset:OrderItemAuto"/>
					</dcsset:order>
					<dcsset:selection>
						<dcsset:item xsi:type="dcsset:SelectedItemAuto"/>
					</dcsset:selection>
				</dcsset:item>
			</dcsset:item>
		</dcsset:settings>
	</settingsVariant>
</DataCompositionSchema>
```

Поле `ДРФО` в выборку варианта намеренно НЕ включено (доступно пользователю для добавления). НЕ добавлять корзины-периоды (День/Месяц/...).

- [ ] **Step 4.3: Commit**

```bash
git add "_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП"
git commit -m "feat: СКД отчёта сверки ВзСС ЕРП<->ЗУП (поля, итоги, параметр ТолькоСотрудникиИзЗУП)"
```

---

### Task 5: ObjectModule.bsl

**Files:**
- Create: `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Ext/ObjectModule.bsl`

- [ ] **Step 5.1: Написать модуль.** Тексты запросов — РОВНО те, что прошли pretest (Task 1). Блок маппинга в `ПолучитьОстаткиЗУП` — вариант по вердикту Task 2 (ниже показаны оба, оставить один):

```bsl
#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ОбработчикиСобытий

Процедура ПриКомпоновкеРезультата(ДокументРезультат, ДанныеРасшифровки, СтандартнаяОбработка)

	СтандартнаяОбработка = Ложь;
	УстановитьПривилегированныйРежим(Истина);

	ПараметрПериодОтчета = КомпоновкаДанныхКлиентСервер.ПолучитьПараметр(КомпоновщикНастроек, "Период").Значение;
	НачалоПериода = ПараметрПериодОтчета.ДатаНачала;
	КонецПериода = КонецДня(ПараметрПериодОтчета.ДатаОкончания);

	ТолькоСотрудникиИзЗУП = Ложь;
	ПарамФлаг = КомпоновкаДанныхКлиентСервер.ПолучитьПараметр(КомпоновщикНастроек, "ТолькоСотрудникиИзЗУП");
	Если ПарамФлаг <> Неопределено Тогда
		ТолькоСотрудникиИзЗУП = (ПарамФлаг.Использование И ПарамФлаг.Значение = Истина);
	КонецЕсли;

	// 1. ЕРП: А_ВзаиморасчетыССотрудниками БЕЗ фильтра ФормаPL (Ф1+Ф2 = "Общий"), разрез ФЛ
	ТаблицаЕРП = ПолучитьОстаткиЕРП(НачалоПериода, КонецПериода);

	// 2. zup_2: ВзаиморасчетыСРаботниками (разрез Физлицо.КодПоДРФО) + множество ДРФО справочника
	СтрокаПодключенияЗУП = "Srvr=""SQLSERVER"";Ref=""zup_2"";Usr=""cfo"";Pwd=""2442""";
	V8 = Новый COMОбъект("V83.COMConnector");
	Попытка
		V83 = V8.Connect(СокрЛП(СтрокаПодключенияЗУП));
	Исключение
		Сообщить("Помилка підключення до zup_2: " + ОписаниеОшибки());
		Возврат;
	КонецПопытки;

	ТаблицаЗУП = ПолучитьОстаткиЗУП(V83, НачалоПериода, КонецПериода);
	МножествоДРФОЗУП = ПолучитьДРФОСправочникаЗУП(V83);
	V83 = Неопределено;
	V8 = Неопределено;

	// 2.1 Конвертация: ДРФО -> ссылки ФЛ ЕРП
	ТаблицаЗУП = КонвертироватьТЗ_ЗУП(ТаблицаЗУП);

	// 3. Объединение по ДРФО + фильтр "Только сотрудники из ЗУП"
	ТаблицаСравнение = ОбъединитьТаблицы(ТаблицаЕРП, ТаблицаЗУП, МножествоДРФОЗУП, ТолькоСотрудникиИзЗУП);

	// 4. Передача в СКД
	ВнешниеНаборыДанных = Новый Структура("Сотрудники_Сравнение", ТаблицаСравнение);

	НастройкиОтчета = КомпоновщикНастроек.ПолучитьНастройки();
	КомпоновщикМакета = Новый КомпоновщикМакетаКомпоновкиДанных;
	МакетКомпоновки = КомпоновщикМакета.Выполнить(СхемаКомпоновкиДанных, НастройкиОтчета, ДанныеРасшифровки);
	ПроцессорКомпоновки = Новый ПроцессорКомпоновкиДанных;
	ПроцессорКомпоновки.Инициализировать(МакетКомпоновки, ВнешниеНаборыДанных, ДанныеРасшифровки);
	ПроцессорВывода = Новый ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент;
	ПроцессорВывода.УстановитьДокумент(ДокументРезультат);
	ПроцессорВывода.Вывести(ПроцессорКомпоновки);

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// ЕРП: РегНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты БЕЗ фильтра ФормаPL.
// Конвенция: + = долг работодателя; Приход = начисление, Расход = выплата/удержание.
Функция ПолучитьОстаткиЕРП(НачалоПериода, КонецПериода)

	ТипЧисло = ОбщегоНазначения.ОписаниеТипаЧисло(15, 2);
	ТипСтрока = Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(150));

	Результат = Новый ТаблицаЗначений;
	Результат.Колонки.Добавить("ДРФО", ТипСтрока);
	Результат.Колонки.Добавить("ФизическоеЛицо", Новый ОписаниеТипов("СправочникСсылка.ФизическиеЛица"));
	Результат.Колонки.Добавить("НачальныйОстаток_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("Начисления_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("Выплаты_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("КонечныйОстаток_ЕРП", ТипЧисло);

	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("НачалоПериода", НачалоПериода);
	Запрос.УстановитьПараметр("КонецПериода", КонецПериода);
	Запрос.Текст =
	"ВЫБРАТЬ
	|	Ост.ФизическоеЛицо КАК ФизическоеЛицо,
	|	Ост.ФизическоеЛицо.КодПоДРФО КАК ДРФО,
	|	СУММА(Ост.СуммаВзаиморасчетовНачальныйОстаток) КАК НачальныйОстаток,
	|	СУММА(Ост.СуммаВзаиморасчетовПриход) КАК Начисления,
	|	СУММА(Ост.СуммаВзаиморасчетовРасход) КАК Выплаты,
	|	СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КонечныйОстаток
	|ИЗ
	|	РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(
	|		&НачалоПериода, &КонецПериода, , , ) КАК Ост
	|СГРУППИРОВАТЬ ПО
	|	Ост.ФизическоеЛицо,
	|	Ост.ФизическоеЛицо.КодПоДРФО";

	Попытка
		Выборка = Запрос.Выполнить().Выбрать();
	Исключение
		Сообщить("Помилка запиту ЕРП (А_ВзаиморасчетыССотрудниками): " + ОписаниеОшибки());
		Возврат Результат;
	КонецПопытки;

	Пока Выборка.Следующий() Цикл
		НоваяСтрока = Результат.Добавить();
		НоваяСтрока.ДРФО = СокрЛП(Выборка.ДРФО);
		НоваяСтрока.ФизическоеЛицо = Выборка.ФизическоеЛицо;
		НоваяСтрока.НачальныйОстаток_ЕРП = Выборка.НачальныйОстаток;
		НоваяСтрока.Начисления_ЕРП = Выборка.Начисления;
		НоваяСтрока.Выплаты_ЕРП = Выборка.Выплаты;
		НоваяСтрока.КонечныйОстаток_ЕРП = Выборка.КонечныйОстаток;
	КонецЦикла;

	Возврат Результат;

КонецФункции

// zup_2 (COM): РегНакопления.ВзаиморасчетыСРаботниками.ОстаткиИОбороты, разрез Физлицо.КодПоДРФО.
// ТЗ содержит ТОЛЬКО примитивы (строки/числа) -> сериализация между базами безопасна.
// Маппинг знаков к конвенции ЕРП зафиксирован тестом test_sverka_vzss_zup_pretest2_signs.py.
Функция ПолучитьОстаткиЗУП(V83, НачалоПериода, КонецПериода)

	ТипЧисло = ОбщегоНазначения.ОписаниеТипаЧисло(15, 2);
	ТипСтрока = Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(150));

	Результат = Новый ТаблицаЗначений;
	Результат.Колонки.Добавить("ДРФО", ТипСтрока);
	Результат.Колонки.Добавить("ФЛ_Наименование", ТипСтрока);
	Результат.Колонки.Добавить("НачальныйОстаток_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("Начисления_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("Выплаты_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("КонечныйОстаток_ЗУП", ТипЧисло);

	ЗапросЗУП = V83.NewObject("Запрос");
	ЗапросЗУП.УстановитьПараметр("НачалоПериода", НачалоПериода);
	ЗапросЗУП.УстановитьПараметр("КонецПериода", КонецПериода);
	ЗапросЗУП.Текст =
	"ВЫБРАТЬ
	|	Ост.Физлицо.КодПоДРФО КАК ДРФО,
	|	Ост.Физлицо.Наименование КАК ФЛ_Наименование,
	|	СУММА(Ост.СуммаУпрНачальныйОстаток) КАК НачОст,
	|	СУММА(Ост.СуммаУпрПриход) КАК Приход,
	|	СУММА(Ост.СуммаУпрРасход) КАК Расход,
	|	СУММА(Ост.СуммаУпрКонечныйОстаток) КАК КонОст
	|ИЗ
	|	РегистрНакопления.ВзаиморасчетыСРаботниками.ОстаткиИОбороты(
	|		&НачалоПериода, &КонецПериода, , , Физлицо.КодПоДРФО <> """""""") КАК Ост
	|СГРУППИРОВАТЬ ПО
	|	Ост.Физлицо.КодПоДРФО,
	|	Ост.Физлицо.Наименование";

	Попытка
		ТЗ = ЗапросЗУП.Выполнить().Выгрузить();
		стр_ТЗ = V83.ЗначениеВСтрокуВнутр(ТЗ);
		ТаблицаЗУП = ЗначениеИзСтрокиВнутр(стр_ТЗ);
	Исключение
		Сообщить("Помилка запиту zup_2 (ВзаиморасчетыСРаботниками): " + ОписаниеОшибки());
		Возврат Результат;
	КонецПопытки;

	Для Каждого СтрЗ Из ТаблицаЗУП Цикл
		ДРФО = СокрЛП(СтрЗ.ДРФО);
		Если ПустаяСтрока(ДРФО) Тогда
			Продолжить;
		КонецЕсли;
		НоваяСтрока = Результат.Добавить();
		НоваяСтрока.ДРФО = ДРФО;
		НоваяСтрока.ФЛ_Наименование = СокрЛП(СтрЗ.ФЛ_Наименование);
		// === Маппинг знаков: ВАРИАНТ ПО ВЕРДИКТУ Task 2 (оставить один) ===
		// T0 (identity):
		//НоваяСтрока.НачальныйОстаток_ЗУП = СтрЗ.НачОст;
		//НоваяСтрока.Начисления_ЗУП = СтрЗ.Приход;
		//НоваяСтрока.Выплаты_ЗУП = СтрЗ.Расход;
		//НоваяСтрока.КонечныйОстаток_ЗУП = СтрЗ.КонОст;
		// T1 (инверсия остатков + swap приход/расход):
		НоваяСтрока.НачальныйОстаток_ЗУП = -СтрЗ.НачОст;
		НоваяСтрока.Начисления_ЗУП = СтрЗ.Расход;
		НоваяСтрока.Выплаты_ЗУП = СтрЗ.Приход;
		НоваяСтрока.КонечныйОстаток_ЗУП = -СтрЗ.КонОст;
	КонецЦикла;

	Возврат Результат;

КонецФункции

// zup_2 (COM): множество непустых КодПоДРФО справочника ФизическиеЛица (без пометки удаления).
// Семантика флага "Только сотрудники из ЗУП": членство в СПРАВОЧНИКЕ, а не наличие движений.
Функция ПолучитьДРФОСправочникаЗУП(V83)

	Множество = Новый Соответствие;

	ЗапросФЛ = V83.NewObject("Запрос");
	ЗапросФЛ.Текст =
	"ВЫБРАТЬ РАЗЛИЧНЫЕ
	|	ФЛ.КодПоДРФО КАК ДРФО
	|ИЗ
	|	Справочник.ФизическиеЛица КАК ФЛ
	|ГДЕ
	|	ФЛ.КодПоДРФО <> """"""""
	|	И НЕ ФЛ.ПометкаУдаления";

	Попытка
		ТЗ = ЗапросФЛ.Выполнить().Выгрузить();
		стр_ТЗ = V83.ЗначениеВСтрокуВнутр(ТЗ);
		ТаблицаФЛ = ЗначениеИзСтрокиВнутр(стр_ТЗ);
	Исключение
		Сообщить("Помилка запиту zup_2 (справочник ФизическиеЛица): " + ОписаниеОшибки());
		Возврат Множество;
	КонецПопытки;

	Для Каждого СтрФ Из ТаблицаФЛ Цикл
		Множество.Вставить(СокрЛП(СтрФ.ДРФО), Истина);
	КонецЦикла;

	Возврат Множество;

КонецФункции

// Конвертация zup_2: ДРФО -> ФизическиеЛица.Ссылка ЕРП (ВТ + ЛЕВОЕ СОЕДИНЕНИЕ, МИНИМУМ от дублей).
Функция КонвертироватьТЗ_ЗУП(ТаблицаЗУП)

	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("ТЗ", ТаблицаЗУП);
	Запрос.Текст =
	"ВЫБРАТЬ
	|	ВТ.ДРФО КАК ДРФО,
	|	ВТ.ФЛ_Наименование КАК ФЛ_Наименование,
	|	ВТ.НачальныйОстаток_ЗУП КАК НачальныйОстаток_ЗУП,
	|	ВТ.Начисления_ЗУП КАК Начисления_ЗУП,
	|	ВТ.Выплаты_ЗУП КАК Выплаты_ЗУП,
	|	ВТ.КонечныйОстаток_ЗУП КАК КонечныйОстаток_ЗУП
	|ПОМЕСТИТЬ ВТ_ЗУП
	|ИЗ
	|	&ТЗ КАК ВТ
	|;
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ВТ.ДРФО КАК ДРФО,
	|	ВТ.ФЛ_Наименование КАК ФЛ_Наименование,
	|	МИНИМУМ(ФЛ.Ссылка) КАК ФизическоеЛицо,
	|	ВТ.НачальныйОстаток_ЗУП КАК НачальныйОстаток_ЗУП,
	|	ВТ.Начисления_ЗУП КАК Начисления_ЗУП,
	|	ВТ.Выплаты_ЗУП КАК Выплаты_ЗУП,
	|	ВТ.КонечныйОстаток_ЗУП КАК КонечныйОстаток_ЗУП
	|ИЗ
	|	ВТ_ЗУП КАК ВТ
	|		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ФизическиеЛица КАК ФЛ
	|		ПО ВТ.ДРФО = ФЛ.КодПоДРФО
	|СГРУППИРОВАТЬ ПО
	|	ВТ.ДРФО,
	|	ВТ.ФЛ_Наименование,
	|	ВТ.НачальныйОстаток_ЗУП,
	|	ВТ.Начисления_ЗУП,
	|	ВТ.Выплаты_ЗУП,
	|	ВТ.КонечныйОстаток_ЗУП";

	Возврат Запрос.Выполнить().Выгрузить();

КонецФункции

// Объединение ЕРП + ЗУП по ДРФО, 4 разницы (ЕРП - ЗУП), фильтр по флагу.
Функция ОбъединитьТаблицы(ТаблицаЕРП, ТаблицаЗУП, МножествоДРФОЗУП, ТолькоСотрудникиИзЗУП)

	ТипЧисло = ОбщегоНазначения.ОписаниеТипаЧисло(15, 2);
	ТипСтрока = Новый ОписаниеТипов("Строка", , Новый КвалификаторыСтроки(150));

	Результат = Новый ТаблицаЗначений;
	Результат.Колонки.Добавить("ФизическоеЛицо", Новый ОписаниеТипов("СправочникСсылка.ФизическиеЛица"));
	Результат.Колонки.Добавить("ДРФО", ТипСтрока);
	Результат.Колонки.Добавить("ФЛ_НаименованиеЗУП", ТипСтрока);
	Результат.Колонки.Добавить("НачальныйОстаток_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("НачальныйОстаток_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("РазницаНачальныйОстаток", ТипЧисло);
	Результат.Колонки.Добавить("Начисления_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("Начисления_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("РазницаНачисления", ТипЧисло);
	Результат.Колонки.Добавить("Выплаты_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("Выплаты_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("РазницаВыплаты", ТипЧисло);
	Результат.Колонки.Добавить("КонечныйОстаток_ЕРП", ТипЧисло);
	Результат.Колонки.Добавить("КонечныйОстаток_ЗУП", ТипЧисло);
	Результат.Колонки.Добавить("РазницаКонечныйОстаток", ТипЧисло);

	Индекс = Новый Соответствие;

	Для Каждого СтрЕ Из ТаблицаЕРП Цикл
		НоваяСтрока = Результат.Добавить();
		НоваяСтрока.ФизическоеЛицо = СтрЕ.ФизическоеЛицо;
		НоваяСтрока.ДРФО = СтрЕ.ДРФО;
		НоваяСтрока.НачальныйОстаток_ЕРП = СтрЕ.НачальныйОстаток_ЕРП;
		НоваяСтрока.Начисления_ЕРП = СтрЕ.Начисления_ЕРП;
		НоваяСтрока.Выплаты_ЕРП = СтрЕ.Выплаты_ЕРП;
		НоваяСтрока.КонечныйОстаток_ЕРП = СтрЕ.КонечныйОстаток_ЕРП;
		Если НЕ ПустаяСтрока(СтрЕ.ДРФО) Тогда
			Индекс.Вставить(СтрЕ.ДРФО, Результат.Индекс(НоваяСтрока));
		КонецЕсли;
	КонецЦикла;

	Для Каждого СтрЗ Из ТаблицаЗУП Цикл
		Ключ = СокрЛП(СтрЗ.ДРФО);
		НомерСтроки = Индекс.Получить(Ключ);
		Если НомерСтроки <> Неопределено Тогда
			СтрокаРез = Результат.Получить(НомерСтроки);
		Иначе
			СтрокаРез = Результат.Добавить();
			СтрокаРез.ФизическоеЛицо = СтрЗ.ФизическоеЛицо;
			СтрокаРез.ДРФО = Ключ;
		КонецЕсли;
		СтрокаРез.ФЛ_НаименованиеЗУП = СтрЗ.ФЛ_Наименование;
		СтрокаРез.НачальныйОстаток_ЗУП = СтрЗ.НачальныйОстаток_ЗУП;
		СтрокаРез.Начисления_ЗУП = СтрЗ.Начисления_ЗУП;
		СтрокаРез.Выплаты_ЗУП = СтрЗ.Выплаты_ЗУП;
		СтрокаРез.КонечныйОстаток_ЗУП = СтрЗ.КонечныйОстаток_ЗУП;
	КонецЦикла;

	Для Каждого СтрР Из Результат Цикл
		СтрР.РазницаНачальныйОстаток = СтрР.НачальныйОстаток_ЕРП - СтрР.НачальныйОстаток_ЗУП;
		СтрР.РазницаНачисления = СтрР.Начисления_ЕРП - СтрР.Начисления_ЗУП;
		СтрР.РазницаВыплаты = СтрР.Выплаты_ЕРП - СтрР.Выплаты_ЗУП;
		СтрР.РазницаКонечныйОстаток = СтрР.КонечныйОстаток_ЕРП - СтрР.КонечныйОстаток_ЗУП;
	КонецЦикла;

	Если ТолькоСотрудникиИзЗУП Тогда
		Для Сч = 1 По Результат.Количество() Цикл
			НомерОбратный = Результат.Количество() - Сч;
			Если МножествоДРФОЗУП.Получить(Результат[НомерОбратный].ДРФО) = Неопределено Тогда
				Результат.Удалить(НомерОбратный);
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;

	Возврат Результат;

КонецФункции

#КонецОбласти

#КонецЕсли
```

**ВНИМАНИЕ к кавычкам в BSL-строках:** в строковых литералах BSL кавычка удваивается, поэтому 1С-условие `КодПоДРФО <> ""` внутри `"..."`-литерала модуля записывается как `<> """"""""` НЕВЕРНО — правильно `<> """"`. В коде выше при переносе в файл проверить: в литерале запроса пустая 1С-строка = 4 кавычки подряд (`""""`). Сверить глазами с эталоном `ПолучитьОстаткиBuhBud` (строка `И Организации.КодПоЕДРПОУ <> """"`).

- [ ] **Step 5.2: Review модуля** — прогнать `/1c-bsl-review` по файлу; запрещённые имена (`Ссылка`, `МетаДанные`, `Выполнить`) не используются; тексты запросов идентичны pretest.

- [ ] **Step 5.3: Commit**

```bash
git add "_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП/Ext/ObjectModule.bsl"
git commit -m "feat: ObjectModule отчёта сверки ВзСС ЕРП(Ф1+Ф2)<->zup_2 + флаг ТолькоСотрудникиИзЗУП"
```

---

### Task 6: Кодировка, валидация, сборка .erf

**Files:**
- Create: `_Rarzrabotki/Python/test/fix_encoding_sverka_zup.py`
- Output: `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.erf`

- [ ] **Step 6.1: Скрипт перекодировки BOM+CRLF**

```python
# -*- coding: utf-8 -*-
"""Перекодировка исходников отчёта в UTF-8 BOM + CRLF (требование 1С)."""
import io
import os

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Отчеты")
NAME = "А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП"
FILES = [
    NAME + ".xml",
    os.path.join(NAME, "Ext", "ObjectModule.bsl"),
    os.path.join(NAME, "Ext", "Help.xml"),
    os.path.join(NAME, "Ext", "Help", "ru.html"),
    os.path.join(NAME, "Ext", "Help", "uk.html"),
    os.path.join(NAME, "Templates", "ОсновнаяСхемаКомпоновкиДанных.xml"),
    os.path.join(NAME, "Templates", "ОсновнаяСхемаКомпоновкиДанных", "Ext", "Template.xml"),
]
for rel in FILES:
    p = os.path.normpath(os.path.join(BASE, rel))
    with io.open(p, "r", encoding="utf-8-sig") as fh:
        text = fh.read()
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    with io.open(p, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(text)
    print("OK", rel)
print("DONE")
```

Run: `python "_Rarzrabotki\Python\test\fix_encoding_sverka_zup.py"` → 7 × `OK`, `DONE`.

- [ ] **Step 6.2: erf-validate** — скил `erf-validate` по корневому XML отчёта. Ожидаемо: 0 ошибок (ложные WARN про External*-типы игнорировать — известный false positive).

- [ ] **Step 6.3: erf-build против реальной BaseERP** — скил `erf-build`: корневой XML → `.erf`, параметры базы `-InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" -UserName "Администратор" -Password "24043"` (стаб-база падает «Неизвестная версия формата 2.17»). Expected: файл `_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.erf` создан.

- [ ] **Step 6.4: Commit**

```bash
git add _Rarzrabotki/Python/test/fix_encoding_sverka_zup.py "_Rarzrabotki/Отчеты/А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.erf"
git commit -m "build: А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.erf (erf-validate 0 ошибок)"
```

---

### Task 7: Acceptance — COM-зеркало сверки и флага

**Files:**
- Create: `_Rarzrabotki/Python/test/test_sverka_vzss_zup_acceptance.py`

- [ ] **Step 7.1: Написать acceptance** (константа `MAPPING` = вердикт Task 2; ниже по умолчанию `"T1"` — заменить при вердикте T0):

```python
# -*- coding: utf-8 -*-
"""Acceptance: COM-зеркало отчёта А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.
Период: декабрь 2025. Проверки: согласованность Σ, поведение флага, дубликаты ДРФО."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MAPPING = "T1"  # вердикт test_sverka_vzss_zup_pretest2_signs.py
TOL = 0.01

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
zup = v8.Connect('Srvr="SQLSERVER";Ref="zup_2";Usr="cfo";Pwd="2442"')

def fetch(conn, text, cols):
    q = conn.NewObject("Запрос")
    q.Text = text
    t = q.Execute().Выгрузить()
    rows = []
    for i in range(t.Количество()):
        r = t.Получить(i)
        rows.append(tuple(getattr(r, c) for c in cols))
    return rows

erp_rows = {}
dup_erp = 0
for fl_drfo, n, p, r, k in fetch(erp, """ВЫБРАТЬ
	Ост.ФизическоеЛицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаВзаиморасчетовНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаВзаиморасчетовПриход) КАК Приход,
	СУММА(Ост.СуммаВзаиморасчетовРасход) КАК Расход,
	СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , ) КАК Ост
СГРУППИРОВАТЬ ПО Ост.ФизическоеЛицо.КодПоДРФО""", ("ДРФО", "НачОст", "Приход", "Расход", "КонОст")):
    key = (fl_drfo or "").strip()
    if key in erp_rows:
        dup_erp += 1
        old = erp_rows[key]
        erp_rows[key] = tuple(a + float(b or 0) for a, b in zip(old, (n, p, r, k)))
    else:
        erp_rows[key] = tuple(float(x or 0) for x in (n, p, r, k))

zup_rows = {}
for drfo, n, p, r, k in fetch(zup, """ВЫБРАТЬ
	Ост.Физлицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаУпрНачальныйОстаток) КАК НачОст,
	СУММА(Ост.СуммаУпрПриход) КАК Приход,
	СУММА(Ост.СуммаУпрРасход) КАК Расход,
	СУММА(Ост.СуммаУпрКонечныйОстаток) КАК КонОст
ИЗ
	РегистрНакопления.ВзаиморасчетыСРаботниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , Физлицо.КодПоДРФО <> "") КАК Ост
СГРУППИРОВАТЬ ПО Ост.Физлицо.КодПоДРФО""", ("ДРФО", "НачОст", "Приход", "Расход", "КонОст")):
    key = (drfo or "").strip()
    vals = tuple(float(x or 0) for x in (n, p, r, k))
    if MAPPING == "T1":
        vals = (-vals[0], vals[2], vals[1], -vals[3])
    if key in zup_rows:
        old = zup_rows[key]
        vals = tuple(a + b for a, b in zip(old, vals))
    zup_rows[key] = vals

membership = set()
for (drfo,) in fetch(zup, """ВЫБРАТЬ РАЗЛИЧНЫЕ
	ФЛ.КодПоДРФО КАК ДРФО
ИЗ Справочник.ФизическиеЛица КАК ФЛ
ГДЕ ФЛ.КодПоДРФО <> "" И НЕ ФЛ.ПометкаУдаления""", ("ДРФО",)):
    membership.add((drfo or "").strip())

all_keys = set(erp_rows) | set(zup_rows)
matched = set(erp_rows) & set(zup_rows)
flag_keys = {k for k in all_keys if k in membership}

Z4 = (0.0, 0.0, 0.0, 0.0)
sum_d = [0.0] * 4
worst = []
for k in all_keys:
    e = erp_rows.get(k, Z4)
    z = zup_rows.get(k, Z4)
    d = tuple(a - b for a, b in zip(e, z))
    for i in range(4):
        sum_d[i] += d[i]
    worst.append((abs(d[3]), k, d[3]))
worst.sort(reverse=True)

print(f"Строк: всего {len(all_keys)}, matched {len(matched)}, ЕРП-only {len(set(erp_rows) - set(zup_rows))}, ЗУП-only {len(set(zup_rows) - set(erp_rows))}")
print(f"Дубликаты ДРФО в ЕРП (слиты): {dup_erp}")
print(f"С флагом 'Только сотрудники из ЗУП': {len(flag_keys)} строк (скрыто {len(all_keys) - len(flag_keys)})")
print(f"Σ Разница: НачОст={sum_d[0]:,.2f} Начисл={sum_d[1]:,.2f} Выпл={sum_d[2]:,.2f} КонОст={sum_d[3]:,.2f}")

se = [sum(v[i] for v in erp_rows.values()) for i in range(4)]
sz = [sum(v[i] for v in zup_rows.values()) for i in range(4)]
ok_sum = all(abs(se[i] - sz[i] - sum_d[i]) <= TOL for i in range(4))
ok_flag = flag_keys <= membership and all(k in membership for k in flag_keys)
print(f"\nТоп-10 |Δ КонОст|:")
for absd, k, d in worst[:10]:
    print(f"  {k}: {d:,.2f}")

print(f"\nCHECK Σ(ЕРП)-Σ(ЗУП)==Σ(Разница): {'OK' if ok_sum else 'FAIL'}")
print(f"CHECK флаг подмножество членства: {'OK' if ok_flag else 'FAIL'}")
verdict = ok_sum and ok_flag
print("ACCEPTANCE: " + ("PASS" if verdict else "FAIL"))
sys.exit(0 if verdict else 1)
```

- [ ] **Step 7.2: Запустить**

Run: `python "_Rarzrabotki\Python\test\test_sverka_vzss_zup_acceptance.py"`
Expected: `ACCEPTANCE: PASS`, exit 0. Величины Σ Разница НЕ обязаны быть нулевыми (отчёт диагностический — Ф2-премии из РаспределениеКазна отсутствуют в zup_2 by-design); важна внутренняя согласованность и работа флага.

- [ ] **Step 7.3: Commit**

```bash
git add _Rarzrabotki/Python/test/test_sverka_vzss_zup_acceptance.py
git commit -m "test: acceptance COM-зеркало отчёта сверки ВзСС ЕРП<->ЗУП — PASS"
```

---

### Task 8: Rule #4 — копия в основную конфигурацию, финализация

- [ ] **Step 8.1: Копировать в основной каталог**

```powershell
$wt = "C:\Configuration_downloads\BASERP25\.claude\worktrees\happy-babbage-50f9cc"
$main = "C:\Configuration_downloads\BASERP25"
Copy-Item -Force "$wt\_Rarzrabotki\Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.xml" "$main\_Rarzrabotki\Отчеты\"
Copy-Item -Force "$wt\_Rarzrabotki\Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.erf" "$main\_Rarzrabotki\Отчеты\"
Copy-Item -Recurse -Force "$wt\_Rarzrabotki\Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП" "$main\_Rarzrabotki\Отчеты\"
Copy-Item -Force "$wt\_Rarzrabotki\Python\test\test_sverka_vzss_zup_pretest.py","$wt\_Rarzrabotki\Python\test\test_sverka_vzss_zup_pretest2_signs.py","$wt\_Rarzrabotki\Python\test\test_sverka_vzss_zup_acceptance.py","$wt\_Rarzrabotki\Python\test\fix_encoding_sverka_zup.py" "$main\_Rarzrabotki\Python\test\"
```

Expected: файлы существуют в `$main\_Rarzrabotki\Отчеты\` и `$main\_Rarzrabotki\Python\test\`.

- [ ] **Step 8.2: Финализация ветки** — скил `superpowers:finishing-a-development-branch`: merge `claude/happy-babbage-50f9cc` → `main`, push (по памяти — автономно). Напомнить пользователю: подключить `.erf` в «Дополнительные отчёты» и прогнать в UI на декабре 2025.

---

## Self-Review (выполнен)

- **Spec coverage:** ЕРП Ф1+Ф2 (Task 1/5), zup_2 регистр+справочник (Task 1/5), знаковый маппинг (Task 2/5), флаг с семантикой «членство в справочнике» (Task 4/5/7), СКД-структура и параметр (Task 4), Help (Task 3), erf против реальной базы + BOM/CRLF (Task 6), acceptance (Task 7), Rule #4 (Task 8). Пробелов нет.
- **Placeholders:** отсутствуют; единственная вилка (T0/T1) — оба варианта кода даны, выбор детерминирован выходом Task 2.
- **Type consistency:** имена колонок ТЗ (`ДРФО`, `ФЛ_Наименование`→`ФЛ_НаименованиеЗУП` при слиянии, `*_ЕРП`/`*_ЗУП`, `Разница*`) сквозные между Task 1, 4, 5, 7; имя набора `Сотрудники_Сравнение` совпадает в BSL и СКД.
