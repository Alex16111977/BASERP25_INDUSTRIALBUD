# План: печать «Аналіз для цеху» + корректная ЕкономіяСума (Документ.РасчетКомплектаций, BuhBud)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 10-я печатная форма документа РасчетКомплектаций для кладовщика (только количество, «Коментар цеху» под рукописные объяснения понаднормы) + формула `ЕкономіяСума` по плановым ценам.

**Architecture:** СКД-макет `МакетАнализЦех` (клон `МакетАнализССОдна`, построенный transform-скриптом), экспортная функция печати в ObjectModule, кнопка в попапе «Друк» формы. Формула экономии — точечная правка `ЗафиксироватьЭкономиюЭталона`. Проверки — Python COM smoke против живой базы BuhBud.

**Tech Stack:** BSL (BAS Бухгалтерія, дамп `_Rarzrabotki/BASEBuh`), СКД XML, Python + win32com (`V83.COMConnector`), скиллы db-load-xml / db-update / form-validate.

**Спека:** `docs/superpowers/specs/2026-08-06-raschet-komplektacij-pechat-ceh-ekonomia-design.md`

**Соединение с базой (все smoke):** `Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"` (fallback `Srvr="SQLSERVER"`).

**Пути.** Worktree: `C:\Configuration_downloads\BASERP25\.claude\worktrees\writeoff-report-by-department-85bf18`. Основная конфигурация: `C:\Configuration_downloads\BASERP25`. Ниже относительные пути — от корня repo; правки делаются в worktree, потом копируются (Task 6).

**Эталонные факты для проверок (диагностика 2026-08-06, док №3 = дом №1):**
- Док РасчетКомплектаций №00000000003, ТЧ 174 строки, 98 эталонов; итоги ТЧ: Остаток 7975.299, ВНорме 6183.118, Понад 1792.181, Економія 1604.067, Норма 7787.185.
- Группа «Бойлер електричний»: Норма 2 (СС: «Бойлер на 30 л» + «Бойлер на 82 л»), Видано 1, В нормі 1, Економія 1; карточка «Електричний водонагрівач O`Pro Slim PC 30, 30л».
- Нормы групп: Вимикач 10, Вітробарєр 42, Хомут 14, Герметик 10.
- Понаднормовые: АВР (понад 6), Гайка (понад 5).
- Новая сумма экономии по дому №1 (пропорциональная): ~34 691.33 грн (справочно; smoke формулы проверяет инварианты, не абсолют).

---

### Task 1: Builder СКД-макета `МакетАнализЦех`

**Files:**
- Create: `_Rarzrabotki/Python/test/doc_kompl_build_maket_ceh.py`
- Create (генерируются скриптом): `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Templates/МакетАнализЦех/Ext/Template.xml`, `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Templates/МакетАнализЦех.xml`
- Modify (скриптом): `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml` (регистрация `<Template>`)

- [ ] **Step 1.1: Написать builder-скрипт**

Скрипт трансформирует текст `МакетАнализССОдна/Ext/Template.xml` (кодировка utf-8-sig, сохранять с BOM как исходник). UUID обёртки — фиксированный `64dd02bb-7ffb-4d3d-b256-4a4b0eb21246` (стабилен при перезапусках — грабля «стабильный uuid»).

```python
# -*- coding: utf-8 -*-
"""Строит Templates/МакетАнализЦех из МакетАнализССОдна (печать для кладовщика).

Колонки: Норма СС | Видано | Одиниця | В нормі | Понад норму | Економія | Коментар цеху.
Норма/Економія — только на строке группы. Жирные группы, жёлтым только понаднорма.
Идемпотентен: перезапись целевых файлов, регистрация в корневом xml — однократная.
"""
import re
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BASE = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\writeoff-report-by-department-85bf18\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций"
SRC_TPL = BASE + r"\Templates\МакетАнализССОдна\Ext\Template.xml"
SRC_WRAP = BASE + r"\Templates\МакетАнализССОдна.xml"
DST_DIR = BASE + r"\Templates\МакетАнализЦех"
DST_TPL = DST_DIR + r"\Ext\Template.xml"
DST_WRAP = BASE + r"\Templates\МакетАнализЦех.xml"
ROOT_XML = BASE + r"\..\РасчетКомплектаций.xml"
UUID_NEW = "64dd02bb-7ffb-4d3d-b256-4a4b0eb21246"

t = open(SRC_TPL, encoding="utf-8-sig").read()

def cut_block(text, open_tag, close_tag, marker):
    """Вырезает первый блок open_tag...close_tag, содержащий marker."""
    start = 0
    while True:
        i = text.find(open_tag, start)
        assert i >= 0, f"нет блока с {marker!r}"
        j = text.find(close_tag, i) + len(close_tag)
        if marker in text[i:j]:
            # съесть перевод строки с отступом перед блоком
            k = text.rfind("\n", 0, i)
            return text[:k] + text[j:], text[i:j]
        start = j

def get_block(text, open_tag, close_tag, marker):
    start = 0
    while True:
        i = text.find(open_tag, start)
        assert i >= 0, f"нет блока с {marker!r}"
        j = text.find(close_tag, i) + len(close_tag)
        if marker in text[i:j]:
            return text[i:j]
        start = j

# --- 1. Удалить ненужные поля набора данных и вычисляемые поля ---
for fld in ("НормаНоменкл", "ВНормеСумма", "ПонадНормуСумма", "ЕкономіяСума",
            "НормаСуммаНоменкл", "СуммаОстатка"):
    t, _ = cut_block(t, "<field xsi:type=", "</field>", f"<dataPath>{fld}</dataPath>")
t, _ = cut_block(t, "<calculatedField>", "</calculatedField>", "<dataPath>РизницяСС</dataPath>")
t, _ = cut_block(t, "<calculatedField>", "</calculatedField>", "<dataPath>ПроцВиконанняСС</dataPath>")

# --- 2. Добавить поле Норма (клон блока ВНорме) с титулом «Норма СС» ---
vnorme = get_block(t, "<field xsi:type=", "</field>", "<dataPath>ВНорме</dataPath>")
norma = vnorme.replace("<dataPath>ВНорме</dataPath>", "<dataPath>Норма</dataPath>")
norma = norma.replace("<field>ВНорме</field>", "<field>Норма</field>")
norma = re.sub(r"<v8:content>[^<]*</v8:content>", "<v8:content>Норма СС</v8:content>", norma)
i = t.find("<field xsi:type=")  # вставить перед первым полем
k = t.rfind("\n", 0, i)
indent = t[k + 1:i]
t = t[:i] + norma + "\n" + indent + t[i:]

# --- 3. Добавить вычисляемое поле КоментарЦеху (пустая строка) ---
comment_field = (
    "<calculatedField>\n"
    "\t\t<dataPath>КоментарЦеху</dataPath>\n"
    "\t\t<expression>\"\"</expression>\n"
    "\t\t<title xsi:type=\"v8:LocalStringType\">\n"
    "\t\t\t<v8:item>\n"
    "\t\t\t\t<v8:lang>uk</v8:lang>\n"
    "\t\t\t\t<v8:content>Коментар цеху</v8:content>\n"
    "\t\t\t</v8:item>\n"
    "\t\t</title>\n"
    "\t</calculatedField>")
anchor = "<totalField>"
i = t.find(anchor)
assert i > 0
k = t.rfind("\n", 0, i)
indent = t[k + 1:i]
t = t[:i] + comment_field + "\n" + indent + t[i:]

# --- 4. totalFields: удалить лишние, добавить Норма ---
for fld in ("НормаНоменкл", "ВНормеСумма", "ПонадНормуСумма", "ЕкономіяСума",
            "НормаСуммаНоменкл", "СуммаОстатка"):
    t, _ = cut_block(t, "<totalField>", "</totalField>", f"<dataPath>{fld}</dataPath>")
tf_vnorme = get_block(t, "<totalField>", "</totalField>", "<dataPath>ВНорме</dataPath>")
tf_norma = tf_vnorme.replace("ВНорме", "Норма")
i = t.find("<totalField>")
k = t.rfind("\n", 0, i)
indent = t[k + 1:i]
t = t[:i] + tf_norma + "\n" + indent + t[i:]

# --- 5. Глобальный selection: 7 полей в порядке колонок ---
def selection_xml(fields, indent):
    items = "".join(
        f"{indent}\t<dcsset:item xsi:type=\"dcsset:SelectedItemField\">\n"
        f"{indent}\t\t<dcsset:field>{f}</dcsset:field>\n"
        f"{indent}\t</dcsset:item>\n" for f in fields)
    return f"<dcsset:selection>\n{items}{indent}</dcsset:selection>"

ALL7 = ["Норма", "Остаток", "ЕдиницаЕдина", "ВНорме", "ПонадНорму", "Экономия", "КоментарЦеху"]
DET5 = ["Остаток", "ЕдиницаЕдина", "ВНорме", "ПонадНорму", "КоментарЦеху"]

i = t.find("<dcsset:selection>")
j = t.find("</dcsset:selection>", i) + len("</dcsset:selection>")
t = t[:i] + selection_xml(ALL7, "\t\t\t") + t[j:]

# --- 6. Титул Остаток -> Видано ---
ost = get_block(t, "<field xsi:type=", "</field>", "<dataPath>Остаток</dataPath>")
t = t.replace(ost, re.sub(r"<v8:content>[^<]*</v8:content>",
                          "<v8:content>Видано</v8:content>", ost), 1)

# --- 7. Условное оформление: удалить зелёное, янтарное перевести на ПонадНорму ---
t, _ = cut_block(t, "<dcsset:item>", "</dcsset:item>", "#E2EFDA")
t = t.replace("<dcsset:left xsi:type=\"dcscor:Field\">ПонадНормуСумма</dcsset:left>",
              "<dcsset:left xsi:type=\"dcscor:Field\">ПонадНорму</dcsset:left>")

# --- 8. Заголовок и представление варианта ---
t = t.replace("<v8:content>Аналіз СС (одна одиниця)</v8:content>",
              "<v8:content>Аналіз залишків для списання за СС</v8:content>")

# --- 9. Selection уровней структуры (2 Auto: группа и детали) + жирность группы ---
auto_sel = ("<dcsset:selection>\n\t\t\t\t\t<dcsset:item xsi:type=\"dcsset:SelectedItemAuto\"/>"
            "\n\t\t\t\t</dcsset:selection>")
assert t.count(auto_sel) == 1, f"ожидался 1 Auto-selection группы, найдено {t.count(auto_sel)}"
bold_ca = (
    "<dcsset:conditionalAppearance>\n"
    "\t\t\t\t\t<dcsset:item>\n"
    "\t\t\t\t\t\t<dcsset:selection/>\n"
    "\t\t\t\t\t\t<dcsset:appearance>\n"
    "\t\t\t\t\t\t\t<dcscor:item xsi:type=\"dcsset:SettingsParameterValue\">\n"
    "\t\t\t\t\t\t\t\t<dcscor:parameter>Шрифт</dcscor:parameter>\n"
    "\t\t\t\t\t\t\t\t<dcscor:value xsi:type=\"v8ui:Font\" ref=\"sys:DefaultGUIFont\" bold=\"true\" italic=\"false\" underline=\"false\" strikeout=\"false\" kind=\"WindowsFont\"/>\n"
    "\t\t\t\t\t\t\t</dcscor:item>\n"
    "\t\t\t\t\t\t</dcsset:appearance>\n"
    "\t\t\t\t\t\t<dcsset:useInFieldsHeader>DontUse</dcsset:useInFieldsHeader>\n"
    "\t\t\t\t\t\t<dcsset:useInHeader>DontUse</dcsset:useInHeader>\n"
    "\t\t\t\t\t\t<dcsset:useInParameters>DontUse</dcsset:useInParameters>\n"
    "\t\t\t\t\t</dcsset:item>\n"
    "\t\t\t\t</dcsset:conditionalAppearance>")
t = t.replace(auto_sel, selection_xml(ALL7, "\t\t\t\t") + "\n\t\t\t\t" + bold_ca, 1)

auto_sel_det = ("<dcsset:selection>\n\t\t\t\t\t\t<dcsset:item xsi:type=\"dcsset:SelectedItemAuto\"/>"
                "\n\t\t\t\t\t</dcsset:selection>")
assert t.count(auto_sel_det) == 1, f"ожидался 1 Auto-selection деталей, найдено {t.count(auto_sel_det)}"
t = t.replace(auto_sel_det, selection_xml(DET5, "\t\t\t\t\t"), 1)

# --- 10. Записать Template.xml ---
import os
os.makedirs(DST_DIR + r"\Ext", exist_ok=True)
with open(DST_TPL, "w", encoding="utf-8-sig", newline="\n") as f:
    f.write(t)

# --- 11. Обёртка с фиксированным uuid ---
w = open(SRC_WRAP, encoding="utf-8-sig").read()
w = re.sub(r'uuid="[0-9a-f-]+"', f'uuid="{UUID_NEW}"', w)
w = w.replace("<Name>МакетАнализССОдна</Name>", "<Name>МакетАнализЦех</Name>")
w = w.replace("<v8:content>Макет аналіз СС (одна одиниця)</v8:content>",
              "<v8:content>Макет аналіз для цеху</v8:content>")
with open(DST_WRAP, "w", encoding="utf-8-sig", newline="\n") as f:
    f.write(w)

# --- 12. Регистрация в корневом xml документа ---
root_path = os.path.normpath(ROOT_XML)
r = open(root_path, encoding="utf-8-sig").read()
if "<Template>МакетАнализЦех</Template>" not in r:
    r = r.replace("<Template>МакетАнализССОдна</Template>",
                  "<Template>МакетАнализССОдна</Template>\n\t\t\t<Template>МакетАнализЦех</Template>")
    with open(root_path, "w", encoding="utf-8-sig", newline="\n") as f:
        f.write(r)
    print("Корневой xml: макет зарегистрирован")
else:
    print("Корневой xml: регистрация уже есть")

# --- 13. Самопроверки ---
import xml.etree.ElementTree as ET
ET.fromstring(open(DST_TPL, encoding="utf-8-sig").read().encode("utf-8"))
chk = open(DST_TPL, encoding="utf-8-sig").read()
for absent in ("НормаНоменкл", "ВНормеСумма", "ПонадНормуСумма", "ЕкономіяСума",
               "НормаСуммаНоменкл", "СуммаОстатка", "РизницяСС", "ПроцВиконанняСС", "#E2EFDA"):
    assert absent not in chk, f"осталось: {absent}"
for present in ("<dataPath>Норма</dataPath>", "КоментарЦеху", "Видано", "Норма СС",
                "#FFE699", 'bold="true"',
                "<dcsset:left xsi:type=\"dcscor:Field\">ПонадНорму</dcsset:left>"):
    assert present in chk, f"нет: {present}"
assert chk.count("SelectedItemAuto") == 0, "остались Auto-selection"
print("OK: МакетАнализЦех построен и проверен")
```

- [ ] **Step 1.2: Запустить builder**

Run: `python -X utf8 _Rarzrabotki/Python/test/doc_kompl_build_maket_ceh.py` (из корня worktree)
Expected: `Корневой xml: макет зарегистрирован` + `OK: МакетАнализЦех построен и проверен`. При assert-падении — читать сообщение, поправить якорь в скрипте (отступы Auto-selection сверить с фактическим файлом), перезапустить.

- [ ] **Step 1.3: Контроль diff корневого xml**

Run: `git diff --stat _Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml`
Expected: 1 строка добавлена (`<Template>МакетАнализЦех</Template>`), больше НИЧЕГО (защита от затирания живых реквизитов).

- [ ] **Step 1.4: Commit**

```bash
git add _Rarzrabotki/Python/test/doc_kompl_build_maket_ceh.py "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Templates/МакетАнализЦех.xml" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Templates/МакетАнализЦех/"
git commit -m "feat(raschet-kompl): СКД-макет МакетАнализЦех (печать для кладовщика)"
```

---

### Task 2: ObjectModule — функция печати `СформироватьПечатьАнализЦех`

**Files:**
- Modify: `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Ext/ObjectModule.bsl` (вставка после `КонецФункции` функции `СформироватьПечатьАнализССОдна`, ~стр. 518)

- [ ] **Step 2.1: Вставить функцию**

Найти конец функции `СформироватьПечатьАнализССОдна` (строки 497–518: `Возврат ТабДок;` → `КонецФункции`) и вставить ПОСЛЕ него:

```bsl

// Друкована форма «для цеху» (кількісна): видається комірнику; по жовтих (понаднормових)
// рядках він пише пояснення від руки в колонку «Коментар цеху». Сум немає — тільки кількість.
// «Норма СС» і «Економія» — лише на рядку групи (загальна назва), на картках порожньо.
Функция СформироватьПечатьАнализЦех(ДанныеАнализа) Экспорт

	ТЗ = ДанныеВТаблицуПланФакт(ДанныеАнализа);
	Схема = ПолучитьМакет("МакетАнализЦех");

	КомпоновщикМакета = Новый КомпоновщикМакетаКомпоновкиДанных;
	Настройки = Схема.НастройкиПоУмолчанию;
	УстановитьЗаголовокЗвіту(Настройки, "Аналіз залишків для списання за СС");
	Макет = КомпоновщикМакета.Выполнить(Схема, Настройки, , , Тип("ГенераторМакетаКомпоновкиДанных"));

	ВнешниеНаборы = Новый Структура("Данные", ТЗ);
	Процессор = Новый ПроцессорКомпоновкиДанных;
	Процессор.Инициализировать(Макет, ВнешниеНаборы, , Истина);

	ТабДок = Новый ТабличныйДокумент;
	ПроцессорВывода = Новый ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент;
	ПроцессорВывода.УстановитьДокумент(ТабДок);
	ПроцессорВывода.Вывести(Процессор);

	// Широка колонка «Коментар цеху» під рукопис (ШиринаКолонки — на НЕоб'єднаній клітинці даних).
	Для НомерСтроки = 1 По Мин(ТабДок.ВысотаТаблицы, 30) Цикл
		Для НомерКолонки = 1 По ТабДок.ШиринаТаблицы Цикл
			Если СокрЛП(ТабДок.Область(НомерСтроки, НомерКолонки).Текст) = "Коментар цеху" Тогда
				ТабДок.Область(НомерСтроки + 1, НомерКолонки).ШиринаКолонки = 35;
				Возврат ТабДок;
			КонецЕсли;
		КонецЦикла;
	КонецЦикла;

	Возврат ТабДок;

КонецФункции
```

- [ ] **Step 2.2: Проверка парности операторных скобок**

Run: `grep -c "^Функция\|^КонецФункции" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Ext/ObjectModule.bsl"` — оба счётчика выросли на 1 против исходных; плюс grep-чек опечаток: `grep -n "КонецФункция\|КонецПроцедуры$" file | head` — «КонецФункция» быть НЕ должно.

- [ ] **Step 2.3: Commit**

```bash
git add "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Ext/ObjectModule.bsl"
git commit -m "feat(raschet-kompl): СформироватьПечатьАнализЦех — печать для кладовщика"
```

---

### Task 3: ObjectModule — формула `ЕкономіяСума` по плановым ценам

**Files:**
- Modify: `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Ext/ObjectModule.bsl` (`ЗаполнитьАнализ` ~стр. 953–1052, `ЗафиксироватьЭкономиюЭталона` ~стр. 1054–1061; номера строк сдвинулись на +33 после Task 2 — искать по тексту)

- [ ] **Step 3.1: Инициализация НормаЭталона + удаление СуммаОстаткаЭталона (3 правки в `ЗаполнитьАнализ`)**

Правка А — блок инициализации перед циклом. Было:
```bsl
	ТекЭталон = Неопределено;
	ОстатокНормы = 0;
	НормаСумЭталона = 0;
	СуммаОстаткаЭталона = 0;
```
Стало:
```bsl
	ТекЭталон = Неопределено;
	ОстатокНормы = 0;
	НормаЭталона = 0;
	НормаСумЭталона = 0;
```

Правка Б — оба вызова фиксации (при смене эталона и после цикла). Было (2 вхождения):
```bsl
		ЗафиксироватьЭкономиюЭталона(ПерваяСтрокаЭталона, ОстатокНормы, НормаСумЭталона, СуммаОстаткаЭталона);
```
и
```bsl
	ЗафиксироватьЭкономиюЭталона(ПерваяСтрокаЭталона, ОстатокНормы, НормаСумЭталона, СуммаОстаткаЭталона);
```
Стало (соответственно, с сохранением отступов):
```bsl
		ЗафиксироватьЭкономиюЭталона(ПерваяСтрокаЭталона, ОстатокНормы, НормаЭталона, НормаСумЭталона);
```
и
```bsl
	ЗафиксироватьЭкономиюЭталона(ПерваяСтрокаЭталона, ОстатокНормы, НормаЭталона, НормаСумЭталона);
```
(Вызов при смене эталона исполняется ДО переприсвоения `НормаЭталона` — передаёт значения предыдущего эталона, это корректно.)

Правка В — убрать накопление суммы остатка эталона. Было:
```bsl
			ОстатокНормы = НормаЭталона;
			СуммаОстаткаЭталона = 0;
			ПервыйРяд = Ложь;
```
Стало:
```bsl
			ОстатокНормы = НормаЭталона;
			ПервыйРяд = Ложь;
```
И было:
```bsl
		ПонадНормуСумма = Стр.СуммаОстатка - ВНормеСумма;
		СуммаОстаткаЭталона = СуммаОстаткаЭталона + Стр.СуммаОстатка;
```
Стало:
```bsl
		ПонадНормуСумма = Стр.СуммаОстатка - ВНормеСумма;
```

- [ ] **Step 3.2: Новое тело `ЗафиксироватьЭкономиюЭталона`**

Было:
```bsl
// Пишет неиспользованную норму (кол + сумма) на представительную строку эталона.
Процедура ЗафиксироватьЭкономиюЭталона(СтрокаЭталона, ОстатокНормы, НормаСумЭталона, СуммаОстаткаЭталона)
	Если СтрокаЭталона <> Неопределено Тогда
		СтрокаЭталона.Экономия = ?(ОстатокНормы > 0, ОстатокНормы, 0);
		РазницаСум = НормаСумЭталона - СуммаОстаткаЭталона;
		СтрокаЭталона.ЕкономіяСума = ?(РазницаСум > 0, РазницаСум, 0);
	КонецЕсли;
КонецПроцедуры
```
Стало:
```bsl
// Пишет неиспользованную норму (кол + сумма) на представительную строку эталона.
// Сумма — КІЛЬКІСНА оцінка за плановими цінами СС: Економія × (НормаСум/Норма).
// Цінову різницю план/факт в економію НЕ підмішуємо (рішення бухгалтера 2026-08-06):
// раніше «Гайка» показувала 11 906 грн економії при перевитраті 5 шт.
Процедура ЗафиксироватьЭкономиюЭталона(СтрокаЭталона, ОстатокНормы, НормаЭталона, НормаСумЭталона)
	Если СтрокаЭталона <> Неопределено Тогда
		СтрокаЭталона.Экономия = ?(ОстатокНормы > 0, ОстатокНормы, 0);
		Если ОстатокНормы > 0 И НормаЭталона > 0 Тогда
			СтрокаЭталона.ЕкономіяСума = Окр(ОстатокНормы * НормаСумЭталона / НормаЭталона, 2);
		Иначе
			СтрокаЭталона.ЕкономіяСума = 0;
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры
```

- [ ] **Step 3.3: Контроль**

Run: `grep -n "СуммаОстаткаЭталона" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Ext/ObjectModule.bsl"`
Expected: пусто (переменная удалена всюду). `grep -c "ЗафиксироватьЭкономиюЭталона"` = 3 (объявление + 2 вызова).

- [ ] **Step 3.4: Commit**

```bash
git add "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Ext/ObjectModule.bsl"
git commit -m "fix(raschet-kompl): ЕкономіяСума = кількість × нормова ціна (без цінової домішки)"
```

---

### Task 4: Форма — команда и кнопка «Аналіз для цеху (з коментарем)»

**Files:**
- Modify: `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form.xml` (кнопка после `КнопкаПечатьМ29`, команда после `ПечатьМ29`; занятые id ≤ 573 → новые 574/575/576)
- Modify: `_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form/Module.bsl` (обработчики после `ПечатьАнализССОднаНаСервере`)

- [ ] **Step 4.1: Form.xml — кнопка в попапе «Друк»**

После закрывающего `</Button>` кнопки `КнопкаПечатьМ29` (якорь — строка `<ExtendedTooltip name="КнопкаПечатьМ29РасширеннаяПодсказка" id="572"/>` + следующий `</Button>`) вставить:

```xml
					<Button name="КнопкаПечатьАнализЦех" id="574">
						<Type>UsualButton</Type>
						<CommandName>Form.Command.ПечатьАнализЦех</CommandName>
						<Title>
							<v8:item>
								<v8:lang>uk</v8:lang>
								<v8:content>Аналіз для цеху (з коментарем)</v8:content>
							</v8:item>
						</Title>
						<ExtendedTooltip name="КнопкаПечатьАнализЦехРасширеннаяПодсказка" id="575"/>
					</Button>
```

- [ ] **Step 4.2: Form.xml — команда**

После закрывающего `</Command>` команды `ПечатьМ29` (id="573") вставить:

```xml
		<Command name="ПечатьАнализЦех" id="576">
			<Title>
				<v8:item>
					<v8:lang>uk</v8:lang>
					<v8:content>Аналіз для цеху (з коментарем)</v8:content>
				</v8:item>
			</Title>
			<Action>ПечатьАнализЦех</Action>
		</Command>
```

- [ ] **Step 4.3: Module.bsl — обработчики**

После `КонецФункции` функции `ПечатьАнализССОднаНаСервере` (~стр. 212) вставить:

```bsl

&НаКлиенте
Процедура ПечатьАнализЦех(Команда)
	ТабДок = ПечатьАнализЦехНаСервере();
	Если ТабДок <> Неопределено Тогда
		ТабДок.Показать("Аналіз залишків для списання за СС");
	КонецЕсли;
КонецПроцедуры

&НаСервере
Функция ПечатьАнализЦехНаСервере()
	Если ТаблицаАнализа.Количество() = 0 Тогда
		Сообщить("Спочатку натисніть «Розрахувати».");
		Возврат Неопределено;
	КонецЕсли;
	ОбъектОбработки = РеквизитФормыВЗначение("Объект");
	Возврат ОбъектОбработки.СформироватьПечатьАнализЦех(ТаблицаАнализа);
КонецФункции
```

- [ ] **Step 4.4: form-validate**

Вызвать скилл `form-validate` для `Forms/ФормаДокумента`. Expected: без НОВЫХ ошибок (известное ложное срабатывание «~Список.X not found» игнорировать — memory `form_validate_tilde_spisok_false_positive`). Проверить уникальность id: `grep -o 'id="57[0-9]"' Form.xml | sort | uniq -d` → пусто.

- [ ] **Step 4.5: Commit**

```bash
git add "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form.xml" "_Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form/Module.bsl"
git commit -m "feat(raschet-kompl): кнопка «Аналіз для цеху (з коментарем)» у попапі Друк"
```

---

### Task 5: Ревью кода BSL

- [ ] **Step 5.1:** Вызвать скилл `1c-bsl-review` для изменённого `ObjectModule.bsl` (блоки: новая функция печати, ЗаполнитьАнализ, ЗафиксироватьЭкономиюЭталона) и `Module.bsl` формы. Исправить замечания уровня «критично»; замечания уровня «стиль» — на усмотрение (не переписывать чужой стиль модуля).
- [ ] **Step 5.2:** Если были правки — commit `refactor(raschet-kompl): зауваження bsl-review`.

---

### Task 6: Копирование в основную конфигурацию и загрузка в базу

**Files (worktree → main):** те же 6 файлов.

- [ ] **Step 6.1: Проверить, что Конфигуратор не держит базу**

Run: `tasklist | grep -i 1cv8` (PowerShell: `Get-Process 1cv8* -ErrorAction SilentlyContinue`). Если есть процессы Конфигуратора у пользователя — ОСТАНОВИТЬСЯ и сообщить пользователю (открытый Конфигуратор со старой копией откатывает загрузку — случалось дважды). Клиентские сеансы Enterprise допустимы (`db-update -Dynamic+`).

- [ ] **Step 6.2: Скопировать файлы (Rule #4)**

```bash
W="C:/Configuration_downloads/BASERP25/.claude/worktrees/writeoff-report-by-department-85bf18/_Rarzrabotki/BASEBuh/Documents"
M="C:/Configuration_downloads/BASERP25/_Rarzrabotki/BASEBuh/Documents"
cp "$W/РасчетКомплектаций.xml" "$M/РасчетКомплектаций.xml"
cp "$W/РасчетКомплектаций/Ext/ObjectModule.bsl" "$M/РасчетКомплектаций/Ext/ObjectModule.bsl"
mkdir -p "$M/РасчетКомплектаций/Templates/МакетАнализЦех/Ext"
cp "$W/РасчетКомплектаций/Templates/МакетАнализЦех.xml" "$M/РасчетКомплектаций/Templates/МакетАнализЦех.xml"
cp "$W/РасчетКомплектаций/Templates/МакетАнализЦех/Ext/Template.xml" "$M/РасчетКомплектаций/Templates/МакетАнализЦех/Ext/Template.xml"
cp "$W/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form.xml" "$M/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form.xml"
cp "$W/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form/Module.bsl" "$M/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form/Module.bsl"
```
Перед копированием корневого xml: `git -C C:/Configuration_downloads/BASERP25 diff --stat _Rarzrabotki/BASEBuh/Documents/РасчетКомплектаций.xml` в ОСНОВНОЙ папке — если там локальные незакоммиченные изменения объекта (чужие правки), СТОП и разобраться.

- [ ] **Step 6.3: db-load-xml Partial**

Вызвать скилл `db-load-xml` против `bas_industrialbud` (cfo/2442) со списком файлов (относительно каталога дампа `_Rarzrabotki/BASEBuh`):
```
Documents/РасчетКомплектаций.xml
Documents/РасчетКомплектаций/Ext/ObjectModule.bsl
Documents/РасчетКомплектаций/Templates/МакетАнализЦех.xml
Documents/РасчетКомплектаций/Templates/МакетАнализЦех/Ext/Template.xml
Documents/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form.xml
Documents/РасчетКомплектаций/Forms/ФормаДокумента/Ext/Form/Module.bsl
```
Родительский `Documents/РасчетКомплектаций.xml` обязателен — новый макет (иначе «нельзя добавлять объекты метаданных без загрузки родительского объекта»). Expected: exit 0.

- [ ] **Step 6.4: db-update**

Вызвать скилл `db-update` с `-Dynamic+` (добавление объектов проходит без монопольного режима). Expected: exit 0.

---

### Task 7: Smoke новой печати + кросс-сверка с отчётом

**Files:**
- Create: `_Rarzrabotki/Python/test/smoke_doc_kompl_pechat_ceh.py`

- [ ] **Step 7.1: Написать smoke**

```python
# -*- coding: utf-8 -*-
"""Smoke печати «Аналіз для цеху» (МакетАнализЦех) на живом доке №3 (дом №1).

Проверки: колонки, Бойлер (норма/економія только на группе, жирный), заливка только
понаднормовых, Разом = сохранённая ТЧ, заголовок, кросс-сверка норм с отчётом
А_ОтчетПоСписаниюНаПроизводствоБухгалтерский (ПланКол по исполнителю №1).
"""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ОтчетПоСписаниюНаПроизводствоБухгалтерский.erf"

v8 = win32com.client.Dispatch("V83.COMConnector")
try:
    erp = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
except Exception:
    erp = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = erp.String

FAILS = []
def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f" | {detail}" if detail else ""))
    if not cond:
        FAILS.append(name)

def num(txt):
    txt = (txt or "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None

q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д
ГДЕ Д.Номер = "00000000003" И НЕ Д.ПометкаУдаления"""
doc = q.Execute().Выгрузить().Получить(0).Ссылка.ПолучитьОбъект()
data = doc.ТабличнаяЧастьОстатков.Выгрузить()

td = doc.СформироватьПечатьАнализЦех(data)
check("печать сформирована", td is not None)
H, W = td.ВысотаТаблицы, td.ШиринаТаблицы
print(f"ТабДок {H}x{W}")

def cell(r, c):
    return (td.Область(r, c, r, c).Текст or "").strip()

# 1. Строка заголовков колонок и карта колонок
hdr_row, cols = None, {}
for r in range(1, min(H, 15) + 1):
    row_texts = {c: cell(r, c) for c in range(1, W + 1)}
    if "Видано" in row_texts.values():
        hdr_row = r
        for c, t in row_texts.items():
            if t:
                cols[t] = c
        break
check("строка заголовков найдена", hdr_row is not None, str(cols))
for t in ("Норма СС", "Видано", "Одиниця виміру", "В нормі", "Понад норму",
          "Економія", "Коментар цеху"):
    check(f"колонка {t!r}", t in cols)
for t in ("Згідно з СС", "В нормі, грн", "Факт. списання"):
    check(f"нет колонки {t!r}", t not in cols)

# 2. Заголовок печати
head = " ".join(cell(r, c) for r in range(1, hdr_row) for c in range(1, W + 1) if cell(r, c))
check("заголовок: назва", "Аналіз залишків для списання за СС" in head)
check("заголовок: підрозділ", "МД IRS 2026" in head)
check("заголовок: склад №1", "15 м №1" in head)

def find_row(text, after=0):
    for r in range(max(hdr_row + 1, after + 1), H + 1):
        for c in range(1, min(W, 3) + 1):
            if cell(r, c) == text:
                return r
    return None

def bold(r):
    for c in range(1, min(W, 3) + 1):
        if cell(r, c):
            return bool(td.Область(r, c, r, c).Шрифт.Жирный)
    return False

def fill_rgb(r):
    a = td.Область(r, cols["Видано"], r, cols["Видано"])
    цв = a.ЦветФона
    return (цв.R, цв.G, цв.B)

# 3. Бойлер: группа vs карточка
rg = find_row("Бойлер електричний")
check("группа Бойлер найдена", rg is not None)
if rg:
    check("Бойлер: Норма СС=2", num(cell(rg, cols["Норма СС"])) == 2)
    check("Бойлер: Видано=1", num(cell(rg, cols["Видано"])) == 1)
    check("Бойлер: В нормі=1", num(cell(rg, cols["В нормі"])) == 1)
    check("Бойлер: Економія=1", num(cell(rg, cols["Економія"])) == 1)
    check("Бойлер: группа жирная", bold(rg))
    rc = find_row("Електричний водонагрівач O`Pro Slim PC 30, 30л", rg)
    check("карточка водонагрівача найдена", rc is not None)
    if rc:
        check("карточка: Економія пусто", cell(rc, cols["Економія"]) == "")
        check("карточка: Норма СС пусто", cell(rc, cols["Норма СС"]) == "")
        check("карточка: не жирная", not bold(rc))
        check("карточка: Коментар цеху пуст", cell(rc, cols["Коментар цеху"]) == "")

# 4. Нормы других групп (для кросс-сверки)
group_norms = {}
for name in ("Вимикач", "Вітробарєр", "Хомут"):
    r = find_row(name)
    if r:
        group_norms[name] = num(cell(r, cols["Норма СС"]))
check("Вимикач: Норма СС=10", group_norms.get("Вимикач") == 10)
check("Вітробарєр: Норма СС=42", group_norms.get("Вітробарєр") == 42)
check("Хомут: Норма СС=14", group_norms.get("Хомут") == 14)

# 5. Заливка: понаднормовая жёлтая, нормовая — нет
ra = find_row("АВР")
if ra:
    check("АВР (понад): жёлтая заливка", fill_rgb(ra) == (255, 230, 153), str(fill_rgb(ra)))
rb = find_row("Брус")
if rb:
    check("Брус (в нормі): НЕ жёлтая", fill_rgb(rb) != (255, 230, 153), str(fill_rgb(rb)))

# 6. Разом = сохранённая ТЧ
tot = {"ost": 0.0, "vn": 0.0, "pn": 0.0, "ek": 0.0}
tch = doc.ТабличнаяЧастьОстатков
for i in range(tch.Количество()):
    s = tch.Получить(i)
    tot["ost"] += s.Остаток; tot["vn"] += s.ВНорме
    tot["pn"] += s.ПонадНорму; tot["ek"] += s.Экономия
rr = None
for r in range(H, hdr_row, -1):
    if any("Разом" in cell(r, c) for c in range(1, min(W, 3) + 1)):
        rr = r
        break
check("строка Разом найдена", rr is not None)
if rr:
    for title, key in (("Видано", "ost"), ("В нормі", "vn"),
                       ("Понад норму", "pn"), ("Економія", "ek")):
        v = num(cell(rr, cols[title]))
        check(f"Разом {title} = ТЧ", v is not None and abs(v - tot[key]) < 0.01,
              f"{v} vs {tot[key]:.3f}")

# 7. Кросс-сверка норм с отчётом (План, кіл по исполнителю №1)
try:
    rep = erp.ВнешниеОтчеты.Создать(ERF)
    tz = rep.ПолучитьДанные(None)
    plan = {}
    for i in range(tz.Количество()):
        row = tz.Получить(i)
        if S(row.ВидДокумента) == "0. План" and "15 м №1" in S(row.ПодразделениеИсполнитель):
            key = S(row.ОбщееНазвание)
            plan[key] = plan.get(key, 0.0) + row.ПланКол
    for name, expect in (("Бойлер електричний", 2), ("Вимикач", 10), ("Вітробарєр", 42)):
        check(f"отчёт ПланКол {name}={expect}", abs(plan.get(name, 0) - expect) < 0.001,
              str(plan.get(name)))
    check("печать Норма == отчёт План (Вимикач)", group_norms.get("Вимикач") == plan.get("Вимикач"))
    check("печать Норма == отчёт План (Вітробарєр)", group_norms.get("Вітробарєр") == plan.get("Вітробарєр"))
except Exception as e:
    check("кросс-сверка с отчётом", False, repr(e)[:200])

print(f"\n{'='*50}\nИТОГ: {'ALL PASS' if not FAILS else 'FAILS: ' + ', '.join(FAILS)}")
sys.exit(0 if not FAILS else 1)
```

- [ ] **Step 7.2: Запустить**

Run: `python -X utf8 _Rarzrabotki/Python/test/smoke_doc_kompl_pechat_ceh.py`
Expected: `ALL PASS`, exit 0. Известные точки отладки: (а) СКД молча выпал → проверить UUID обёртки уникален и `<Template>` в корневом xml; (б) «Поле не найдено КоментарЦеху» → выражение calculatedField; (в) кросс-сверка: если `ПолучитьДанные(None)` падает на None-периоде → создать `П = erp.NewObject("СтандартныйПериод")` и передать его (пустой период = весь диапазон); (г) заливка возвращает (-1,-1,-1) → цвет задан не #RRGGBB (грабля §36).

- [ ] **Step 7.3: Commit**

```bash
git add _Rarzrabotki/Python/test/smoke_doc_kompl_pechat_ceh.py
git commit -m "test(raschet-kompl): smoke печати «Аналіз для цеху» + кросс-сверка с отчётом"
```

---

### Task 8: Smoke формулы `ЕкономіяСума`

**Files:**
- Create: `_Rarzrabotki/Python/test/smoke_doc_kompl_ekonomia_summa.py`

- [ ] **Step 8.1: Написать smoke** (scratch-док get-or-create по маркеру, документы НЕ удалять — правило `feedback_no_doc_delete_in_tests`)

```python
# -*- coding: utf-8 -*-
"""Smoke формулы ЕкономіяСума (= Економія × НормаСум/Норма) после пересчёта.

Scratch-док с маркером SMOKE_EKON_SUM_v1 (get-or-create), спецификация и склад дома №1.
Инварианты (живые данные дрейфуют — абсолюты не проверяем):
 A. ЕкономіяСума > 0 <=> Экономия > 0 (по каждому эталону);
 B. ЕкономіяСума == Окр(Экономия × НормаСумма/Норма, 2) ± 0.02;
 C. построчно ВНорме + ПонадНорму == Остаток;
 D. есть >= 1 эталон с Экономия > 0 и ЕкономіяСума > 0.
"""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

MARKER = "SMOKE_EKON_SUM_v1"

v8 = win32com.client.Dispatch("V83.COMConnector")
try:
    erp = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
except Exception:
    erp = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = erp.String

FAILS = []
def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f" | {detail}" if detail else ""))
    if not cond:
        FAILS.append(name)

# эталонный док №3 — источник спецификации и склада
q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д
ГДЕ Д.Номер = "00000000003" И НЕ Д.ПометкаУдаления"""
doc3 = q.Execute().Выгрузить().Получить(0).Ссылка.ПолучитьОбъект()

# get-or-create scratch-дока по маркеру (Комментарий — неограниченная строка: ВЫРАЗИТЬ)
q2 = erp.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д
ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления"""
q2.SetParameter("М", MARKER)
r2 = q2.Execute().Выгрузить()
if r2.Количество() > 0:
    doc = r2.Получить(0).Ссылка.ПолучитьОбъект()
    print(f"scratch-док найден: №{S(doc.Номер)}")
else:
    doc = erp.Документы.РасчетКомплектаций.СоздатьДокумент()
    doc.Дата = erp.ТекущаяДата()
    doc.Заполнить(None)
    doc.Комментарий = MARKER
    print("scratch-док создаётся")

doc.Спецификация = doc3.Спецификация
doc.Период = doc3.Период
doc.Организация = doc3.Организация
doc.СкладыОстатков.Очистить()
for i in range(doc3.СкладыОстатков.Количество()):
    doc.СкладыОстатков.Добавить().Склад = doc3.СкладыОстатков.Получить(i).Склад
doc.СчетаОстатков.Очистить()
for i in range(doc3.СчетаОстатков.Количество()):
    doc.СчетаОстатков.Добавить().Счет = doc3.СчетаОстатков.Получить(i).Счет
doc.СчетаМалоценки.Очистить()
for i in range(doc3.СчетаМалоценки.Количество()):
    doc.СчетаМалоценки.Добавить().Счет = doc3.СчетаМалоценки.Получить(i).Счет

doc.РассчитатьАнализ()
doc.Записать()

tch = doc.ТабличнаяЧастьОстатков
print(f"строк после пересчёта: {tch.Количество()}")
et = {}
bad_c = 0
for i in range(tch.Количество()):
    s = tch.Получить(i)
    d = et.setdefault(S(s.ОбщееНазвание),
                      {"norma": 0.0, "nsum": 0.0, "ek": 0.0, "eksum": 0.0})
    d["norma"] += s.Норма; d["nsum"] += s.НормаСумма
    d["ek"] += s.Экономия; d["eksum"] += s.ЕкономіяСума
    if abs(s.ВНорме + s.ПонадНорму - s.Остаток) > 0.001:
        bad_c += 1
check("C: построчный баланс", bad_c == 0, f"нарушений {bad_c}")

bad_a, bad_b, pos = 0, 0, 0
for name, d in et.items():
    if (d["ek"] > 0.0005) != (d["eksum"] > 0.005):
        bad_a += 1
        print(f"  A-нарушение {name!r}: ек={d['ek']} сум={d['eksum']}")
    if d["norma"] > 0:
        expect = round(d["ek"] * d["nsum"] / d["norma"], 2)
        if abs(d["eksum"] - expect) > 0.02:
            bad_b += 1
            print(f"  B-нарушение {name!r}: сум={d['eksum']} ожид={expect}")
    if d["ek"] > 0 and d["eksum"] > 0:
        pos += 1
check("A: ЕкономіяСума>0 <=> Економія>0", bad_a == 0, f"нарушений {bad_a}")
check("B: сумма = кол × нормовая цена", bad_b == 0, f"нарушений {bad_b}")
check("D: есть эталоны с экономией", pos >= 1, f"pos={pos}")
tot = sum(d["eksum"] for d in et.values())
print(f"Σ ЕкономіяСума по scratch-доку: {tot:.2f} грн (справочно; по ТЧ дока №3 старая была 109072.97)")

print(f"\n{'='*50}\nИТОГ: {'ALL PASS' if not FAILS else 'FAILS: ' + ', '.join(FAILS)}")
sys.exit(0 if not FAILS else 1)
```

- [ ] **Step 8.2: Запустить**

Run: `python -X utf8 _Rarzrabotki/Python/test/smoke_doc_kompl_ekonomia_summa.py`
Expected: `ALL PASS`, exit 0. Примечание: остатки дома №1 на 03.08 после проведённых списаний могли обнулиться — это НЕ ошибка (инварианты держатся на любых данных); если строк 0 — сменить в scratch-доке `Период` на `doc3.Период` минус 1 день (остатки до списания) и перезапустить.

- [ ] **Step 8.3: Commit**

```bash
git add _Rarzrabotki/Python/test/smoke_doc_kompl_ekonomia_summa.py
git commit -m "test(raschet-kompl): smoke формулы ЕкономіяСума"
```

---

### Task 9: Регрессия существующих печатей и расчёта

- [ ] **Step 9.1:** `python -X utf8 _Rarzrabotki/Python/test/smoke_doc_kompl_analiz.py` → PASS (инварианты расчёта целы; тест НЕ проверяет старую формулу ЕкономіяСума — если упадёт именно на ней, обновить ожидание теста на новую формулу и зафиксировать в коммите).
- [ ] **Step 9.2:** `python -X utf8 _Rarzrabotki/Python/test/smoke_doc_kompl_pechat.py` → PASS (7 существующих печатей не сломаны).
- [ ] **Step 9.3:** `python -X utf8 _Rarzrabotki/Python/test/smoke_doc_kompl_m29.py` → PASS (М-29 независим).
- [ ] **Step 9.4:** Итоговый прогон нового смоука ещё раз (`smoke_doc_kompl_pechat_ceh.py`) → PASS. Commit при правках тестов.

---

### Task 10: Знания, память, интеграция

- [ ] **Step 10.1: База знаний** — `_Rarzrabotki/notebook/knowledge_Бухгалтерия_РасчетКомплектаций/`:
  - `08_pechatnye_formy.md`: добавить печать №10 (метод, макет, колонки, «только количество», ширина Коментар цеху кодом после вывода СКД, per-level selection, жирные группы через СА группировки);
  - `03_algoritm.md`: раздел «Економія» — новая формула суммы + причина (решение бухгалтера 2026-08-06, кейс «Гайка»);
  - `KNOWLEDGE_MAP.md`: строка экспортного метода `СформироватьПечатьАнализЦех`, макет, максимальный id формы теперь **576**;
  - `LESSONS.md`: новые грабли, если встретились при выполнении.
- [ ] **Step 10.2: Память** — обновить `raschet_komplektacij_dokument.md` (10-я печать, формула, id 576) одним абзацем.
- [ ] **Step 10.3: Спека** — отметить статус «✅ реализовано» в разделе доработок KNOWLEDGE_MAP (таблица «Доработки»).
- [ ] **Step 10.4: Commit** знаний: `docs(raschet-kompl): знання — печать для цеху + формула економії`.
- [ ] **Step 10.5: Интеграция веток** — по канону CLAUDE.md (worktree → `claude/main`): в основном каталоге `git status`; если он на `claude/main` и чист — `git merge --no-ff <ветка worktree>` (или cherry-pick коммитов задачи). НЕ пушить `main` (112MB-правило); origin — только по явной просьбе пользователя. Сообщить пользователю итог: что загружено в базу, какие коммиты, что осталось (правка 18 нулевых строк СС — техотдел).

---

## Self-review (выполнен)

- Спека-покрытие: печать (Task 1,2,4), только количество + Норма СС + Коментар цеху (Task 1), економія/норма на группе (Task 1 selections), ч/б стили (Task 1 шаг 7,9), формула (Task 3), приёмка+кросс-сверка (Task 7,8,9), деплой (Task 6), знания (Task 10). Эталоны без выдачи — не показываются: отдельного кода не требует (их нет в ТЧ) ✓.
- Заглушек нет; сигнатуры согласованы: `СформироватьПечатьАнализЦех(ДанныеАнализа)` = Task 2/4/7; `ЗафиксироватьЭкономиюЭталона(СтрокаЭталона, ОстатокНормы, НормаЭталона, НормаСумЭталона)` = Task 3 (объявление + 2 вызова).
- Риск-точки помечены прямо в шагах: якоря builder-скрипта (1.2), пустой `<dcsset:filter/>` не используется (СА группы без filter), порядок элементов СА группы проверяется загрузкой (6.3) и выводом (7.2); ширина колонки — кодом после вывода (2.1), а не через СКД.
