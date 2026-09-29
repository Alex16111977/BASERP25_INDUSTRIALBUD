# -*- coding: utf-8 -*-
"""Перекомпонування форми документа «Фінансовий звіт виробництва» (29.09.2026) за каноном РасчетКомплектаций.
Точковий патч живого Form.xml (не перегенерація). Ідемпотентність: повторний запуск — assert «вже зроблено».
Вкладки: Звіт (Розрахунок + документи рядка) → Контроль → Казна → Параметри; коментар — внизу форми."""
import io, re, sys, uuid
from xml.dom import minidom

sys.stdout.reconfigure(encoding="utf-8")
ФОРМА = (r"C:\Configuration_downloads\BASERP25\.claude\worktrees\director-report-letter-689b66\_Rarzrabotki\BASEBuh"
         r"\Documents\А_ФинансовыйОтчетПроизводства\Forms\ФормаДокумента\Ext\Form.xml")
s = io.open(ФОРМА, encoding="utf-8-sig", newline="").read()
CRLF = "\r\n" in s
s = s.replace("\r\n", "\n")
assert 'name="ГруппаНастройкиРасчета"' not in s, "форму вже перекомпоновано"


def блок(tag, name):
    """(початок рядка, кінець рядка) елемента з урахуванням вкладених однойменних тегів."""
    m = re.search(rf'\n([ \t]*)<{tag} name="{re.escape(name)}" id="\d+"(/?)>', s)
    assert m, (tag, name)
    start = m.start() + 1
    if m.group(2) == "/":
        return start, s.index("\n", m.end()) + 1
    pat = re.compile(rf"<{tag}\b[^>]*?(/?)>|</{tag}>")
    pos, depth = m.end(), 1
    while depth:
        mm = pat.search(s, pos)
        if mm.group(0).startswith("</"):
            depth -= 1
        elif mm.group(1) != "/":
            depth += 1
        pos = mm.end()
    return start, s.index("\n", pos) + 1


def вирізати(tag, name):
    global s
    a, b = блок(tag, name)
    текст = s[a:b]
    s = s[:a] + s[b:]
    return текст


def відступ(текст):
    return re.match(r"[ \t]*", текст).group(0)


def переіндентувати(текст, новий):
    старий = відступ(текст)
    return "".join((новий + р[len(старий):] if р.startswith(старий) else р) if р.strip() else р
                   for р in текст.splitlines(keepends=True))


def вставити_після(tag, name, текст):
    global s
    a, b = блок(tag, name)
    s = s[:b] + переіндентувати(текст, відступ(s[a:b])) + s[b:]


def вставити_перед(tag, name, текст):
    global s
    a, b = блок(tag, name)
    s = s[:a] + переіндентувати(текст, відступ(s[a:b])) + s[a:]


def новий_id():
    global наступний_id
    наступний_id += 1
    return наступний_id


наступний_id = max(int(x) for x in re.findall(r'<(?!Command\b|Attribute\b|Column\b)\w+ name="[^"]+" id="(\d+)"', s))


def заголовок(текст_блоку, старий, новий):
    assert текст_блоку.count(f"<v8:content>{старий}</v8:content>") >= 1, старий
    return текст_блоку.replace(f"<v8:content>{старий}</v8:content>", f"<v8:content>{новий}</v8:content>")


# --- 1. Командна панель: Excel — окремою кнопкою після «Розрахувати»; «Заповнити підрозділи» — на вкладку «Параметри»
excel = вирізати("Button", "КнопкаСохранитьExcel")
excel = excel.replace("<Type>UsualButton</Type>", "<Type>UsualButton</Type>\n\t<Representation>Text</Representation>".replace(
    "\n\t", "\n" + відступ(excel) + "\t"), 1)
вставити_після("Button", "КнопкаРассчитать", excel)
заповнити = вирізати("Button", "КнопкаЗаполнитьПодразделения")
заповнити = заповнити.replace("<Type>UsualButton</Type>", "<Type>Hyperlink</Type>")
заповнити = re.sub(r"\n[ \t]*<Representation>Text</Representation>", "", заповнити)
заповнити = заголовок(заповнити, "Заповнити підрозділи", "Заповнити за напрямком")
казна_кн_a, казна_кн_b = блок("Button", "КнопкаЗаполнитьИзКазны")
кн = s[казна_кн_a:казна_кн_b]
s = s[:казна_кн_a] + заголовок(кн, "Заповнити", "Заповнити з казни") + s[казна_кн_b:]
вставити_після("Button", "КнопкаЗаполнитьИзКазны", заповнити)

# --- 2. Шапка: другорядні поля — у групу «Налаштування розрахунку» на вкладці «Параметри»
поля = {n: вирізати(t, n) for t, n in (("InputField", "ПолеНачалоПериода"), ("CheckBoxField", "ПолеВключатьНепроведенные"),
                                        ("CheckBoxField", "ПолеЗаполнятьТолькоСДвижениями"),
                                        ("InputField", "ПолеДокументСравнения"), ("InputField", "ПолеОтветственный"))}
a, b = блок("UsualGroup", "ГруппаШапкаЛево")
шаблон_підгрупи = s[a:b]
шапка_підгрупи = шаблон_підгрупи[:шаблон_підгрупи.index("<ChildItems>")]
a, b = блок("UsualGroup", "ГруппаШапка")
шаблон_групи = s[a:b]
шапка_групи = шаблон_групи[:шаблон_групи.index("<ChildItems>")]


def група(шапка, ім, діти, стара_назва):
    т = шапка.replace(f'name="{стара_назва}"', f'name="{ім}"')
    т = re.sub(rf'<UsualGroup name="{ім}" id="\d+">', f'<UsualGroup name="{ім}" id="{новий_id()}">', т, count=1)
    т = re.sub(r'<ExtendedTooltip name="[^"]+" id="\d+"/>',
               f'<ExtendedTooltip name="{ім}РасширеннаяПодсказка" id="{новий_id()}"/>', т, count=1)
    в = відступ(т)
    return т + "<ChildItems>\n" + "".join(переіндентувати(д, в + "\t\t") for д in діти) + в + "\t</ChildItems>\n" + в + "</UsualGroup>\n"


ліво = група(шапка_підгрупи, "ГруппаНастройкиЛево", [поля["ПолеНачалоПериода"], поля["ПолеВключатьНепроведенные"],
                                                      поля["ПолеЗаполнятьТолькоСДвижениями"]], "ГруппаШапкаЛево")
право = група(шапка_підгрупи, "ГруппаНастройкиПраво", [поля["ПолеДокументСравнения"], поля["ПолеОтветственный"]],
              "ГруппаШапкаЛево")
налаштування = група(шапка_групи, "ГруппаНастройкиРасчета", [ліво, право], "ГруппаШапка")
вставити_перед("UsualGroup", "ГруппаКомандыПодразделений", налаштування)

# --- 3. Сторінки: Звіт (Розрахунок + документи рядка) → Контроль → Казна → Параметри; без «Розшифровки» і «Коментаря»
статті = вирізати("Table", "СтатьиИсключенияКазны")
вставити_після("UsualGroup", "ГруппаУточнения", статті)
розш = вирізати("Table", "РасшифровкаДокументов")
розш = re.sub(r'(<Table name="РасшифровкаДокументов" id="\d+">\n(?:[ \t]*<CommandBarLocation>[^\n]*\n)?'
              r'[ \t]*<ReadOnly>true</ReadOnly>\n)', lambda m: m.group(1)
              + відступ(розш) + "\t<HeightInTableRows>8</HeightInTableRows>\n"
              + відступ(розш) + "\t<VerticalStretch>false</VerticalStretch>\n", розш, count=1)
for колонка in ("РасшПодразделение", "РасшКонтрагент", "РасшДоговор"):
    розш = re.sub(rf'(\n([ \t]*)<InputField name="{колонка}" id="\d+">\n)',
                  lambda m: m.group(1) + m.group(2) + "\t<Visible>false</Visible>\n", розш, count=1)
вставити_після("Table", "Расчет", розш)
# подвійний клік по рядку розрахунку — параметри договору
a, b = блок("Table", "Расчет")
т = s[a:b]
т = т.replace("<Event name=\"OnActivateRow\">РасчетПриАктивизацииСтроки</Event>",
              "<Event name=\"OnActivateRow\">РасчетПриАктивизацииСтроки</Event>\n" + відступ(т) + "\t\t\t"
              + "<Event name=\"Selection\">РасчетВыборСтроки</Event>", 1)
s = s[:a] + т + s[b:]

коментар = вирізати("InputField", "ПолеКомментарий").replace("<TitleLocation>None</TitleLocation>",
                                                             "<TitleLocation>Left</TitleLocation>")
вирізати("Page", "СтраницаРасшифровка")
вирізати("Page", "СтраницаКомментарий")
сторінки = {n: вирізати("Page", n) for n in ("СтраницаРасчет", "СтраницаКазна", "СтраницаРезультат", "СтраницаКонтроль")}
сторінки["СтраницаРезультат"] = заголовок(сторінки["СтраницаРезультат"], "Результат", "Звіт")
сторінки["СтраницаРасчет"] = заголовок(сторінки["СтраницаРасчет"], "Розрахунок", "Параметри")
a, b = блок("Pages", "Страницы")
т = s[a:b]
кінець_дітей = т.rindex("</ChildItems>")
в_стор = відступ(т) + "\t\t"
т = т[:кінець_дітей].rstrip("\t") + "".join(переіндентувати(сторінки[n], в_стор) for n in
                                            ("СтраницаРезультат", "СтраницаКонтроль", "СтраницаКазна", "СтраницаРасчет")) \
    + відступ(т) + "\t" + т[кінець_дітей:]
s = s[:a] + т + s[b:]
вставити_після("Pages", "Страницы", коментар)

# --- перевірки
ЕЛ = r'<(?!Command\b|Attribute\b|Column\b)\w+ name="[^"]+" id="'   # у реквізитів/команд — свої простори id
dup = [x for x in set(re.findall(ЕЛ + r'(\d+)"', s)) if len(re.findall(ЕЛ + x + '"', s)) > 1]
assert not dup, dup
for n in ("СтраницаРасшифровка", "СтраницаКомментарий", "ГруппаФильтрРасшифровки"):
    assert f'name="{n}"' not in s, n
порядок = re.findall(r'<Page name="(\w+)"', s)
assert порядок == ["СтраницаРезультат", "СтраницаКонтроль", "СтраницаКазна", "СтраницаРасчет"], порядок
if CRLF:
    s = s.replace("\n", "\r\n")
io.open(ФОРМА, "w", encoding="utf-8-sig", newline="").write(s)
minidom.parse(ФОРМА)
print("OK: сторінки", порядок, "; Excel у панелі; налаштування на «Параметри»; розшифровка під звітом; коментар внизу")
