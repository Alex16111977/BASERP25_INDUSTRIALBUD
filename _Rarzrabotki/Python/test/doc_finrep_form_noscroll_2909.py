# -*- coding: utf-8 -*-
"""Форма фінзвіту без прокрутки (канон bas-form-layout / РасчетКомплектаций), 29.09.2026. Точковий патч Form.xml:
корінь VerticalScroll=useIfNecessary; шапка в один ряд; налаштування — 2 ряди; «Параметри»: пара
[Підрозділи | Договори додатково] — єдиний «гумовий» елемент, пара [Рахунки | Статті-виключення] по 3 рядки;
розшифровка під звітом — 6 рядків; коментар однорядковий."""
import io, re, sys
from xml.dom import minidom

sys.stdout.reconfigure(encoding="utf-8")
src = io.open(__file__.replace("doc_finrep_form_noscroll_2909.py", "doc_finrep_form_redesign_2909.py"),
              encoding="utf-8").read()
ФОРМА = re.search(r'ФОРМА = \((.*?)\)\n', src, re.S).group(0)
exec(ФОРМА)
s = io.open(ФОРМА, encoding="utf-8-sig", newline="").read()
CRLF = "\r\n" in s
s = s.replace("\r\n", "\n")
exec(src[src.index("def блок("):src.index("наступний_id = ")])   # блок/вирізати/вставити_*/переіндентувати
assert 'name="ГруппаПодразделенияИДоговоры"' not in s, "вже виконано"
наступний_id = max(int(x) for x in re.findall(r'<(?!Command\b|Attribute\b|Column\b)\w+ name="[^"]+" id="(\d+)"', s))


def новий_id():
    global наступний_id
    наступний_id += 1
    return наступний_id


def замінити_в_блоці(tag, name, old, new):
    global s
    a, b = блок(tag, name)
    т = s[a:b]
    assert old in т, (name, old)
    s = s[:a] + т.replace(old, new, 1) + s[b:]


# 1. корінь: прокрутка — лише аварійний фолбек
assert s.count("\n\t<CommandBarLocation>Top</CommandBarLocation>\n") == 1
s = s.replace("\n\t<CommandBarLocation>Top</CommandBarLocation>\n",
              "\n\t<CommandBarLocation>Top</CommandBarLocation>\n\t<VerticalScroll>useIfNecessary</VerticalScroll>\n", 1)
# 2. шапка в один ряд
замінити_в_блоці("UsualGroup", "ГруппаШапкаЛево", "<Group>Vertical</Group>", "<Group>Horizontal</Group>")
# 3. налаштування — два ряди
замінити_в_блоці("UsualGroup", "ГруппаНастройкиРасчета", "<Group>Horizontal</Group>", "<Group>Vertical</Group>")
замінити_в_блоці("UsualGroup", "ГруппаНастройкиЛево", "<Group>Vertical</Group>", "<Group>Horizontal</Group>")
замінити_в_блоці("UsualGroup", "ГруппаНастройкиПраво", "<Group>Vertical</Group>", "<Group>Horizontal</Group>")
# 4. «Параметри»: [Підрозділи | Договори додатково] — гумовий ряд; [Рахунки | Статті] — 3 рядки
a, b = блок("UsualGroup", "ГруппаУточнения")
шапка_групи = s[a:b][:s[a:b].index("<ChildItems>")]
допдог = вирізати("Table", "ДоговорыДополнительно")
допдог = re.sub(r"\n[ \t]*<HeightInTableRows>3</HeightInTableRows>\n[ \t]*<VerticalStretch>false</VerticalStretch>", "",
                допдог, count=1)
підрозділи = вирізати("Table", "Подразделения")
ім = "ГруппаПодразделенияИДоговоры"
нова = шапка_групи.replace('name="ГруппаУточнения"', f'name="{ім}"')
нова = re.sub(rf'<UsualGroup name="{ім}" id="\d+">', f'<UsualGroup name="{ім}" id="{новий_id()}">', нова, count=1)
нова = re.sub(r'<ExtendedTooltip name="[^"]+" id="\d+"/>', f'<ExtendedTooltip name="{ім}РасширеннаяПодсказка" id="{новий_id()}"/>',
              нова, count=1)
в = відступ(нова)
нова = нова + "<ChildItems>\n" + переіндентувати(підрозділи, в + "\t\t") + переіндентувати(допдог, в + "\t\t") \
    + в + "\t</ChildItems>\n" + в + "</UsualGroup>\n"
вставити_перед("UsualGroup", "ГруппаУточнения", нова)
статті = вирізати("Table", "СтатьиИсключенияКазны")
вставити_після("Table", "СчетаРасчетов", статті)
# 5. розшифровка під звітом — 6 рядків
замінити_в_блоці("Table", "РасшифровкаДокументов", "<HeightInTableRows>8</HeightInTableRows>",
                 "<HeightInTableRows>6</HeightInTableRows>")
# 6. коментар однорядковий
замінити_в_блоці("InputField", "ПолеКомментарий", "<HorizontalStretch>true</HorizontalStretch>",
                 "<HorizontalStretch>true</HorizontalStretch>\n\t\t<VerticalStretch>false</VerticalStretch>\n\t\t<MultiLine>false</MultiLine>")

ЕЛ = r'<(?!Command\b|Attribute\b|Column\b)\w+ name="[^"]+" id="'
dup = [x for x in set(re.findall(ЕЛ + r'(\d+)"', s)) if len(re.findall(ЕЛ + x + '"', s)) > 1]
assert not dup, dup
if CRLF:
    s = s.replace("\n", "\r\n")
io.open(ФОРМА, "w", encoding="utf-8-sig", newline="").write(s)
minidom.parse(ФОРМА)
print("OK: без прокрутки — шапка 1 ряд, налаштування 2 ряди, пари таблиць на «Параметрах», розшифровка 6, коментар 1 рядок")
