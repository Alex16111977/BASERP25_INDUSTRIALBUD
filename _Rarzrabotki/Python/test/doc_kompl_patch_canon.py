# -*- coding: utf-8 -*-
"""Перекомпоновка формы РасчетКомплектаций по канону типовых («один резиновый элемент»):
1) корень: VerticalScroll=useIfNecessary;
2) «Параметри»: 4 мини-таблицы в горизонтальный ряд, VerticalStretch=false;
3) «Аналіз»: таблица без панели и пустого подвала (stretch по умолчанию), «Підсумки» в одну строку (Left);
4) списания: итоги в подвал таблиц (Footer + FooterDataPath=Total*), группы итогов и 4 реквизита удалены;
5) «Документи»: таблицам VerticalStretch=false; заголовки групп убраны.
Модуль: удаление ПересчитатьИтогиСписания и вызовов. Патч ТЕКУЩЕЙ формы из базы (19.07 11:46)."""
import re, sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

D = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext"
FX = D + r"\Form.xml"
FM = D + r"\Form\Module.bsl"


def read(p):
    raw = open(p, 'rb').read()
    return (raw.decode('utf-8-sig').replace('\r\n', '\n'),
            raw.startswith(b'\xef\xbb\xbf'), b'\r\n' in raw)


def write(p, text, bom, crlf):
    if crlf:
        text = text.replace('\n', '\r\n')
    open(p, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + text.encode('utf-8'))


def cut_node(text, open_tag_start, close_literal):
    """Вырезает узел от позиции открытия до соответствующего close_literal (без вложенных того же типа)."""
    j = text.index(close_literal, open_tag_start) + len(close_literal)
    return text[:open_tag_start] + text[j:], text[open_tag_start:j]


text, bom, crlf = read(FX)
max_id = max(int(m) for m in re.findall(r'id="(-?\d+)"', text))
nid = iter(range(max_id + 1, max_id + 30))

# --- 1. VerticalScroll (позиция как в типовых: перед AutoTime) ---
assert "VerticalScroll" not in text
text = text.replace("\t<AutoTime>", "\t<VerticalScroll>useIfNecessary</VerticalScroll>\n\t<AutoTime>", 1)
print("OK 1: VerticalScroll=useIfNecessary")

# --- 2. «Параметри»: обернуть 4 таблицы в горизонтальную группу, VerticalStretch=false ---
start = text.index('\t\t\t\t\t\t<Table name="СкладыОстатков"')
end_marker = '</Table>\n'
# конец таблицы Этапы
et = text.index('<Table name="Этапы"')
end = text.index(end_marker, et) + len(end_marker)
# найдём отступ конца: у Этапы </Table> с 6 табами
block = text[start:end]
assert block.count('<Table name="') == 4
# VerticalStretch=false каждой (после HeightInTableRows)
block = block.replace("<HeightInTableRows>3</HeightInTableRows>\n",
                      "<HeightInTableRows>3</HeightInTableRows>\n"
                      + "\t\t\t\t\t\t\t\t<VerticalStretch>false</VerticalStretch>\n")
# сдвиг на 1 таб глубже
block = "".join(("\t" + line if line.strip() else line) for line in block.splitlines(keepends=True))
g1, g2 = next(nid), next(nid)
wrapped = (f'\t\t\t\t\t\t<UsualGroup name="ГруппаТаблицыНастроек" id="{g1}">\n'
           f'\t\t\t\t\t\t\t<Group>Horizontal</Group>\n'
           f'\t\t\t\t\t\t\t<Behavior>Usual</Behavior>\n'
           f'\t\t\t\t\t\t\t<Representation>None</Representation>\n'
           f'\t\t\t\t\t\t\t<ShowTitle>false</ShowTitle>\n'
           f'\t\t\t\t\t\t\t<ExtendedTooltip name="ГруппаТаблицыНастроекРасширеннаяПодсказка" id="{g2}"/>\n'
           f'\t\t\t\t\t\t\t<ChildItems>\n'
           + block +
           f'\t\t\t\t\t\t\t</ChildItems>\n'
           f'\t\t\t\t\t\t</UsualGroup>\n')
text = text[:start] + wrapped + text[end:]
print("OK 2: 4 мини-таблицы в горизонтальном ряду, stretch=false")

# --- 3a. ТаблицаАнализа: без командной панели и пустого подвала ---
anchor = ('<Table name="ТаблицаАнализа" id="45">\n'
          '\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n')
assert anchor in text
text = text.replace(anchor,
                    '<Table name="ТаблицаАнализа" id="45">\n'
                    '\t\t\t\t\t\t\t<CommandBarLocation>None</CommandBarLocation>\n'
                    '\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n', 1)
footer = '\t\t\t\t\t\t\t<Footer>true</Footer>\n'
assert text.count(footer) == 1
text = text.replace(footer, '', 1)
print("OK 3a: панель и пустой подвал таблицы анализа убраны")

# --- 3b. Заголовки групп Відбір/Підсумки/Шапка убрать ---
removed_titles = 0
for gname in ("ГруппаШапка", "ГруппаВідбір", "ГруппаПідсумки"):
    gi = text.find(f'<UsualGroup name="{gname}"')
    if gi < 0:
        continue
    ti = text.find("<Title>", gi)
    ci = text.find("<ChildItems>", gi)
    if 0 < ti < ci:
        line_start = text.rindex("\n", 0, ti) + 1
        tend = text.index("</Title>\n", ti) + len("</Title>\n")
        text = text[:line_start] + text[tend:]
        removed_titles += 1
print(f"OK 3b: заголовков групп убрано: {removed_titles}")

# --- 4. Удалить группы итогов списаний (целиком) ---
for gname in ("ГруппаИтогиПоНормам", "ГруппаИтогиСверхНормы"):
    gi = text.index(f'<UsualGroup name="{gname}"')
    line_start = text.rindex("\n", 0, gi) + 1
    text, cut = cut_node(text, line_start, "\t\t\t\t\t\t</UsualGroup>\n")
    assert gname in cut
print("OK 4a: группы итогов списаний удалены")

# --- 4b. Подвал с автоитогами таблицам списания ---
for ts in ("СписаниеПоНормам", "СписаниеСверхНормы"):
    dp = f"\t\t\t\t\t\t\t<DataPath>Объект.{ts}</DataPath>\n"
    assert text.count(dp) == 1, ts
    text = text.replace(dp, f"\t\t\t\t\t\t\t<Footer>true</Footer>\n" + dp, 1)
    for col in ("Количество", "Сумма"):
        cdp = f"<DataPath>Объект.{ts}.{col}</DataPath>\n"
        i = text.index(cdp)
        j = i + len(cdp)
        indent = "\t" * 9
        text = text[:j] + f"{indent}<FooterDataPath>Объект.{ts}.Total{col}</FooterDataPath>\n" + text[j:]
print("OK 4b: Footer + Total-итоги у таблиц списания")

# --- 3c. Поля «Підсумки»: подписи слева (одна строка) ---
n_top = text.count("<TitleLocation>Top</TitleLocation>")
text = text.replace("<TitleLocation>Top</TitleLocation>", "<TitleLocation>Left</TitleLocation>")
print(f"OK 3c: TitleLocation Top->Left: {n_top}")

# --- 5. «Документи»: таблицам stretch=false ---
n_doc = 0
for ts in ("ДокументиКомплектації", "ДокументиМалоценки"):
    gi = text.index(f'<Table name="{ts}"')
    hi = text.index("<HeightInTableRows>3</HeightInTableRows>\n", gi)
    j = hi + len("<HeightInTableRows>3</HeightInTableRows>\n")
    text = text[:j] + "\t\t\t\t\t\t\t<VerticalStretch>false</VerticalStretch>\n" + text[j:]
    n_doc += 1
print(f"OK 5: stretch=false у таблиц документов: {n_doc}")

# --- 6. Удалить 4 реквизита формы Итого(Норм/Сверх)* ---
for attr in ("ИтогоНормКоличество", "ИтогоНормСумма", "ИтогоСверхКоличество", "ИтогоСверхСумма"):
    ai = text.index(f'<Attribute name="{attr}"')
    line_start = text.rindex("\n", 0, ai) + 1
    text, cut = cut_node(text, line_start, "\t\t</Attribute>\n")
    assert attr in cut
print("OK 6: 4 реквизита итогов списаний удалены")

# --- Контроли Form.xml ---
ids = re.findall(r'id="(-?\d+)"', text)
assert len(ids) == len(set(ids)), "дубли id"
assert text.count("<Footer>true</Footer>") == 2
assert text.count("<FooterDataPath>") == 4
assert "ГруппаИтогиПоНормам" not in text and "ГруппаИтогиСверхНормы" not in text
assert "ИтогоНормКоличество" not in text
write(FX, text, bom, crlf)
xml.dom.minidom.parse(FX)
print(f"Form.xml OK: id={len(ids)}, max={max(int(x) for x in ids)}")

# --- Модуль: удалить ПересчитатьИтогиСписания и вызовы ---
mod, mbom, mcrlf = read(FM)
n_calls = mod.count("\tПересчитатьИтогиСписания();\n")
mod = mod.replace("\tПересчитатьИтогиСписания();\n", "")
pi = mod.index("Процедура ПересчитатьИтогиСписания()")
# захватить директиву и комментарий выше
line_start = mod.rindex("&НаКлиенте", 0, pi)
pend = mod.index("КонецПроцедуры", pi) + len("КонецПроцедуры")
# хвостовые переводы строк
while pend < len(mod) and mod[pend] == "\n":
    pend += 1
mod = mod[:line_start] + mod[pend:]
n_proc = len(re.findall(r"^(?:Процедура|Функция) ", mod, re.M))
n_end = len(re.findall(r"^Конец(?:Процедуры|Функции)", mod, re.M))
assert n_proc == n_end, f"непарность {n_proc}/{n_end}"
assert "ПересчитатьИтогиСписания" not in mod
write(FM, mod, mbom, mcrlf)
print(f"Module.bsl OK: убрано вызовов={n_calls}, процедур={n_proc}")
print("CANON PATCH PASS")
