# -*- coding: utf-8 -*-
"""«Параметри»: 4 таблицы настроек в сетку 2×2 (канон: 2 таблицы в ряд, ChildItemsWidth=Equal),
чтобы не было ни горизонтальной, ни вертикальной прокрутки.
Ряд1: Склади залишків | Рахунки залишків
Ряд2: Рахунки малоцінки | (Етапи + чекбокс «Виключати вибрані етапи» ПОД Етапи).
Патч текущей формы из базы."""
import re, sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

F = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml"

raw = open(F, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
crlf = b'\r\n' in raw
text = raw.decode('utf-8-sig').replace('\r\n', '\n')

max_id = max(int(m) for m in re.findall(r'id="(-?\d+)"', text))
nid = iter(range(max_id + 1, max_id + 30))


def extract(node_open_re, close_literal, base_indent):
    """Вырезает узел (по regex открытия) целиком со строки; возвращает (text_без_узла, вырезанный_текст)."""
    global text
    m = re.search(node_open_re, text)
    assert m, node_open_re
    ls = text.rindex("\n", 0, m.start()) + 1
    je = text.index(base_indent + close_literal, m.start()) + len(base_indent + close_literal)
    node = text[ls:je]
    text = text[:ls] + text[je:]
    return node


def reindent(node, delta_tabs):
    pad = "\t" * delta_tabs
    return "".join((pad + line if line.strip() else line) for line in node.splitlines(keepends=True))


# --- вырезать 4 таблицы (7 табов) и чекбокс (6 табов) ---
t_sklady = extract(r'\t{7}<Table name="СкладыОстатков" id="\d+">', "</Table>\n", "\t" * 7)
t_scheta = extract(r'\t{7}<Table name="СчетаОстатков" id="\d+">', "</Table>\n", "\t" * 7)
t_malo = extract(r'\t{7}<Table name="СчетаМалоценки" id="\d+">', "</Table>\n", "\t" * 7)
t_etapy = extract(r'\t{7}<Table name="Этапы" id="\d+">', "</Table>\n", "\t" * 7)
cb = extract(r'\t{6}<CheckBoxField name="ИсключатьЭтапы" id="\d+">', "</CheckBoxField>\n", "\t" * 6)

# от старой пустой группы ГруппаТаблицыНастроек остались только props без детей — удалить её целиком
gi = text.index('\t\t\t\t\t\t<UsualGroup name="ГруппаТаблицыНастроек"')
ge = text.index('\t\t\t\t\t\t</UsualGroup>\n', gi) + len('\t\t\t\t\t\t</UsualGroup>\n')
assert 'ГруппаТаблицыНастроек' in text[gi:ge]
anchor_point = gi  # сюда вставим 2 новых ряда
text = text[:gi] + text[ge:]


def hgroup(name, children, gid_hint):
    g, e = next(nid), next(nid)
    return (f'\t\t\t\t\t\t<UsualGroup name="{name}" id="{g}">\n'
            f'\t\t\t\t\t\t\t<Group>Horizontal</Group>\n'
            f'\t\t\t\t\t\t\t<Behavior>Usual</Behavior>\n'
            f'\t\t\t\t\t\t\t<Representation>None</Representation>\n'
            f'\t\t\t\t\t\t\t<ChildItemsWidth>Equal</ChildItemsWidth>\n'
            f'\t\t\t\t\t\t\t<ShowTitle>false</ShowTitle>\n'
            f'\t\t\t\t\t\t\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{e}"/>\n'
            f'\t\t\t\t\t\t\t<ChildItems>\n'
            + children +
            f'\t\t\t\t\t\t\t</ChildItems>\n'
            f'\t\t\t\t\t\t</UsualGroup>\n')


# Ряд 1: Склади | Рахунки (таблицы остаются на 7 табах — как раз дети горизонтальной группы 6 таба)
row1 = hgroup("ГруппаНастрПервыйРяд", t_sklady + t_scheta, None)

# Ряд 2: Рахунки малоцінки | [вертикальная группа: Етапи + чекбокс]
gv, ev = next(nid), next(nid)
# Этапы и чекбокс вкладываются на 1 таб глубже (8 табов)
vgroup = (f'\t\t\t\t\t\t\t<UsualGroup name="ГруппаЕтапиІФлаг" id="{gv}">\n'
          f'\t\t\t\t\t\t\t\t<Group>Vertical</Group>\n'
          f'\t\t\t\t\t\t\t\t<Behavior>Usual</Behavior>\n'
          f'\t\t\t\t\t\t\t\t<Representation>None</Representation>\n'
          f'\t\t\t\t\t\t\t\t<ShowTitle>false</ShowTitle>\n'
          f'\t\t\t\t\t\t\t\t<ExtendedTooltip name="ГруппаЕтапиІФлагРасширеннаяПодсказка" id="{ev}"/>\n'
          f'\t\t\t\t\t\t\t\t<ChildItems>\n'
          + reindent(t_etapy, 1)
          + reindent(cb, 2)
          + f'\t\t\t\t\t\t\t\t</ChildItems>\n'
          f'\t\t\t\t\t\t\t</UsualGroup>\n')
row2 = hgroup("ГруппаНастрВторойРяд", t_malo + vgroup, None)

text = text[:anchor_point] + row1 + row2 + text[anchor_point:]

# --- контроли ---
ids = re.findall(r'id="(-?\d+)"', text)
assert len(ids) == len(set(ids)), "дубли id"
assert text.count('<ChildItemsWidth>Equal</ChildItemsWidth>') == 2
for n in ("СкладыОстатков", "СчетаОстатков", "СчетаМалоценки", "Этапы", "ИсключатьЭтапы"):
    assert text.count(f'name="{n}"') >= 1, n
assert 'ГруппаТаблицыНастроек' not in text
data = text.replace('\n', '\r\n') if crlf else text
open(F, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + data.encode('utf-8'))
xml.dom.minidom.parse(F)
print(f"GRID 2x2 OK: id={len(ids)}, max={max(int(x) for x in ids)}, 2 ряда ChildItemsWidth=Equal, чекбокс под Етапи")
