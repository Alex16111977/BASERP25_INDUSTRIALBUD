# -*- coding: utf-8 -*-
"""Стандартные кнопки документа («Записать и закрыть»/«Записать») в форме РасчетКомплектаций —
как у КомплектацияНоменклатуры: ButtonGroup с CommandSource=Form в AutoCommandBar
+ перенос наших кнопок (Розрахувати / Заповнити / Друк) в верхнюю панель,
нижняя CommandBar «ОсновнаяКоманднаяПанель» удаляется.
Патч ТЕКУЩЕЙ (Designer-нормализованной) формы — НЕ перегенерация."""
import re, sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

F = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml"

raw = open(F, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
crlf = b'\r\n' in raw
text = raw.decode('utf-8-sig').replace('\r\n', '\n')

max_id = max(int(m) for m in re.findall(r'id="(-?\d+)"', text))
nid = iter(range(max_id + 1, max_id + 50))

# --- 1. Вырезать CommandBar ОсновнаяКоманднаяПанель, забрать его кнопки ---
cb_start = text.index('\t\t<CommandBar name="ОсновнаяКоманднаяПанель" id="1">')
cb_end = text.index('\t\t</CommandBar>\n', cb_start) + len('\t\t</CommandBar>\n')
cb_node = text[cb_start:cb_end]
ci = cb_node.index('<ChildItems>') + len('<ChildItems>\n')
cj = cb_node.rindex('\t\t\t</ChildItems>')
buttons = cb_node[ci:cj]
# кнопки были на 4 табах (CommandBar 2 -> ChildItems 3 -> кнопки 4);
# в AutoCommandBar (1 таб -> ChildItems 2) им нужно 3 таба: убрать один \t
buttons = "".join((line[1:] if line.startswith('\t') else line)
                  for line in buttons.splitlines(keepends=True))
text = text[:cb_start] + text[cb_end:]

# --- 2. Пересобрать AutoCommandBar ---
acb_old = ('\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n'
           '\t\t<HorizontalAlign>Right</HorizontalAlign>\n'
           '\t\t<Autofill>false</Autofill>\n'
           '\t</AutoCommandBar>\n')
assert acb_old in text, "не найден узел AutoCommandBar"
g1, g2, g3, g4 = next(nid), next(nid), next(nid), next(nid)
acb_new = f'''\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">
\t\t<Autofill>false</Autofill>
\t\t<ChildItems>
\t\t\t<ButtonGroup name="ГруппаКнопкиКоманднойПанели" id="{g1}">
\t\t\t\t<Title>
\t\t\t\t\t<v8:item>
\t\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t\t<v8:content>Кнопки командной панели</v8:content>
\t\t\t\t\t</v8:item>
\t\t\t\t\t<v8:item>
\t\t\t\t\t\t<v8:lang>uk</v8:lang>
\t\t\t\t\t\t<v8:content>Кнопки командної панелі</v8:content>
\t\t\t\t\t</v8:item>
\t\t\t\t</Title>
\t\t\t\t<CommandSource>Form</CommandSource>
\t\t\t\t<ExtendedTooltip name="ГруппаКнопкиКоманднойПанелиРасширеннаяПодсказка" id="{g2}"/>
\t\t\t</ButtonGroup>
{buttons}\t\t\t<Button name="ФормаСправка" id="{g3}">
\t\t\t\t<Type>CommandBarButton</Type>
\t\t\t\t<CommandName>Form.StandardCommand.Help</CommandName>
\t\t\t\t<LocationInCommandBar>InCommandBarAndInAdditionalSubmenu</LocationInCommandBar>
\t\t\t\t<ExtendedTooltip name="ФормаСправкаРасширеннаяПодсказка" id="{g4}"/>
\t\t\t</Button>
\t\t</ChildItems>
\t</AutoCommandBar>
'''
text = text.replace(acb_old, acb_new, 1)

# --- 3. Высоты (ноутбук): большие таблицы — авто-растяжение со скроллом внутри,
#     мелкие настроечные — 3 строки ---
n14 = text.count("<HeightInTableRows>14</HeightInTableRows>\n")
n12 = text.count("<HeightInTableRows>12</HeightInTableRows>\n")
assert n14 == 1 and n12 == 2, f"высоты: 14x{n14}, 12x{n12}"
text = re.sub(r"[\t]*<HeightInTableRows>1[24]</HeightInTableRows>\n", "", text)
text = text.replace("<HeightInTableRows>6</HeightInTableRows>", "<HeightInTableRows>3</HeightInTableRows>")
text = text.replace("<HeightInTableRows>4</HeightInTableRows>", "<HeightInTableRows>3</HeightInTableRows>")

# --- 4. Убрать заголовок «Документ» у группы шапки (съедает строку) ---
hdr = text.index('<UsualGroup name="ГруппаШапка"')
t_start = text.index('\t\t\t<Title>', hdr)
t_end = text.index('\t\t\t</Title>\n', t_start) + len('\t\t\t</Title>\n')
text = text[:t_start] + text[t_end:]

# --- Контроли ---
ids = re.findall(r'id="(-?\d+)"', text)
assert len(ids) == len(set(ids)), "дубли id!"
assert 'CommandBar name="ОсновнаяКоманднаяПанель"' not in text
assert text.count('Form.Command.Рассчитать') == 1
assert text.count('<CommandSource>Form</CommandSource>') == 1
assert "<HeightInTableRows>14" not in text and "<HeightInTableRows>12" not in text
data = text.replace('\n', '\r\n') if crlf else text
open(F, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + data.encode('utf-8'))
xml.dom.minidom.parse(F)
print(f"PATCH OK: id уникальны ({len(ids)}), max id={max(int(x) for x in ids)}, кнопки в AutoCommandBar, CommandBar удалена")
