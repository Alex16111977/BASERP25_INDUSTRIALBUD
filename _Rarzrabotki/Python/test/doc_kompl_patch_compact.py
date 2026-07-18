# -*- coding: utf-8 -*-
"""Ещё ~80px высоты формы РасчетКомплектаций (ноутбук, «не видно итоги»):
1) ТаблицаАнализа: убрать командную панель (read-only — «Добавить» всегда серые) и пустой Footer;
2) убрать строку-заголовок группы «Відбір» (подписи полей остаются)."""
import sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

F = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml"

raw = open(F, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
crlf = b'\r\n' in raw
text = raw.decode('utf-8-sig').replace('\r\n', '\n')

# 1a. CommandBarLocation=None ПЕРЕД ReadOnly (порядок схемы: Representation, CommandBarLocation, ReadOnly)
anchor = ('\t\t\t\t\t\t<Table name="ТаблицаАнализа" id="45">\n'
          '\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n')
assert anchor in text, "не найден узел ТаблицаАнализа"
text = text.replace(anchor,
                    '\t\t\t\t\t\t<Table name="ТаблицаАнализа" id="45">\n'
                    '\t\t\t\t\t\t\t<CommandBarLocation>None</CommandBarLocation>\n'
                    '\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n', 1)
print("OK: командная панель таблицы анализа скрыта")

# 1b. Пустой подвал (Footer без FooterDataPath — итоги живут в группе «Підсумки»)
footer = '\t\t\t\t\t\t\t<Footer>true</Footer>\n'
assert text.count(footer) == 1
text = text.replace(footer, '', 1)
print("OK: пустой Footer убран")

# 2. Заголовок группы «Відбір»
hdr = text.index('<UsualGroup name="ГруппаВідбір"')
t_start = text.index('\t\t\t\t\t\t\t<Title>', hdr)
t_end = text.index('\t\t\t\t\t\t\t</Title>\n', t_start) + len('\t\t\t\t\t\t\t</Title>\n')
seg = text[t_start:t_end]
assert "Відбір" in seg, "Title группы Відбір не совпал"
text = text[:t_start] + text[t_end:]
print("OK: заголовок «Відбір» убран")

data = text.replace('\n', '\r\n') if crlf else text
open(F, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + data.encode('utf-8'))
xml.dom.minidom.parse(F)
print("COMPACT PATCH PASS")
