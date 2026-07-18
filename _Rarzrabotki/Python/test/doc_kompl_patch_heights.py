# -*- coding: utf-8 -*-
"""Высоты больших таблиц РасчетКомплектаций: HeightInTableRows=5 (база) + VerticalStretch=true —
таблица растягивается ТОЛЬКО на свободное место, итоги/подвал видны без прокрутки формы.
Порядок свойств по эталону конфигурации: Table -> ReadOnly? -> [Width] -> HeightInTableRows ->
VerticalStretch -> DataPath (DataProcessors/НачалоРаботы)."""
import re, sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

F = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml"

raw = open(F, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
crlf = b'\r\n' in raw
text = raw.decode('utf-8-sig').replace('\r\n', '\n')

for name in ("ТаблицаАнализа", "СписаниеПоНормам", "СписаниеСверхНормы"):
    m = re.search(rf'([\t]+)<Table name="{name}" id="\d+">\n', text)
    assert m, f"нет таблицы {name}"
    indent = m.group(1) + "\t"
    pos = m.end()
    # если сразу идёт ReadOnly — вставляем после него
    m_ro = re.match(rf'{re.escape(indent)}<ReadOnly>true</ReadOnly>\n', text[pos:])
    if m_ro:
        pos += m_ro.end()
    seg = text[pos:pos + 400]
    assert "<HeightInTableRows>" not in seg.split("</Table>")[0].split("<ChildItems>")[0], f"{name}: высота уже задана"
    ins = (f"{indent}<HeightInTableRows>5</HeightInTableRows>\n"
           f"{indent}<VerticalStretch>true</VerticalStretch>\n")
    text = text[:pos] + ins + text[pos:]
    print(f"OK {name}: 5 строк + VerticalStretch")

data = text.replace('\n', '\r\n') if crlf else text
open(F, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + data.encode('utf-8'))
xml.dom.minidom.parse(F)
print("HEIGHTS PATCH PASS")
