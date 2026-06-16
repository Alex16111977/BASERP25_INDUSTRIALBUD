# -*- coding: utf-8 -*-
import xml.etree.ElementTree as ET
import sys
sys.stdout.reconfigure(encoding='utf-8')

base = r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость'
files = [
    base + '.xml',
    base + r'\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml',
    base + r'\Forms\ФормаОтчетаУправляемая\Ext\Form.xml',
]
ok = True
for p in files:
    try:
        ET.parse(p)
        print('WELL-FORMED:', p.split('\\')[-1])
    except Exception as e:
        ok = False
        print('MALFORMED :', p, '->', e)
print('ИТОГ:', 'ВСЕ OK' if ok else 'ЕСТЬ ОШИБКИ')
