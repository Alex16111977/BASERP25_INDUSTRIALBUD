# -*- coding: utf-8 -*-
import re, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"
with open(PATH, encoding="utf-8") as f:
    xml = f.read()
# удалить блок <field>...Документ...</field>
new, n = re.subn(
    r'\s*<field xsi:type="DataSetFieldField">\s*<dataPath>Документ</dataPath>.*?</field>',
    "", xml, count=1, flags=re.DOTALL)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(new)
import xml.etree.ElementTree as ET
ET.fromstring(new)
print(f"Удалено полей Документ: {n}. Осталось упоминаний 'Документ' в файле: {new.count('Документ')}. XML well-formed.")
