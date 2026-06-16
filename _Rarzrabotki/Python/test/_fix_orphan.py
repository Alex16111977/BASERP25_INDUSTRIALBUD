# -*- coding: utf-8 -*-
import re, sys
import xml.etree.ElementTree as ET
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"
with open(PATH, encoding="utf-8") as f:
    xml = f.read()
# Удалить осиротевший блок <title>...Документ...</title></field> сразу после <name>НаборДанных1</name>
new, n = re.subn(
    r'(<name>НаборДанных1</name>)\s*<title xsi:type="v8:LocalStringType">.*?</title>\s*</field>',
    r'\1', xml, count=1, flags=re.DOTALL)
assert n == 1, f"осиротевший блок не найден (n={n})"
# Проверка ДО записи
ET.fromstring(new)
# поле Документ должно уйти; "Документ.ПриобретениеТоваровУслуг" в запросе — это норм
assert "<dataPath>Документ</dataPath>" not in new, "поле Документ ещё есть"
assert "<v8:content>Документ</v8:content>" not in new, "title Документ ещё есть"
with open(PATH, "w", encoding="utf-8") as f:
    f.write(new)
print("OK: осиротевший блок удалён, XML well-formed, поля Документ нет (Документ.ПТУ в запросе — норма).")
