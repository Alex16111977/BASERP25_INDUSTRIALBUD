# -*- coding: utf-8 -*-
import xml.dom.minidom as minidom, re, sys
if sys.stdout.encoding != 'utf-8': sys.stdout.reconfigure(encoding='utf-8')
p = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета.xml"
doc = minidom.parse(p)
print("PARSE OK")
txt = open(p, encoding='utf-8').read()
# уникальность всех uuid
uuids = re.findall(r'uuid="([0-9a-fA-F-]+)"', txt)
dups = {u for u in uuids if uuids.count(u) > 1}
print("Дубли uuid:", dups if dups else "нет")
# новые ТЧ присутствуют
for name in ["СчетаМалоценки", "ДокументиМалоценки"]:
    print(f"ТЧ {name}:", "OK" if f"<Name>{name}</Name>" in txt else "НЕ НАЙДЕНА")
# все TabularSection
ts = re.findall(r'<Name>([^<]+)</Name>\s*<Synonym>', txt)
print("Всего <TabularSection uuid> :", txt.count("<TabularSection uuid="))
print("RESULT:", "PASS" if (not dups and "<Name>СчетаМалоценки</Name>" in txt and "<Name>ДокументиМалоценки</Name>" in txt) else "FAIL")
