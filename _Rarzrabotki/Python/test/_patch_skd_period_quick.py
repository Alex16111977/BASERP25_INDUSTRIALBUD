# -*- coding: utf-8 -*-
"""Добавить параметр Период как быстрый пользовательский настройку в вариант СКД."""
import re, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"

BLOCK = """
				<dcsset:dataParameters>
					<dcscor:item xsi:type="dcsset:SettingsParameterValue">
						<dcscor:parameter>Период</dcscor:parameter>
						<dcscor:value xsi:type="v8:StandardPeriod">
							<v8:variant xsi:type="v8:StandardPeriodVariant">ThisMonth</v8:variant>
						</dcscor:value>
						<dcsset:userSettingID>b00d430c-6e74-4105-9656-63c1baca039e</dcsset:userSettingID>
					</dcscor:item>
				</dcsset:dataParameters>"""

with open(PATH, "r", encoding="utf-8") as f:
    xml = f.read()

if "<dcsset:dataParameters>" in xml:
    print("dataParameters уже есть — пропуск")
else:
    new, n = re.subn(r"(<dcsset:settings[^>]*>)", r"\1" + BLOCK, xml, count=1)
    assert n == 1, f"settings-тег не найден (n={n})"
    with open(PATH, "w", encoding="utf-8") as f:
        f.write(new)
    print("OK: Период добавлен как быстрая настройка.")

# контроль
with open(PATH, "r", encoding="utf-8") as f:
    chk = f.read()
import xml.etree.ElementTree as ET
ET.fromstring(chk)
print("XML well-formed. userSettingID присутствует:", "b00d430c-6e74-4105-9656-63c1baca039e" in chk)
