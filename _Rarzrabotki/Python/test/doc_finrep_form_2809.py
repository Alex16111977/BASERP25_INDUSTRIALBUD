# -*- coding: utf-8 -*-
"""Form.xml фінзвіту: колонка «Підрозділ казни» у таблиці розшифровки (після «Сума Кт»).
Ідемпотентний; id — max+1..+3; перевірка minidom і assert після запису."""
import re, io, sys
from xml.dom import minidom
sys.stdout.reconfigure(encoding="utf-8")
П = sys.argv[1] if len(sys.argv) > 1 else (
    r"C:\Configuration_downloads\BASERP25\.claude\worktrees\director-report-letter-689b66\_Rarzrabotki\BASEBuh"
    r"\Documents\А_ФинансовыйОтчетПроизводства\Forms\ФормаДокумента\Ext\Form.xml")
т = io.open(П, encoding="utf-8-sig", newline="").read()
пс = "\r\n" if "\r\n" in т else "\n"
if 'name="РасшПодразделениеКазны"' not in т:
    мх = max(int(x) for x in re.findall(r'\bid="(\d+)"', т))
    в8, в9 = "\t" * 8, "\t" * 9
    якір = f'{в9}<ExtendedTooltip name="РасшСуммаКтРасширеннаяПодсказка" id="235"/>{пс}{в8}</InputField>'
    assert т.count(якір) == 1, "якір РасшСуммаКт"
    нове = (якір + пс
            + f'{в8}<InputField name="РасшПодразделениеКазны" id="{мх + 1}">{пс}'
            + f'{в9}<DataPath>Объект.РасшифровкаДокументов.ПодразделениеКазны</DataPath>{пс}'
            + f'{в9}<ContextMenu name="РасшПодразделениеКазныКонтекстноеМеню" id="{мх + 2}"/>{пс}'
            + f'{в9}<ExtendedTooltip name="РасшПодразделениеКазныРасширеннаяПодсказка" id="{мх + 3}"/>{пс}'
            + f'{в8}</InputField>')
    т = т.replace(якір, нове, 1)
    io.open(П, "w", encoding="utf-8-sig", newline="").write(т)
minidom.parse(П)
т = io.open(П, encoding="utf-8-sig", newline="").read()
assert т.count('name="РасшПодразделениеКазны"') == 1
# id у Form.xml мають окремі простори (елементи / реквізити / команди) — повтори між ними штатні.
# Нові id = max+1..+3 по ВСЬОМУ файлу, тож кожен має зустрічатися рівно один раз.
блок = re.search(r'<InputField name="РасшПодразделениеКазны"[\s\S]*?</InputField>', т).group(0)
for нов in re.findall(r'\bid="(\d+)"', блок):
    assert т.count(f'id="{нов}"') == 1, f"id {нов} не унікальний"
print("OK Form.xml: колонка РасшПодразделениеКазны; нові id:", re.findall(r'\bid="(\d+)"', блок))
