# -*- coding: utf-8 -*-
"""Итоги «Аналіз залишків» — в подвал таблицы (канон ИнвентаризацияРасчетовСКонтрагентами:
Footer=true у ValueTable-таблицы + FooterDataPath колонок на реквизиты формы Итого*).
Удаляем широкую группу «Підсумки» (8 полей) — она даёт форменную горизонтальную прокрутку.
Реквизиты Итого* и их подсчёт в ПрименитьОтборИИтоги остаются (питают подвал)."""
import re, sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

F = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml"

raw = open(F, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
crlf = b'\r\n' in raw
text = raw.decode('utf-8-sig').replace('\n', '\n').replace('\r\n', '\n')

# --- 1. Footer=true у ТаблицаАнализа (после ReadOnly, перед DataPath — порядок схемы) ---
anchor = ('<Table name="ТаблицаАнализа" id="45">\n'
          '\t\t\t\t\t\t\t<CommandBarLocation>None</CommandBarLocation>\n'
          '\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n')
assert anchor in text, "не найдена таблица анализа"
text = text.replace(anchor, anchor + '\t\t\t\t\t\t\t<Footer>true</Footer>\n', 1)

# --- 2. FooterDataPath на 8 числовых колонок (перед <ContextMenu колонки) ---
MAP = {
    "ТАОстаток": "ИтогоЗалишок",
    "ТАВНорме": "ИтогоВНорме",
    "ТАПонадНорму": "ИтогоПонад",
    "ТАЭкономия": "ИтогоЭкономия",
    "ТАСуммаОстатка": "ИтогоСуммаОстатка",
    "ТАВНормеСумма": "ИтогоВНормеСума",
    "ТАПонадНормуСумма": "ИтогоПонадСума",
    "ТАЕкономіяСума": "ИтогоЕкономіяСума",
}
for col, itog in MAP.items():
    m = re.search(rf'<InputField name="{col}" id="\d+">', text)
    assert m, col
    end = text.index("</InputField>", m.start())
    # позиция ContextMenu этой колонки (внутри её InputField)
    cm = text.index("<ContextMenu ", m.start())
    assert cm < end, f"{col}: ContextMenu вне поля"
    indent = "\t" * 9
    assert "<FooterDataPath>" not in text[m.start():end], f"{col}: FooterDataPath уже есть"
    text = text[:cm] + f"<FooterDataPath>{itog}</FooterDataPath>\n{indent}" + text[cm:]
print(f"OK: Footer + {len(MAP)} FooterDataPath")

# --- 3. Удалить группу «Підсумки» целиком ---
gi = text.index('<UsualGroup name="ГруппаПідсумки"')
ls = text.rindex("\n", 0, gi) + 1
indent = text[ls:gi]                       # табы перед <UsualGroup
close = indent + "</UsualGroup>\n"
je = text.index(close, gi) + len(close)
cut = text[ls:je]
assert "ГруппаПідсумки" in cut and cut.count("<UsualGroup") == 1, "группа Підсумки содержит вложенные группы — проверить"
text = text[:ls] + text[je:]
print("OK: группа «Підсумки» удалена")

# --- контроли ---
ids = re.findall(r'id="(-?\d+)"', text)
assert len(ids) == len(set(ids)), "дубли id"
# 8 новых (аналіз) + 4 существующих (списания) = 12
assert text.count("<FooterDataPath>") == 12, text.count("<FooterDataPath>")
for itog in MAP.values():
    assert f"<FooterDataPath>{itog}</FooterDataPath>" in text, itog
assert "ГруппаПідсумки" not in text
# реквизиты Итого* должны остаться (питают подвал)
for a in MAP.values():
    assert f'<Attribute name="{a}"' in text, f"пропал реквизит {a}"
data = text.replace('\n', '\r\n') if crlf else text
open(F, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + data.encode('utf-8'))
xml.dom.minidom.parse(F)
print(f"FOOTER TOTALS OK: id={len(ids)}, реквизиты Итого* сохранены, широкий ряд удалён")
