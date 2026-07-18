# -*- coding: utf-8 -*-
"""Добавляет в таблицу формы ДокументиМалоценки поля ФизЛицо и СпособОтраженияРасходов
сразу после Склада. Клонирует InputField ДМСклад. Идемпотентно."""
import re, sys, xml.dom.minidom as minidom
if sys.stdout.encoding != 'utf-8': sys.stdout.reconfigure(encoding='utf-8')
P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Forms\Форма\Ext\Form.xml"
raw = open(P, 'rb').read(); bom = raw.startswith(b'\xef\xbb\xbf'); txt = raw.decode('utf-8-sig')

if 'name="ДМФизЛицо"' in txt:
    print("SKIP: поля уже есть"); sys.exit(0)

ids = [int(x) for x in re.findall(r'id="(\d+)"', txt)]
nid = [max(ids) + 1]
def nxt():
    v = nid[0]; nid[0] += 1; return v

m = re.search(r'([\t ]*)<InputField name="ДМСклад".*?</InputField>', txt, re.DOTALL)
indent = m.group(1); sklad_block = m.group(0)

def field(fname, datapath):
    f_id, cm_id, tt_id = nxt(), nxt(), nxt()
    return (f'{indent}<InputField name="{fname}" id="{f_id}">\n'
            f'{indent}\t<DataPath>{datapath}</DataPath>\n'
            f'{indent}\t<ContextMenu name="{fname}КонтекстноеМеню" id="{cm_id}"/>\n'
            f'{indent}\t<ExtendedTooltip name="{fname}РасширеннаяПодсказка" id="{tt_id}"/>\n'
            f'{indent}</InputField>')

физлицо = field("ДМФизЛицо", "Объект.ДокументиМалоценки.ФизЛицо")
способ  = field("ДМСпособОтраженияРасходов", "Объект.ДокументиМалоценки.СпособОтраженияРасходов")
txt = txt.replace(sklad_block, sklad_block + "\n" + физлицо + "\n" + способ, 1)

data = txt.encode('utf-8'); data = (b'\xef\xbb\xbf' + data) if bom else data
open(P, 'wb').write(data)

minidom.parseString(open(P, encoding='utf-8-sig').read())
t2 = open(P, encoding='utf-8-sig').read()
allids = re.findall(r'id="(\d+)"', t2); dups = {i for i in allids if allids.count(i) > 1}
print("PARSE OK; BOM:", bom, "| дубли id:", dups if dups else "нет")
print("поля:", 'name="ДМФизЛицо"' in t2, 'name="ДМСпособОтраженияРасходов"' in t2)
print("новые id:", max(ids)+1, "..", nid[0]-1)
print("RESULT:", "PASS" if not dups and 'name="ДМФизЛицо"' in t2 else "FAIL")
