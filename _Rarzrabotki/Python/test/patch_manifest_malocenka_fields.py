# -*- coding: utf-8 -*-
"""Добавляет в ТЧ ДокументиМалоценки колонки ФизЛицо (a4..009) и СпособОтраженияРасходов (a4..010).
Клонирует атрибут Склад (a4..006). Идемпотентно."""
import re, sys, xml.dom.minidom as minidom
if sys.stdout.encoding != 'utf-8': sys.stdout.reconfigure(encoding='utf-8')
P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета.xml"
raw = open(P, 'rb').read(); bom = raw.startswith(b'\xef\xbb\xbf'); txt = raw.decode('utf-8-sig')

if 'a4000000-0000-4000-8000-000000000009' in txt:
    print("SKIP: колонки уже есть"); sys.exit(0)

# блок ТЧ ДокументиМалоценки
m = re.search(r'<TabularSection uuid="a4000000-0000-4000-8000-000000000001">.*?</TabularSection>', txt, re.DOTALL)
block = m.group(0)
# шаблон — атрибут Склад (a4..006)
ma = re.search(r'([\t ]*)<Attribute uuid="a4000000-0000-4000-8000-000000000006">.*?</Attribute>', block, re.DOTALL)
sklad = ma.group(0)

def clone_attr(uuid_suffix, name, synonym, cfg_type):
    a = sklad
    a = a.replace('a4000000-0000-4000-8000-000000000006', f'a4000000-0000-4000-8000-00000000000{uuid_suffix}')
    a = re.sub(r'<Name>Склад</Name>', f'<Name>{name}</Name>', a, count=1)
    a = re.sub(r'<v8:content>Склад</v8:content>', f'<v8:content>{synonym}</v8:content>', a, count=1)
    a = re.sub(r'<v8:Type>cfg:CatalogRef\.Склады</v8:Type>', f'<v8:Type>cfg:{cfg_type}</v8:Type>', a, count=1)
    return a

физлицо = clone_attr('9', 'ФизЛицо', 'МВО (фізична особа)', 'CatalogRef.ФизическиеЛица')
способ  = clone_attr('10'.replace('10','A') if False else '10', 'СпособОтраженияРасходов',
                     'Спосіб відображення витрат', 'CatalogRef.СпособыОтраженияРасходовПоАмортизации')
# uuid ...010 (не ...0010): суффикс из двух цифр
способ = способ.replace('a4000000-0000-4000-8000-0000000000010', 'a4000000-0000-4000-8000-000000000010')

new_block = block.replace('</Attribute>\n\t\t\t\t\t</ChildObjects>',
                          '</Attribute>\n' + физлицо + '\n' + способ + '\n\t\t\t\t\t</ChildObjects>', 1)
if new_block == block:  # fallback если отступ иной
    idx = block.rfind('</Attribute>') + len('</Attribute>')
    new_block = block[:idx] + '\n' + физлицо + '\n' + способ + block[idx:]

txt = txt.replace(block, new_block, 1)
data = txt.encode('utf-8');  data = (b'\xef\xbb\xbf' + data) if bom else data
open(P, 'wb').write(data)

minidom.parseString(open(P, encoding='utf-8-sig').read())
t2 = open(P, encoding='utf-8-sig').read()
uuids = re.findall(r'uuid="([0-9a-fA-F-]+)"', t2); dups = {u for u in uuids if uuids.count(u) > 1}
print("PARSE OK; BOM:", bom)
print("Дубли uuid:", dups if dups else "нет")
for n in ("ФизЛицо", "СпособОтраженияРасходов"):
    print(f"колонка {n}:", "OK" if f"<Name>{n}</Name>" in t2 else "НЕТ")
print("a4..009:", 'a4000000-0000-4000-8000-000000000009' in t2, "| a4..010:", 'a4000000-0000-4000-8000-000000000010' in t2)
print("RESULT:", "PASS" if not dups and '<Name>ФизЛицо</Name>' in t2 and '<Name>СпособОтраженияРасходов</Name>' in t2 else "FAIL")
