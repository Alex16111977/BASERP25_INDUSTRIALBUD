# -*- coding: utf-8 -*-
"""Task 7: клонирование таблиц формы для СчетаМалоценки/ДокументиМалоценки. Идемпотентно (skip если уже есть)."""
import re, sys, xml.dom.minidom as minidom
if sys.stdout.encoding != 'utf-8': sys.stdout.reconfigure(encoding='utf-8')

P = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Forms\Форма\Ext\Form.xml"
raw = open(P, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
txt = raw.decode('utf-8-sig')

if 'name="СчетаМалоценки"' in txt or 'name="ДокументиМалоценки"' in txt:
    print("SKIP: клоны уже присутствуют"); sys.exit(0)

existing_ids = set(int(x) for x in re.findall(r'id="(\d+)"', txt))
next_id = [max(existing_ids) + 1]  # 283

def remap_ids(block):
    def repl(m):
        v = next_id[0]; next_id[0] += 1
        return f'id="{v}"'
    return re.sub(r'id="\d+"', repl, block)

def clone_table(txt, src_name, dst_name, extra_renames, new_title):
    m = re.search(r'([\t ]*)<Table name="' + re.escape(src_name) + r'".*?</Table>', txt, re.DOTALL)
    if not m:
        raise SystemExit(f"FAIL: не найден блок таблицы {src_name}")
    block = m.group(0)
    clone = block
    clone = clone.replace(src_name, dst_name)           # name/DataPath/Item/sub-names
    for a, b in extra_renames:
        clone = clone.replace(a, b)
    clone = remap_ids(clone)                              # sequential id remap
    # заменить/вставить <Title> клона
    title_xml = ('<Title>\n\t\t\t\t\t\t\t\t<v8:item>\n\t\t\t\t\t\t\t\t\t<v8:lang>uk</v8:lang>\n'
                 '\t\t\t\t\t\t\t\t\t<v8:content>' + new_title + '</v8:content>\n'
                 '\t\t\t\t\t\t\t\t</v8:item>\n\t\t\t\t\t\t\t</Title>')
    if '<Title>' in clone:
        clone = re.sub(r'<Title>.*?</Title>', title_xml, clone, count=1, flags=re.DOTALL)
    else:
        clone = re.sub(r'(<DataPath>[^<]+</DataPath>\n)', r'\1\t\t\t\t\t\t\t' + title_xml + '\n', clone, count=1)
    return txt.replace(block, block + '\n' + clone, 1)

# 1) СчетаОстатков -> СчетаМалоценки (после DataPath у СчетаОстатков нет проблем с ДКК-префиксом)
txt = clone_table(txt, "СчетаОстатков", "СчетаМалоценки", [], "Рахунки малоцінки (порожньо = 22)")
# 2) ДокументиКомплектації -> ДокументиМалоценки (+ префикс ДКК -> ДМ у InputField'ов)
txt = clone_table(txt, "ДокументиКомплектації", "ДокументиМалоценки", [("ДКК", "ДМ")],
                  "Документи малоцінки (ПередачаМалоценныхАктивовВЭксплуатацию)")
# 3) обновить заголовок СчетаОстатков 20/22/28 -> 20/28
txt = txt.replace("Рахунки залишків (порожньо = 20/22/28)", "Рахунки залишків для комплектації (порожньо = 20/28)")

# запись (сохранить признак BOM оригинала)
data = txt.encode('utf-8')
if bom: data = b'\xef\xbb\xbf' + data
open(P, 'wb').write(data)

# валидация
minidom.parseString(open(P, encoding='utf-8-sig').read())
ids = re.findall(r'id="(\d+)"', open(P, encoding='utf-8-sig').read())
dups = {i for i in ids if ids.count(i) > 1}
print("PARSE OK; BOM:", bom)
print("Дубли id:", dups if dups else "нет")
print("Новые таблицы:", 'name="СчетаМалоценки"' in txt, 'name="ДокументиМалоценки"' in txt)
print("Диапазон новых id: 283 ..", next_id[0]-1)
print("RESULT:", "PASS" if not dups else "FAIL")
