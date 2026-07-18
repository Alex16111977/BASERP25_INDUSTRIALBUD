# -*- coding: utf-8 -*-
"""Порт 4 СКД-макетов обработки в Документ.РасчетКомплектаций:
копия Ext/Template.xml как есть (uk-титулы/appearance вшиты), обёртки с НОВЫМИ uuid4 (§39),
регистрация <Template> в xml документа, контроль уникальности uuid по ВСЕЙ конфигурации."""
import os, shutil, sys, uuid, glob, re
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

OBR = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета"
CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
DOCXML = CFG + r"\Documents\РасчетКомплектаций.xml"
TPLDIR = CFG + r"\Documents\РасчетКомплектаций\Templates"

MAKETY = [
    ("МакетПланФакт", "Макет план-факт"),
    ("МакетПланФактЕтапи", "Макет план-факт (за етапами)"),
    ("МакетАнализСС", "Макет аналіз СС"),
    ("МакетАнализССОдна", "Макет аналіз СС (одна одиниця)"),
]

NSDECL = ('<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
          'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" '
          'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
          'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
          'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
          'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
          'xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" '
          'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
          'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">')

WRAPPER = '''<?xml version="1.0" encoding="UTF-8"?>
{ns}
	<Template uuid="{uid}">
		<Properties>
			<Name>{name}</Name>
			<Synonym>
				<v8:item>
					<v8:lang>uk</v8:lang>
					<v8:content>{syn}</v8:content>
				</v8:item>
			</Synonym>
			<Comment/>
			<TemplateType>DataCompositionSchema</TemplateType>
		</Properties>
	</Template>
</MetaDataObject>
'''

os.makedirs(TPLDIR, exist_ok=True)
new_uids = []
for name, syn in MAKETY:
    src_body = os.path.join(OBR, "Templates", name, "Ext", "Template.xml")
    dst_body_dir = os.path.join(TPLDIR, name, "Ext")
    os.makedirs(dst_body_dir, exist_ok=True)
    shutil.copyfile(src_body, os.path.join(dst_body_dir, "Template.xml"))
    uid = str(uuid.uuid4())
    new_uids.append(uid)
    wrapper = WRAPPER.format(ns=NSDECL, uid=uid, name=name, syn=syn)
    open(os.path.join(TPLDIR, name + ".xml"), 'wb').write(
        b'\xef\xbb\xbf' + wrapper.encode('utf-8'))
    print(f"OK {name}: uuid={uid}")

# Регистрация в xml документа: 4 строки <Template> перед закрывающим </ChildObjects> документа
raw = open(DOCXML, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
crlf = b'\r\n' in raw
text = raw.decode('utf-8-sig').replace('\r\n', '\n')
assert "<Template>" not in text, "Template уже зарегистрированы?"
anchor = "\t\t</ChildObjects>\n\t</Document>"
assert text.count(anchor) == 1, "не найден якорь конца ChildObjects документа"
lines = "".join(f"\t\t\t<Template>{name}</Template>\n" for name, _ in MAKETY)
text = text.replace(anchor, lines + anchor)
if crlf:
    text = text.replace('\n', '\r\n')
data = text.encode('utf-8')
open(DOCXML, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + data)
xml.dom.minidom.parse(DOCXML)
print("OK регистрация <Template> x4")

# Контроль §39: уникальность Template uuid по ВСЕЙ конфигурации
all_uids = {}
dups = []
for path in glob.glob(CFG + r"\**\Templates\*.xml", recursive=True) + glob.glob(CFG + r"\CommonTemplates\*.xml"):
    try:
        head = open(path, 'rb').read(2000).decode('utf-8-sig', errors='ignore')
    except OSError:
        continue
    m = re.search(r'<Template uuid="([0-9a-f-]{36})"', head)
    if m:
        u = m.group(1)
        if u in all_uids:
            dups.append((u, all_uids[u], path))
        all_uids[u] = path
assert not dups, f"ДУБЛИ uuid макетов: {dups}"
for u in new_uids:
    assert u in all_uids
# parse обёрток
for name, _ in MAKETY:
    xml.dom.minidom.parse(os.path.join(TPLDIR, name + ".xml"))
    xml.dom.minidom.parse(os.path.join(TPLDIR, name, "Ext", "Template.xml"))
print(f"UUID-контроль: {len(all_uids)} макетов в конфигурации, дублей нет")
print("TEMPLATES PASS")
