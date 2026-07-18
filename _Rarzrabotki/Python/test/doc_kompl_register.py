# -*- coding: utf-8 -*-
"""Регистрация Документ.РасчетКомплектаций: подсистема СкладГлавное (Content + CommandInterface)
и роли ДобавлениеИзменениеДанныхБухгалтерии / ЧтениеДанныхБухгалтерии (surgical insert, без RLS)."""
import sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"


def read(path):
    raw = open(path, 'rb').read()
    return raw.decode('utf-8-sig').replace('\r\n', '\n'), raw.startswith(b'\xef\xbb\xbf'), b'\r\n' in raw


def write(path, text, bom, crlf):
    if crlf:
        text = text.replace('\n', '\r\n')
    open(path, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + text.encode('utf-8'))


def insert_once(path, anchor, insertion, after=True):
    text, bom, crlf = read(path)
    if insertion.strip().splitlines()[0].strip() in text:
        print(f"SKIP (уже есть): {path.split(chr(92))[-1]}")
        return
    assert text.count(anchor) == 1, f"якорь не уникален/не найден в {path}: {anchor[:60]}"
    text = text.replace(anchor, (anchor + insertion) if after else (insertion + anchor), 1)
    write(path, text, bom, crlf)
    xml.dom.minidom.parse(path)
    print(f"OK: {path.split(chr(92))[-1]}")


# 1. Подсистема: Content
insert_once(
    CFG + r"\Subsystems\Склад\Subsystems\СкладГлавное.xml",
    '\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">Document.КомплектацияНоменклатуры</xr:Item>\n',
    '\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">Document.РасчетКомплектаций</xr:Item>\n')

# 2. Подсистема: CommandInterface
insert_once(
    CFG + r"\Subsystems\Склад\Subsystems\СкладГлавное\Ext\CommandInterface.xml",
    ('\t\t<Command name="Document.КомплектацияНоменклатуры.StandardCommand.OpenList">\n'
     '\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
     '\t\t</Command>\n'),
    ('\t\t<Command name="Document.РасчетКомплектаций.StandardCommand.OpenList">\n'
     '\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
     '\t\t</Command>\n'))

# 3. Роли (вставка перед </Rights>; без RLS — Организация в документе опциональна)
RIGHTS_FULL = '''	<object>
		<name>Document.РасчетКомплектаций</name>
		<right>
			<name>Read</name>
			<value>true</value>
		</right>
		<right>
			<name>Insert</name>
			<value>true</value>
		</right>
		<right>
			<name>Update</name>
			<value>true</value>
		</right>
		<right>
			<name>View</name>
			<value>true</value>
		</right>
		<right>
			<name>InteractiveInsert</name>
			<value>true</value>
		</right>
		<right>
			<name>Edit</name>
			<value>true</value>
		</right>
		<right>
			<name>InteractiveSetDeletionMark</name>
			<value>true</value>
		</right>
		<right>
			<name>InteractiveClearDeletionMark</name>
			<value>true</value>
		</right>
		<right>
			<name>InputByString</name>
			<value>true</value>
		</right>
	</object>
'''
RIGHTS_READ = '''	<object>
		<name>Document.РасчетКомплектаций</name>
		<right>
			<name>Read</name>
			<value>true</value>
		</right>
		<right>
			<name>View</name>
			<value>true</value>
		</right>
		<right>
			<name>InputByString</name>
			<value>true</value>
		</right>
	</object>
'''
for role, block in (("ДобавлениеИзменениеДанныхБухгалтерии", RIGHTS_FULL),
                    ("ЧтениеДанныхБухгалтерии", RIGHTS_READ)):
    path = CFG + rf"\Roles\{role}\Ext\Rights.xml"
    text, bom, crlf = read(path)
    if "Document.РасчетКомплектаций" in text:
        print(f"SKIP роль {role} (уже есть)")
        continue
    anchor = "</Rights>"
    assert text.count(anchor) == 1
    text = text.replace(anchor, block + anchor, 1)
    write(path, text, bom, crlf)
    xml.dom.minidom.parse(path)
    print(f"OK роль: {role}")

print("REGISTER PASS")
