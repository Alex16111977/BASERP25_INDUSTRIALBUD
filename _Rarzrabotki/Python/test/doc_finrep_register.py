# -*- coding: utf-8 -*-
"""Регистрация Документ.А_ФинансовыйОтчетПроизводства:
подсистема Продажи/РасчетыСКонтрагентами (Content + CommandInterface) и роли
ДобавлениеИзменениеДанныхБухгалтерии / ЧтениеДанныхБухгалтерии — без RLS
(Организация в документе опциональна, RLS по организации сломал бы пустую орг).
Идемпотентно: повторный запуск ничего не дублирует."""
import sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
ДОК = "Document.А_ФинансовыйОтчетПроизводства"


def читать(путь):
    сырое = open(путь, 'rb').read()
    return (сырое.decode('utf-8-sig').replace('\r\n', '\n'),
            сырое.startswith(b'\xef\xbb\xbf'), b'\r\n' in сырое)


def писать(путь, текст, bom, crlf):
    if crlf:
        текст = текст.replace('\n', '\r\n')
    open(путь, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + текст.encode('utf-8'))


def вставить(путь, якорь, вставка):
    текст, bom, crlf = читать(путь)
    if ДОК in текст:
        print(f"SKIP (уже есть): {путь.split(chr(92))[-1]}")
        return
    assert текст.count(якорь) == 1, f"якорь не уникален/не найден в {путь}"
    текст = текст.replace(якорь, якорь + вставка, 1)
    писать(путь, текст, bom, crlf)
    xml.dom.minidom.parse(путь)
    print(f"OK: {путь.split(chr(92))[-1]}")


# 1. Подсистема — Content
вставить(
    CFG + r"\Subsystems\Продажи\Subsystems\РасчетыСКонтрагентами.xml",
    '\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">Document.ИмпРеализацияСтроительныхРаботУслуг</xr:Item>\n',
    f'\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">{ДОК}</xr:Item>\n')

# 2. Подсистема — CommandInterface (порядок в панели навигации)
вставить(
    CFG + r"\Subsystems\Продажи\Subsystems\РасчетыСКонтрагентами\Ext\CommandInterface.xml",
    ('\t\t<Command name="Document.КорректировкаДолга.StandardCommand.OpenList">\n'
     '\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
     '\t\t</Command>\n'),
    (f'\t\t<Command name="{ДОК}.StandardCommand.OpenList">\n'
     '\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
     '\t\t</Command>\n'))

# 3. Роли
ПОЛНЫЕ = """	<object>
		<name>{doc}</name>
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
			<name>Delete</name>
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
			<name>InteractiveDelete</name>
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
""".format(doc=ДОК)

ЧТЕНИЕ = """	<object>
		<name>{doc}</name>
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
""".format(doc=ДОК)

for роль, блок in (("ДобавлениеИзменениеДанныхБухгалтерии", ПОЛНЫЕ),
                   ("ЧтениеДанныхБухгалтерии", ЧТЕНИЕ)):
    путь = CFG + rf"\Roles\{роль}\Ext\Rights.xml"
    текст, bom, crlf = читать(путь)
    if ДОК in текст:
        print(f"SKIP роль {роль} (уже есть)")
        continue
    якорь = "</Rights>"
    assert текст.count(якорь) == 1
    текст = текст.replace(якорь, блок + якорь, 1)
    писать(путь, текст, bom, crlf)
    xml.dom.minidom.parse(путь)
    print(f"OK роль: {роль}")

print("REGISTER PASS")
