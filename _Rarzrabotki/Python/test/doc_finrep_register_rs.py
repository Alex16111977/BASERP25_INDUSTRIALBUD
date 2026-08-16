# -*- coding: utf-8 -*-
"""Регистрация РегистрСведений.А_ПараметрыДоговоровФинотчета:
подсистема Продажи/РасчетыСКонтрагентами (Content + CommandInterface) и роли
ДобавлениеИзменениеДанныхБухгалтерии / ЧтениеДанныхБухгалтерии — без RLS
(по прецеденту документа А_ФинансовыйОтчетПроизводства).
Идемпотентно: повторный запуск ничего не дублирует."""
import sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
РЕГ = "InformationRegister.А_ПараметрыДоговоровФинотчета"
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
    if РЕГ in текст:
        print(f"SKIP (уже есть): {путь.split(chr(92))[-1]}")
        return
    assert текст.count(якорь) == 1, f"якорь не уникален/не найден в {путь}"
    текст = текст.replace(якорь, якорь + вставка, 1)
    писать(путь, текст, bom, crlf)
    xml.dom.minidom.parse(путь)
    print(f"OK: {путь.split(chr(92))[-1]}")


# 1. Подсистема — Content (сразу после документа финотчёта)
вставить(
    CFG + r"\Subsystems\Продажи\Subsystems\РасчетыСКонтрагентами.xml",
    f'\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">{ДОК}</xr:Item>\n',
    f'\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">{РЕГ}</xr:Item>\n')

# 2. Подсистема — CommandInterface (после команды списка документа)
вставить(
    CFG + r"\Subsystems\Продажи\Subsystems\РасчетыСКонтрагентами\Ext\CommandInterface.xml",
    (f'\t\t<Command name="{ДОК}.StandardCommand.OpenList">\n'
     '\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
     '\t\t</Command>\n'),
    (f'\t\t<Command name="{РЕГ}.StandardCommand.OpenList">\n'
     '\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
     '\t\t</Command>\n'))

# 3. Роли
ПОЛНЫЕ = """	<object>
		<name>{reg}</name>
		<right>
			<name>Read</name>
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
			<name>Edit</name>
			<value>true</value>
		</right>
	</object>
""".format(reg=РЕГ)

ЧТЕНИЕ = """	<object>
		<name>{reg}</name>
		<right>
			<name>Read</name>
			<value>true</value>
		</right>
		<right>
			<name>View</name>
			<value>true</value>
		</right>
	</object>
""".format(reg=РЕГ)

for роль, блок in (("ДобавлениеИзменениеДанныхБухгалтерии", ПОЛНЫЕ),
                   ("ЧтениеДанныхБухгалтерии", ЧТЕНИЕ)):
    путь = CFG + rf"\Roles\{роль}\Ext\Rights.xml"
    текст, bom, crlf = читать(путь)
    if РЕГ in текст:
        print(f"SKIP роль {роль} (уже есть)")
        continue
    якорь = "</Rights>"
    assert текст.count(якорь) == 1
    текст = текст.replace(якорь, блок + якорь, 1)
    писать(путь, текст, bom, crlf)
    xml.dom.minidom.parse(путь)
    print(f"OK роль: {роль}")

print("REGISTER RS PASS")
