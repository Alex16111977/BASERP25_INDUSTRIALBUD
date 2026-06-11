# -*- coding: utf-8 -*-
"""Генератор корневого XML внешней обработки А_НачальнаяЗадолженностьПоЗарплатеСозданнаяПоВыплатам.
Формат блоков 1:1 как у донора А_НачальнаяЗадолженностьПоЗарплатеФорма2.xml (версия 2.13).
Пишет UTF-8 BOM + CRLF."""
import io
import os

NAME = "А_НачальнаяЗадолженностьПоЗарплатеСозданнаяПоВыплатам"
OUT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "..", "Обработки", NAME + ".xml"))

HEADER = '''<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">
\t<ExternalDataProcessor uuid="7e4f2a91-c358-4b06-9d17-83a5e0f6b249">
\t\t<InternalInfo>
\t\t\t<xr:ContainedObject>
\t\t\t\t<xr:ClassId>c3831ec8-d8d5-4f93-8a22-f9bfae07327f</xr:ClassId>
\t\t\t\t<xr:ObjectId>2b67d1f4-8e95-4c30-a7d2-50c918e3f6a7</xr:ObjectId>
\t\t\t</xr:ContainedObject>
\t\t\t<xr:GeneratedType name="ExternalDataProcessorObject.{NAME}" category="Object">
\t\t\t\t<xr:TypeId>4d92e7b5-1f60-4a83-b2c9-67e054d8f1a3</xr:TypeId>
\t\t\t\t<xr:ValueId>8c15f3a7-d249-4e06-91b8-2f74a6c0e5d9</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t</InternalInfo>
\t\t<Properties>
\t\t\t<Name>{NAME}</Name>
\t\t\t<Synonym>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>uk</v8:lang>
\t\t\t\t\t<v8:content>Початкова заборгованість по зарплаті за виплатами</v8:content>
\t\t\t\t</v8:item>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t<v8:content>Начальная задолженность по зарплате (по выплатам)</v8:content>
\t\t\t\t</v8:item>
\t\t\t</Synonym>
\t\t\t<Comment/>
\t\t\t<DefaultForm>ExternalDataProcessor.{NAME}.Form.Форма</DefaultForm>
\t\t\t<AuxiliaryForm/>
\t\t</Properties>
\t\t<ChildObjects>
'''.replace("{NAME}", NAME)

FOOTER = '''\t\t\t<Form>Форма</Form>
\t\t</ChildObjects>
\t</ExternalDataProcessor>
</MetaDataObject>
'''


def synonym(ru, uk=None):
    items = [f'<v8:item>\n<v8:lang>ru</v8:lang>\n<v8:content>{ru}</v8:content>\n</v8:item>']
    if uk:
        items.append(f'<v8:item>\n<v8:lang>uk</v8:lang>\n<v8:content>{uk}</v8:content>\n</v8:item>')
    return "<Synonym>\n" + "\n".join(items) + "\n</Synonym>"


def type_block(t):
    """t: ('date','Date'|'DateTime') | ('ref','cfg:...') | ('bool',) | ('num',d,f) | ('str',n)"""
    if t[0] == "date":
        return (f"<Type>\n<v8:Type>xs:dateTime</v8:Type>\n<v8:DateQualifiers>\n"
                f"<v8:DateFractions>{t[1]}</v8:DateFractions>\n</v8:DateQualifiers>\n</Type>")
    if t[0] == "ref":
        return f"<Type>\n<v8:Type>{t[1]}</v8:Type>\n</Type>"
    if t[0] == "bool":
        return "<Type>\n<v8:Type>xs:boolean</v8:Type>\n</Type>"
    if t[0] == "num":
        return (f"<Type>\n<v8:Type>xs:decimal</v8:Type>\n<v8:NumberQualifiers>\n"
                f"<v8:Digits>{t[1]}</v8:Digits>\n<v8:FractionDigits>{t[2]}</v8:FractionDigits>\n"
                f"<v8:AllowedSign>Any</v8:AllowedSign>\n</v8:NumberQualifiers>\n</Type>")
    if t[0] == "str":
        return (f"<Type>\n<v8:Type>xs:string</v8:Type>\n<v8:StringQualifiers>\n"
                f"<v8:Length>{t[1]}</v8:Length>\n<v8:AllowedLength>Variable</v8:AllowedLength>\n"
                f"</v8:StringQualifiers>\n</Type>")
    raise ValueError(t)


def attr_block(uuid, name, ru, uk, t, header_attr):
    """header_attr=True: реквизит шапки (без FillFromFillingValue/FillValue), False: колонка ТЧ."""
    fill = "" if header_attr else "<FillFromFillingValue>false</FillFromFillingValue>\n<FillValue xsi:nil=\"true\"/>\n"
    return f'''<Attribute uuid="{uuid}">
<Properties>
<Name>{name}</Name>
{synonym(ru, uk)}
<Comment/>
{type_block(t)}
<PasswordMode>false</PasswordMode>
<Format/>
<EditFormat/>
<ToolTip/>
<MarkNegatives>false</MarkNegatives>
<Mask/>
<MultiLine>false</MultiLine>
<ExtendedEdit>false</ExtendedEdit>
<MinValue xsi:nil="true"/>
<MaxValue xsi:nil="true"/>
{fill}<FillChecking>DontCheck</FillChecking>
<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
<ChoiceParameterLinks/>
<ChoiceParameters/>
<QuickChoice>Auto</QuickChoice>
<CreateOnInput>Auto</CreateOnInput>
<ChoiceForm/>
<LinkByType/>
<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
</Properties>
</Attribute>'''


STD_LINENUMBER = '''<StandardAttributes>
<xr:StandardAttribute name="LineNumber">
<xr:LinkByType/>
<xr:FillChecking>DontCheck</xr:FillChecking>
<xr:MultiLine>false</xr:MultiLine>
<xr:FillFromFillingValue>false</xr:FillFromFillingValue>
<xr:CreateOnInput>Auto</xr:CreateOnInput>
<xr:MaxValue xsi:nil="true"/>
<xr:ToolTip/>
<xr:ExtendedEdit>false</xr:ExtendedEdit>
<xr:Format/>
<xr:ChoiceForm/>
<xr:QuickChoice>Auto</xr:QuickChoice>
<xr:ChoiceHistoryOnInput>Auto</xr:ChoiceHistoryOnInput>
<xr:EditFormat/>
<xr:PasswordMode>false</xr:PasswordMode>
<xr:DataHistory>Use</xr:DataHistory>
<xr:MarkNegatives>false</xr:MarkNegatives>
<xr:MinValue xsi:nil="true"/>
<xr:Synonym/>
<xr:Comment/>
<xr:FullTextSearch>Use</xr:FullTextSearch>
<xr:ChoiceParameterLinks/>
<xr:FillValue xsi:nil="true"/>
<xr:Mask/>
<xr:ChoiceParameters/>
</xr:StandardAttribute>
</StandardAttributes>'''


def ts_block(uuid, t1, v1, t2, v2, name, ru, uk, cols):
    cols_xml = "\n".join(cols)
    return f'''<TabularSection uuid="{uuid}">
<InternalInfo>
<xr:GeneratedType name="DataProcessorTabularSection.{NAME}.{name}" category="TabularSection">
<xr:TypeId>{t1}</xr:TypeId>
<xr:ValueId>{v1}</xr:ValueId>
</xr:GeneratedType>
<xr:GeneratedType name="DataProcessorTabularSectionRow.{NAME}.{name}" category="TabularSectionRow">
<xr:TypeId>{t2}</xr:TypeId>
<xr:ValueId>{v2}</xr:ValueId>
</xr:GeneratedType>
</InternalInfo>
<Properties>
<Name>{name}</Name>
{synonym(ru, uk)}
<Comment/>
<ToolTip/>
<FillChecking>DontCheck</FillChecking>
{STD_LINENUMBER}
</Properties>
<ChildObjects>
{cols_xml}
</ChildObjects>
</TabularSection>'''


parts = [HEADER]

# --- Реквизиты шапки ---
parts.append(attr_block("f1a2b3c4-d5e6-4789-8abc-def012345601", "МесяцОстатков",
                        "Месяц остатков", "Місяць залишків", ("date", "Date"), True))
parts.append(attr_block("f1a2b3c4-d5e6-4789-8abc-def012345602", "МесяцВедомостей",
                        "Месяц ведомостей", "Місяць відомостей", ("date", "Date"), True))
parts.append(attr_block("f1a2b3c4-d5e6-4789-8abc-def012345603", "ДокументЦель",
                        "Документ начальной задолженности (цель)", "Документ початкової заборгованості (ціль)",
                        ("ref", "cfg:DocumentRef.НачальнаяЗадолженностьПоЗарплате"), True))

# --- ТЧ Ведомости ---
cols_ved = [
    attr_block("a2b3c4d5-e6f7-4801-9234-56789abcde01", "Выбрать", "Выбрать", "Вибрати", ("bool",), False),
    attr_block("a2b3c4d5-e6f7-4801-9234-56789abcde02", "Ведомость", "Ведомость", "Відомість",
               ("ref", "cfg:DocumentRef.ВедомостьНаВыплатуЗарплатыВКассу"), False),
    attr_block("a2b3c4d5-e6f7-4801-9234-56789abcde03", "Дата", "Дата", "Дата", ("date", "DateTime"), False),
    attr_block("a2b3c4d5-e6f7-4801-9234-56789abcde04", "СуммаПоДокументу", "Сумма по документу",
               "Сума за документом", ("num", 15, 2), False),
    attr_block("a2b3c4d5-e6f7-4801-9234-56789abcde05", "КоличествоСтрок", "Кол-во строк",
               "К-сть рядків", ("num", 10, 0), False),
]
parts.append(ts_block("b3c4d5e6-f708-4912-a345-6789abcdef01",
                      "c4d5e6f7-0819-4a23-b456-789abcdef012", "d5e6f708-192a-4b34-c567-89abcdef0123",
                      "e6f70819-2a3b-4c45-d678-9abcdef01234", "f708192a-3b4c-4d56-e789-abcdef012345",
                      "Ведомости", "Ведомости", "Відомості", cols_ved))

# --- ТЧ Превью ---
cols_prev = [
    attr_block("11223344-5566-4778-899a-bbccddeeff01", "Ведомость", "Ведомость", "Відомість",
               ("ref", "cfg:DocumentRef.ВедомостьНаВыплатуЗарплатыВКассу"), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff02", "Сотрудник", "Сотрудник", "Співробітник",
               ("ref", "cfg:CatalogRef.Сотрудники"), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff03", "ФизическоеЛицо", "Физическое лицо", "Фізична особа",
               ("ref", "cfg:CatalogRef.ФизическиеЛица"), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff04", "Подразделение", "Подразделение", "Підрозділ",
               ("ref", "cfg:CatalogRef.ПодразделенияОрганизаций"), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff05", "ПериодВзаиморасчетов", "Период взаиморасчетов",
               "Період взаєморозрахунків", ("date", "Date"), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff06", "ГруппаУчетаНачислений", "Группа учета начислений",
               "Група обліку нарахувань", ("ref", "cfg:CatalogRef.ГруппыУчетаНачисленийИУдержаний"), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff07", "СуммаПоВедомостям", "Сумма по ведомостям",
               "Сума за відомостями", ("num", 15, 2), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff08", "ОстатокЗКВ", "Остаток ЗКВ на дату",
               "Залишок ЗКВ на дату", ("num", 15, 2), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff09", "СуммаКЗаполнению", "Сумма к заполнению",
               "Сума до заповнення", ("num", 15, 2), False),
    attr_block("11223344-5566-4778-899a-bbccddeeff10", "Статус", "Статус", "Статус", ("str", 30), False),
]
parts.append(ts_block("22334455-6677-4889-9aab-ccddeeff0011",
                      "33445566-7788-499a-abbc-ddeeff001122", "44556677-8899-4aab-bccd-eeff00112233",
                      "55667788-99aa-4bbc-cdde-eff001122334", "66778899-aabb-4ccd-deef-f00112233445",
                      "Превью", "Превью (строки документа)", "Прев'ю (рядки документа)", cols_prev))

parts.append(FOOTER)

# Сборка: блоки внутри ChildObjects с отступом 3 таба; внутренние строки — относительный отступ
body_blocks = parts[1:-1]
indented = []
for block in body_blocks:
    lines = block.split("\n")
    out = []
    depth = 0
    for ln in lines:
        stripped = ln.strip()
        if stripped.startswith("</") and depth > 0:
            depth -= 1
        out.append("\t" * (3 + depth) + stripped)
        if (stripped.startswith("<") and not stripped.startswith("</")
                and not stripped.startswith("<?")
                and not ("</" in stripped) and not stripped.endswith("/>")):
            depth += 1
    indented.append("\n".join(out))

xml = parts[0] + "\n".join(indented) + "\n" + parts[-1]
xml = xml.replace("\r\n", "\n").replace("\n", "\r\n")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with io.open(OUT, "w", encoding="utf-8-sig", newline="") as fh:
    fh.write(xml)
print("OK", OUT, len(xml), "chars")
