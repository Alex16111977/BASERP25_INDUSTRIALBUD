# -*- coding: utf-8 -*-
"""Генератор маніфесту зовнішньої обробки «Друк М-29» (ПечатьМ29).

Спека: docs/superpowers/specs/2026-08-03-universal-m29-obrabotka-design.md
Пише `_Rarzrabotki/Обработки/Печать М29.xml` ПОВНІСТЮ (клон СозданиеКомплектацийБух
перезаписується), усі uuid — свіжі.

Запуск: python _Rarzrabotki/Python/test/m29_obr_build_meta.py
"""
import io
import os
import sys
import uuid

sys.stdout.reconfigure(encoding='utf-8')

БАЗА = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки"
ФАЙЛ = os.path.join(БАЗА, "Печать М29.xml")
ИМЯ = "ПечатьМ29"
СИНОНИМ = "Друк М-29"

NS = ('xmlns="http://v8.1c.ru/8.3/MDClasses" '
      'xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
      'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
      'xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" '
      'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" '
      'xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
      'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" '
      'xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
      'xmlns:v8="http://v8.1c.ru/8.1/data/core" '
      'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
      'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" '
      'xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
      'xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" '
      'xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" '
      'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
      'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')


def u():
    return str(uuid.uuid4())


# ---------------------------------------------------------------- типы
def тип_ссылка(*полные_имена):
    return "\n".join(f"{{о}}\t<v8:Type>cfg:{имя}</v8:Type>" for имя in полные_имена)


def тип_булево():
    return "{о}\t<v8:Type>xs:boolean</v8:Type>"


def тип_дата():
    return ("{о}\t<v8:Type>xs:dateTime</v8:Type>\n"
            "{о}\t<v8:DateQualifiers>\n"
            "{о}\t\t<v8:DateFractions>Date</v8:DateFractions>\n"
            "{о}\t</v8:DateQualifiers>")


def тип_строка(длина):
    return ("{о}\t<v8:Type>xs:string</v8:Type>\n"
            "{о}\t<v8:StringQualifiers>\n"
            f"{{о}}\t\t<v8:Length>{длина}</v8:Length>\n"
            "{о}\t\t<v8:AllowedLength>Variable</v8:AllowedLength>\n"
            "{о}\t</v8:StringQualifiers>")


def тип_число(разрядов, дробных):
    return ("{о}\t<v8:Type>xs:decimal</v8:Type>\n"
            "{о}\t<v8:NumberQualifiers>\n"
            f"{{о}}\t\t<v8:Digits>{разрядов}</v8:Digits>\n"
            f"{{о}}\t\t<v8:FractionDigits>{дробных}</v8:FractionDigits>\n"
            "{о}\t\t<v8:AllowedSign>Any</v8:AllowedSign>\n"
            "{о}\t</v8:NumberQualifiers>")


ДОК_ТИПЫ = ("DocumentRef.ТребованиеНакладная",
            "DocumentRef.СписаниеТоваров",
            "DocumentRef.КомплектацияНоменклатуры",
            "DocumentRef.ПередачаМалоценныхАктивовВЭксплуатацию")


# ---------------------------------------------------------------- рендер
def реквизит(имя, синоним, тип_xml, уровень, выбор_групп=False):
    """выбор_групп=True -> ChoiceFoldersAndItems=FoldersAndItems.

    Без цього платформа на спробі обрати папку каже «Выберите элемент, а не группу!»,
    і фільтр В ИЕРАРХИИ стає марним — групу нема як вказати.
    """
    о = "\t" * уровень
    тело = тип_xml.replace("{о}", о + "\t\t\t")
    выбор = "FoldersAndItems" if выбор_групп else "Items"
    return f"""{о}<Attribute uuid="{u()}">
{о}\t<Properties>
{о}\t\t<Name>{имя}</Name>
{о}\t\t<Synonym>
{о}\t\t\t<v8:item>
{о}\t\t\t\t<v8:lang>uk</v8:lang>
{о}\t\t\t\t<v8:content>{синоним}</v8:content>
{о}\t\t\t</v8:item>
{о}\t\t</Synonym>
{о}\t\t<Comment/>
{о}\t\t<Type>
{тело}
{о}\t\t</Type>
{о}\t\t<PasswordMode>false</PasswordMode>
{о}\t\t<Format/>
{о}\t\t<EditFormat/>
{о}\t\t<ToolTip/>
{о}\t\t<MarkNegatives>false</MarkNegatives>
{о}\t\t<Mask/>
{о}\t\t<MultiLine>false</MultiLine>
{о}\t\t<ExtendedEdit>false</ExtendedEdit>
{о}\t\t<MinValue xsi:nil="true"/>
{о}\t\t<MaxValue xsi:nil="true"/>
{о}\t\t<FillChecking>DontCheck</FillChecking>
{о}\t\t<ChoiceFoldersAndItems>{выбор}</ChoiceFoldersAndItems>
{о}\t\t<ChoiceParameterLinks/>
{о}\t\t<ChoiceParameters/>
{о}\t\t<QuickChoice>Auto</QuickChoice>
{о}\t\t<CreateOnInput>Auto</CreateOnInput>
{о}\t\t<ChoiceForm/>
{о}\t\t<LinkByType/>
{о}\t\t<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
{о}\t</Properties>
{о}</Attribute>"""


СТАНДАРТНЫЙ_НОМЕР_СТРОКИ = """\t\t\t\t\t<StandardAttributes>
\t\t\t\t\t\t<xr:StandardAttribute name="LineNumber">
\t\t\t\t\t\t\t<xr:LinkByType/>
\t\t\t\t\t\t\t<xr:FillChecking>DontCheck</xr:FillChecking>
\t\t\t\t\t\t\t<xr:MultiLine>false</xr:MultiLine>
\t\t\t\t\t\t\t<xr:FillFromFillingValue>false</xr:FillFromFillingValue>
\t\t\t\t\t\t\t<xr:CreateOnInput>Auto</xr:CreateOnInput>
\t\t\t\t\t\t\t<xr:MaxValue xsi:nil="true"/>
\t\t\t\t\t\t\t<xr:ToolTip/>
\t\t\t\t\t\t\t<xr:ExtendedEdit>false</xr:ExtendedEdit>
\t\t\t\t\t\t\t<xr:Format/>
\t\t\t\t\t\t\t<xr:ChoiceForm/>
\t\t\t\t\t\t\t<xr:QuickChoice>Auto</xr:QuickChoice>
\t\t\t\t\t\t\t<xr:ChoiceHistoryOnInput>Auto</xr:ChoiceHistoryOnInput>
\t\t\t\t\t\t\t<xr:EditFormat/>
\t\t\t\t\t\t\t<xr:PasswordMode>false</xr:PasswordMode>
\t\t\t\t\t\t\t<xr:DataHistory>Use</xr:DataHistory>
\t\t\t\t\t\t\t<xr:MarkNegatives>false</xr:MarkNegatives>
\t\t\t\t\t\t\t<xr:MinValue xsi:nil="true"/>
\t\t\t\t\t\t\t<xr:Synonym/>
\t\t\t\t\t\t\t<xr:Comment/>
\t\t\t\t\t\t\t<xr:FullTextSearch>Use</xr:FullTextSearch>
\t\t\t\t\t\t\t<xr:ChoiceParameterLinks/>
\t\t\t\t\t\t\t<xr:FillValue xsi:nil="true"/>
\t\t\t\t\t\t\t<xr:Mask/>
\t\t\t\t\t\t\t<xr:ChoiceParameters/>
\t\t\t\t\t\t</xr:StandardAttribute>
\t\t\t\t\t</StandardAttributes>"""


def табличная_часть(имя, синоним, колонки):
    # колонка = (имя, синоним, тип) або (имя, синоним, тип, выбор_групп)
    рекв = "\n".join(реквизит(к[0], к[1], к[2], 5, к[3] if len(к) > 3 else False)
                     for к in колонки)
    return f"""\t\t\t<TabularSection uuid="{u()}">
\t\t\t\t<InternalInfo>
\t\t\t\t\t<xr:GeneratedType name="DataProcessorTabularSection.{ИМЯ}.{имя}" category="TabularSection">
\t\t\t\t\t\t<xr:TypeId>{u()}</xr:TypeId>
\t\t\t\t\t\t<xr:ValueId>{u()}</xr:ValueId>
\t\t\t\t\t</xr:GeneratedType>
\t\t\t\t\t<xr:GeneratedType name="DataProcessorTabularSectionRow.{ИМЯ}.{имя}" category="TabularSectionRow">
\t\t\t\t\t\t<xr:TypeId>{u()}</xr:TypeId>
\t\t\t\t\t\t<xr:ValueId>{u()}</xr:ValueId>
\t\t\t\t\t</xr:GeneratedType>
\t\t\t\t</InternalInfo>
\t\t\t\t<Properties>
\t\t\t\t\t<Name>{имя}</Name>
\t\t\t\t\t<Synonym>
\t\t\t\t\t\t<v8:item>
\t\t\t\t\t\t\t<v8:lang>uk</v8:lang>
\t\t\t\t\t\t\t<v8:content>{синоним}</v8:content>
\t\t\t\t\t\t</v8:item>
\t\t\t\t\t</Synonym>
\t\t\t\t\t<Comment/>
\t\t\t\t\t<ToolTip/>
\t\t\t\t\t<FillChecking>DontCheck</FillChecking>
{СТАНДАРТНЫЙ_НОМЕР_СТРОКИ}
\t\t\t\t</Properties>
\t\t\t\t<ChildObjects>
{рекв}
\t\t\t\t</ChildObjects>
\t\t\t</TabularSection>"""


# ---------------------------------------------------------------- состав
РЕКВИЗИТЫ = [
    ("ПериодС", "Період з", тип_дата()),
    ("ПериодПо", "Період по", тип_дата()),
    ("Организация", "Організація", тип_ссылка("CatalogRef.Организации")),
    ("ВключатьТребованияНакладные", "Вимоги-накладні", тип_булево()),
    ("ВключатьСписанияТоваров", "Списання товарів", тип_булево()),
    ("ВключатьКомплектации", "Комплектації номенклатури", тип_булево()),
    ("ВключатьПередачиМалоценки", "Передачі малоцінки в експлуатацію", тип_булево()),
    ("НаименованиеРаботы", "Найменування роботи (обʼєкт)", тип_строка(150)),
    ("НомерАкта", "До Акту №", тип_строка(30)),
    ("Бухгалтер", "Опрацював  Бухгалтер", тип_ссылка("CatalogRef.ФизическиеЛица")),
    ("Проверил", "Витрати матеріалів перевірив", тип_ссылка("CatalogRef.ФизическиеЛица")),
    ("НачальникВТО", "Начальник ВТВ", тип_ссылка("CatalogRef.ФизическиеЛица")),
    ("Инженер", "Інженер ВТВ", тип_ссылка("CatalogRef.ФизическиеЛица")),
    ("Начальник", "Начальник будівельного-монтажного комплекса", тип_ссылка("CatalogRef.ФизическиеЛица")),
    ("МОЛ", "Матеріально-відповідальна особа, що склала звіт", тип_ссылка("CatalogRef.ФизическиеЛица")),
]

ТАБЛИЧНЫЕ_ЧАСТИ = [
    # Фільтри-списки. Порожній список = без обмеження; заповнений застосовується
    # В ИЕРАРХИИ, тому можна вибрати групу складів / батьківський підрозділ.
    ("Склады", "Склади", [
        # ⚠ FoldersAndItems — інакше платформа не дасть обрати папку складів
        ("Склад", "Склад (можна групу)", тип_ссылка("CatalogRef.Склады"), True),
    ]),
    # ⚠ Справочник.Подразделения має синонім «Підрозділи Казна» — це реквізит
    # Документ.X.Подразделение, ЄДИНИЙ спільний для всіх 4 видів документів.
    # НЕ плутати з ПодразделениеОрганизации (Справочник.ПодразделенияОрганизаций,
    # синонім «Підрозділ»), якого немає у СписаниеТоваров і КомплектацияНоменклатуры.
    ("Подразделения", "Підрозділи", [
        # Подразделения — ієрархія ЕЛЕМЕНТІВ (папок немає), але FoldersAndItems
        # ставимо для симетрії й на випадок зміни виду ієрархії
        ("Подразделение", "Підрозділ (можна батьківський)",
         тип_ссылка("CatalogRef.Подразделения"), True),
    ]),
    ("Документы", "Документи", [
        ("Пометка", "Друкувати", тип_булево()),
        ("Документ", "Документ", тип_ссылка(*ДОК_ТИПЫ)),
        ("ДатаДок", "Дата", тип_дата()),
        ("Организация", "Організація", тип_ссылка("CatalogRef.Организации")),
        ("Склад", "Склад", тип_ссылка("CatalogRef.Склады")),
        ("Подразделение", "Підрозділ", тип_ссылка("CatalogRef.Подразделения")),
        ("СуммаПоПроводкам", "Сума за проводками", тип_число(15, 2)),
    ]),
    ("СчетаМШП", "Рахунки МШП", [
        ("Счет", "Рахунок", тип_ссылка("ChartOfAccountsRef.Хозрасчетный")),
    ]),
]


def main():
    рекв = "\n".join(реквизит(н, с, т, 3) for н, с, т in РЕКВИЗИТЫ)
    тч = "\n".join(табличная_часть(и, с, к) for и, с, к in ТАБЛИЧНЫЕ_ЧАСТИ)

    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject {NS}>
\t<ExternalDataProcessor uuid="{u()}">
\t\t<InternalInfo>
\t\t\t<xr:ContainedObject>
\t\t\t\t<xr:ClassId>c3831ec8-d8d5-4f93-8a22-f9bfae07327f</xr:ClassId>
\t\t\t\t<xr:ObjectId>{u()}</xr:ObjectId>
\t\t\t</xr:ContainedObject>
\t\t\t<xr:GeneratedType name="ExternalDataProcessorObject.{ИМЯ}" category="Object">
\t\t\t\t<xr:TypeId>{u()}</xr:TypeId>
\t\t\t\t<xr:ValueId>{u()}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t</InternalInfo>
\t\t<Properties>
\t\t\t<Name>{ИМЯ}</Name>
\t\t\t<Synonym>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>uk</v8:lang>
\t\t\t\t\t<v8:content>{СИНОНИМ}</v8:content>
\t\t\t\t</v8:item>
\t\t\t</Synonym>
\t\t\t<Comment/>
\t\t\t<DefaultForm>ExternalDataProcessor.{ИМЯ}.Form.Форма</DefaultForm>
\t\t\t<AuxiliaryForm/>
\t\t</Properties>
\t\t<ChildObjects>
{рекв}
{тч}
\t\t\t<Form>Форма</Form>
\t\t</ChildObjects>
\t</ExternalDataProcessor>
</MetaDataObject>"""

    with io.open(ФАЙЛ, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(xml)

    print(f"Записано: {ФАЙЛ}")
    print(f"  реквізитів: {len(РЕКВИЗИТЫ)}")
    for и, с, к in ТАБЛИЧНЫЕ_ЧАСТИ:
        print(f"  ТЧ {и}: {len(к)} колонок")


if __name__ == "__main__":
    main()
