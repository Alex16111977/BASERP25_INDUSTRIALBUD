# -*- coding: utf-8 -*-
"""Генератор исходников внешнего отчёта А_ОтчетПоДебиторскаяЗадолженностьПоставщиковПоСрокамДолгаИстатусам
(корневой XML, метаданные макета, макет СКД). Модуль объекта пишется руками (Ext/ObjectModule.bsl).
Запрос в макете — N=3 из otchet_debitorka_statusy_zapros.текст_запроса (проверен через COM, Rule #-1).
Отборы варианта — рабочие отборы бухгалтера (otchet_debitorka_statusy_otbory.xml, из её сохранённых настроек).
    python otchet_debitorka_statusy_generate.py"""
import io, os, sys
from xml.sax.saxutils import escape

sys.stdout.reconfigure(encoding="utf-8")
ПАПКА = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ПАПКА)
from otchet_debitorka_statusy_zapros import текст_запроса  # noqa: E402

ИМЯ = "А_ОтчетПоДебиторскаяЗадолженностьПоставщиковПоСрокамДолгаИстатусам"
СИНОНИМ = "Дебіторська заборгованість постачальників за строками боргу та статусами"
КОРЕНЬ = os.path.normpath(os.path.join(ПАПКА, "..", "..", "Отчеты"))
# постоянные идентификаторы (сгенерированы один раз 30.09.2026; ClassId — константа класса внешних отчётов)
UUID_ОТЧЕТА, OBJECT_ID, TYPE_ID, VALUE_ID = ("ea218402-0112-48f9-b198-bc316929d7d6", "e50cac45-ebfa-4348-a836-92cec31e2763",
                                            "24fd4f4d-906c-4194-b5ec-9197e0edc31c", "3d3d9c16-474f-4464-aa63-cfd6e9d1b9d0")
UUID_МАКЕТА = "99c7bea6-d109-4c19-a3d2-8936ca661f46"
ID_ПЕРИОД, ID_ОРГ, ID_ИНТЕРВАЛЫ, ID_ДАТАБАС, ID_ОТБОР = (
    "468c0905-b8fa-40f3-a43c-8f1cacbf246c", "55ca0e99-fd95-496c-9436-71199ae078af", "ea3868fe-3f94-498a-a09c-8ceca5ee4a07",
    "4016bfd2-862a-476e-acf0-f9abdcac6570", "6332985a-f8a4-4ab2-98b9-fde9771c7ed0")
CFG = 'xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config"'

NS_MD = ('xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
         'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" '
         'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
         'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
         'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
         'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
         'xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" '
         'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
         'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')


def записать(отн_путь, текст):
    путь = os.path.join(КОРЕНЬ, отн_путь)
    os.makedirs(os.path.dirname(путь), exist_ok=True)
    with io.open(путь, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(текст)
    print("записан:", путь)


def строка_ls(текст, отступ):
    т = escape(текст)
    return (f'{отступ}<v8:item>\n{отступ}\t<v8:lang>ru</v8:lang>\n{отступ}\t<v8:content>{т}</v8:content>\n{отступ}</v8:item>\n'
            f'{отступ}<v8:item>\n{отступ}\t<v8:lang>uk</v8:lang>\n{отступ}\t<v8:content>{т}</v8:content>\n{отступ}</v8:item>\n')


def заголовок(текст, отступ="\t\t\t"):
    return f'{отступ}<title xsi:type="v8:LocalStringType">\n{строка_ls(текст, отступ + chr(9))}{отступ}</title>\n'


def формат(текст, отступ="\t\t\t"):
    return (f'{отступ}<appearance>\n{отступ}\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n'
            f'{отступ}\t\t<dcscor:parameter>Формат</dcscor:parameter>\n{отступ}\t\t<dcscor:value xsi:type="xs:string">{текст}</dcscor:value>\n'
            f'{отступ}\t</dcscor:item>\n{отступ}</appearance>\n')


def тип_ссылки(имя, отступ="\t\t\t"):
    return f'{отступ}<valueType>\n{отступ}\t<v8:Type {CFG}>d5p1:{имя}</v8:Type>\n{отступ}</valueType>\n'


def тип_числа(отступ="\t\t\t", дробь=2):
    return (f'{отступ}<valueType>\n{отступ}\t<v8:Type>xs:decimal</v8:Type>\n{отступ}\t<v8:NumberQualifiers>\n'
            f'{отступ}\t\t<v8:Digits>15</v8:Digits>\n{отступ}\t\t<v8:FractionDigits>{дробь}</v8:FractionDigits>\n'
            f'{отступ}\t\t<v8:AllowedSign>Any</v8:AllowedSign>\n{отступ}\t</v8:NumberQualifiers>\n{отступ}</valueType>\n')


def тип_строки(длина, отступ="\t\t\t"):
    return (f'{отступ}<valueType>\n{отступ}\t<v8:Type>xs:string</v8:Type>\n{отступ}\t<v8:StringQualifiers>\n'
            f'{отступ}\t\t<v8:Length>{длина}</v8:Length>\n{отступ}\t\t<v8:AllowedLength>Variable</v8:AllowedLength>\n'
            f'{отступ}\t</v8:StringQualifiers>\n{отступ}</valueType>\n')


def тип_даты(отступ="\t\t"):
    return (f'{отступ}<valueType>\n{отступ}\t<v8:Type>xs:dateTime</v8:Type>\n{отступ}\t<v8:DateQualifiers>\n'
            f'{отступ}\t\t<v8:DateFractions>DateTime</v8:DateFractions>\n{отступ}\t</v8:DateQualifiers>\n{отступ}</valueType>\n')


ИЗМЕРЕНИЕ = ('\t\t\t<role>\n\t\t\t\t<dcscom:dimension>true</dcscom:dimension>\n\t\t\t\t<dcscom:ignoreNullValues>true</dcscom:ignoreNullValues>\n'
             '\t\t\t\t<dcscom:required>true</dcscom:required>\n\t\t\t</role>\n')


def поле(путь, титул, роль="", тип="", оформление=""):
    return (f'\t\t<field xsi:type="DataSetFieldField">\n\t\t\t<dataPath>{путь}</dataPath>\n\t\t\t<field>{путь}</field>\n'
            f'{заголовок(титул)}{роль}{тип}{оформление}\t\t</field>\n')


ИНТЕРВАЛЫ_ПО_УМОЛЧАНИЮ = ["До 7 днів", "Від 8 до 15 днів", "Від 16 до 30 днів", "Понад 30 днів"]


def макет_скд():
    x = ['<?xml version="1.0" encoding="UTF-8"?>\n'
         '<DataCompositionSchema xmlns="http://v8.1c.ru/8.1/data-composition-system/schema" '
         'xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" '
         'xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" xmlns:v8="http://v8.1c.ru/8.1/data/core" '
         'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
         'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
         '\t<dataSource>\n\t\t<name>ИнформационнаяБаза</name>\n\t\t<dataSourceType>Local</dataSourceType>\n\t</dataSource>\n'
         '\t<dataSet xsi:type="DataSetQuery">\n\t\t<name>ОсновнойНабор</name>\n']
    x.append(поле("Организация", "Організація", ИЗМЕРЕНИЕ, тип_ссылки("CatalogRef.Организации")))
    x.append(поле("Контрагент", "Постачальник", ИЗМЕРЕНИЕ, тип_ссылки("CatalogRef.Контрагенты")))
    x.append(поле("Договор", "Договір", ИЗМЕРЕНИЕ, тип_ссылки("CatalogRef.ДоговорыКонтрагентов")))
    x.append(поле("СчетНаОплату", "Рахунок на оплату", ИЗМЕРЕНИЕ))
    x.append(поле("Статус", "Статус боргу", "\t\t\t<role>\n\t\t\t\t<dcscom:dimension>true</dcscom:dimension>\n\t\t\t</role>\n",
                  тип_строки(20)))
    x.append(поле("ОстатокДолга", "Загальна заборгованість", "", тип_числа(), формат("ЧЦ=15; ЧДЦ=2")))
    for i, титул in enumerate(ИНТЕРВАЛЫ_ПО_УМОЛЧАНИЮ, 1):
        x.append(поле(f"ОстатокПериода{i}", титул, "", тип_числа(), формат("ЧЦ=15; ЧДЦ=0")))
    x.append(f'\t\t<dataSource>ИнформационнаяБаза</dataSource>\n\t\t<query>{escape(текст_запроса(3))}</query>\n'
             '\t\t<autoFillFields>false</autoFillFields>\n\t</dataSet>\n')
    # внешний набор просрочки (таблица из модуля)
    x.append('\t<dataSet xsi:type="DataSetObject">\n\t\t<name>ПросроченнаяЗадолженность</name>\n')
    x.append(поле("Организация", "Організація", "", тип_ссылки("CatalogRef.Организации")))
    x.append(поле("Контрагент", "Постачальник", "", тип_ссылки("CatalogRef.Контрагенты")))
    x.append(поле("Договор", "Договір", "", тип_ссылки("CatalogRef.ДоговорыКонтрагентов")))
    x.append(поле("СчетНаОплату", "Рахунок на оплату"))
    x.append(поле("Статус", "Статус боргу", "", тип_строки(20)))
    x.append(поле("ПросроченнаяЗадолженность", "У т.ч. прострочена заборгованість", "", тип_числа(), формат("ЧЦ=15; ЧДЦ=0")))
    x.append('\t\t<dataSource>ИнформационнаяБаза</dataSource>\n\t\t<objectName>ПросроченнаяЗадолженность</objectName>\n\t</dataSet>\n')
    for ключ in ("Организация", "Контрагент", "Договор", "СчетНаОплату", "Статус"):
        x.append(f'\t<dataSetLink>\n\t\t<sourceDataSet>ОсновнойНабор</sourceDataSet>\n\t\t<destinationDataSet>ПросроченнаяЗадолженность</destinationDataSet>\n'
                 f'\t\t<sourceExpression>{ключ}</sourceExpression>\n\t\t<destinationExpression>{ключ}</destinationExpression>\n\t</dataSetLink>\n')
    # «Статус» колонкой на уровне документа: ресурс считается только в группировке СчетНаОплату
    x.append('\t<calculatedField>\n\t\t<dataPath>СтатусДокумента</dataPath>\n\t\t<expression>Статус</expression>\n'
             f'{заголовок("Статус", chr(9) * 2)}{тип_строки(20, chr(9) * 2)}\t</calculatedField>\n')
    for путь in ["ОстатокДолга"] + [f"ОстатокПериода{i}" for i in range(1, 5)] + ["ПросроченнаяЗадолженность"]:
        x.append(f'\t<totalField>\n\t\t<dataPath>{путь}</dataPath>\n\t\t<expression>Сумма({путь})</expression>\n\t</totalField>\n')
    x.append('\t<totalField>\n\t\t<dataPath>СтатусДокумента</dataPath>\n\t\t<expression>Максимум(Статус)</expression>\n'
             '\t\t<group>СчетНаОплату</group>\n\t</totalField>\n')
    # параметры
    def параметр(имя, титул, тип, значение, скрыт, список=False):
        return (f'\t<parameter>\n\t\t<name>{имя}</name>\n{заголовок(титул, chr(9) * 2)}{тип}{значение}'
                f'\t\t<useRestriction>{"true" if скрыт else "false"}</useRestriction>\n'
                + ('\t\t<valueListAllowed>true</valueListAllowed>\n' if список else '') + '\t</parameter>\n')
    дата_только = ('\t\t<valueType>\n\t\t\t<v8:Type>xs:dateTime</v8:Type>\n\t\t\t<v8:DateQualifiers>\n'
                   '\t\t\t\t<v8:DateFractions>Date</v8:DateFractions>\n\t\t\t</v8:DateQualifiers>\n\t\t</valueType>\n')
    x.append(параметр("Период", "На дату", дата_только, '\t\t<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>\n', False))
    x.append(параметр("Организация", "Організація", тип_ссылки("CatalogRef.Организации", chr(9) * 2),
                      '\t\t<value xsi:nil="true"/>\n', False))
    x.append(параметр("Интервалы", "Інтервали, днів", тип_строки(100, chr(9) * 2),
                      '\t\t<value xsi:type="xs:string">7, 15, 30</value>\n', False))
    x.append(параметр("ДатаНачалаБас", "Дата початку обліку в BAS", дата_только,
                      '\t\t<value xsi:type="xs:dateTime">2023-07-01T00:00:00</value>\n', False))
    x.append(параметр("ВидыСубконтоКД", "Види субконто розрахунків",
                      тип_ссылки("ChartOfCharacteristicTypesRef.ВидыСубконтоХозрасчетные", chr(9) * 2), "", True, True))
    x.append(параметр("ИсключенныеСчета", "Виключені рахунки", тип_ссылки("ChartOfAccountsRef.Хозрасчетный", chr(9) * 2),
                      "", True, True))
    for i in range(1, 4):
        x.append(параметр(f"НачалоИнтервала{i}", f"Початок інтервалу {i}", тип_даты(),
                          '\t\t<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>\n', True))
        x.append(параметр(f"КонецИнтервала{i}", f"Кінець інтервалу {i}", тип_даты(),
                          '\t\t<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>\n', True))
    # вариант «Основний» — рабочие настройки бухгалтера + колонка «Статус»
    отборы = io.open(os.path.join(ПАПКА, "otchet_debitorka_statusy_otbory.xml"), encoding="utf-8").read().strip("\n")
    # Ссылки на данные базы в макете внешнего отчёта не переживают сборку .erf (становятся строками), поэтому
    # значения «Не в списку» для Постачальника и Договір.Підрозділ подставляет модуль (ПередЗагрузкойНастроекВКомпоновщик).
    отборы = "\n".join(л for л in отборы.split("\n")
                       if 'CatalogRef.Контрагенты">' not in л and 'CatalogRef.Подразделения">' not in л)
    отборы = "\n".join("\t\t\t\t" + л for л in отборы.split("\n"))
    отбор_статуса = ('\t\t\t\t<dcsset:item xsi:type="dcsset:FilterItemComparison">\n\t\t\t\t\t<dcsset:use>false</dcsset:use>\n'
                     '\t\t\t\t\t<dcsset:left xsi:type="dcscor:Field">Статус</dcsset:left>\n'
                     '\t\t\t\t\t<dcsset:comparisonType>Equal</dcsset:comparisonType>\n'
                     '\t\t\t\t\t<dcsset:right xsi:type="xs:string">Передоплата</dcsset:right>\n\t\t\t\t</dcsset:item>\n')

    def группировка(поле_гр, тип_гр, вложенная, отступ):
        о = отступ
        return (f'{о}<dcsset:item xsi:type="dcsset:StructureItemGroup">\n{о}\t<dcsset:groupItems>\n'
                f'{о}\t\t<dcsset:item xsi:type="dcsset:GroupItemField">\n{о}\t\t\t<dcsset:field>{поле_гр}</dcsset:field>\n'
                f'{о}\t\t\t<dcsset:groupType>{тип_гр}</dcsset:groupType>\n{о}\t\t\t<dcsset:periodAdditionType>None</dcsset:periodAdditionType>\n'
                f'{о}\t\t\t<dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>\n'
                f'{о}\t\t\t<dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>\n'
                f'{о}\t\t</dcsset:item>\n{о}\t</dcsset:groupItems>\n'
                f'{о}\t<dcsset:order>\n{о}\t\t<dcsset:item xsi:type="dcsset:OrderItemAuto"/>\n{о}\t</dcsset:order>\n'
                f'{о}\t<dcsset:selection>\n{о}\t\t<dcsset:item xsi:type="dcsset:SelectedItemAuto"/>\n{о}\t</dcsset:selection>\n'
                f'{вложенная}{о}</dcsset:item>\n')
    структура = ""
    уровни = [("Договор.Ответственный", "Items"), ("Договор.Подразделение", "Items"), ("Контрагент", "Hierarchy"),
              ("Договор", "Items"), ("СчетНаОплату", "Items")]
    for глубина in range(len(уровни) - 1, -1, -1):
        поле_гр, тип_гр = уровни[глубина]
        структура = группировка(поле_гр, тип_гр, структура, "\t\t\t" + "\t" * глубина)

    def знач_пар(имя, значение, ид, обычный):
        return (f'\t\t\t\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n\t\t\t\t\t<dcscor:parameter>{имя}</dcscor:parameter>\n'
                f'\t\t\t\t\t{значение}\n' + ('\t\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n' if обычный else '')
                + f'\t\t\t\t\t<dcsset:userSettingID>{ид}</dcsset:userSettingID>\n\t\t\t\t</dcscor:item>\n')
    интервалы_выбор = "".join(f'\t\t\t\t\t<dcsset:item xsi:type="dcsset:SelectedItemField">\n\t\t\t\t\t\t<dcsset:field>ОстатокПериода{i}</dcsset:field>\n'
                              '\t\t\t\t\t</dcsset:item>\n' for i in range(1, 5))
    x.append('\t<settingsVariant>\n\t\t<dcsset:name>Основний</dcsset:name>\n\t\t<dcsset:presentation xsi:type="v8:LocalStringType">\n'
             + строка_ls("Основний", "\t\t\t") + '\t\t</dcsset:presentation>\n'
             '\t\t<dcsset:settings xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
             'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows">\n'
             '\t\t\t<dcsset:selection>\n'
             '\t\t\t\t<dcsset:item xsi:type="dcsset:SelectedItemField">\n\t\t\t\t\t<dcsset:field>СчетНаОплату.НомерВходящегоДокумента</dcsset:field>\n'
             '\t\t\t\t\t<dcsset:title>Вх. номер</dcsset:title>\n\t\t\t\t</dcsset:item>\n'
             '\t\t\t\t<dcsset:item xsi:type="dcsset:SelectedItemField">\n\t\t\t\t\t<dcsset:field>ОстатокДолга</dcsset:field>\n\t\t\t\t</dcsset:item>\n'
             '\t\t\t\t<dcsset:item xsi:type="dcsset:SelectedItemFolder">\n\t\t\t\t\t<dcsset:title>Загальна заборгованість по строках боргу</dcsset:title>\n'
             + интервалы_выбор + '\t\t\t\t\t<dcsset:placement>Auto</dcsset:placement>\n\t\t\t\t</dcsset:item>\n'
             '\t\t\t\t<dcsset:item xsi:type="dcsset:SelectedItemField">\n\t\t\t\t\t<dcsset:use>false</dcsset:use>\n'
             '\t\t\t\t\t<dcsset:field>ПросроченнаяЗадолженность</dcsset:field>\n\t\t\t\t</dcsset:item>\n'
             '\t\t\t\t<dcsset:item xsi:type="dcsset:SelectedItemField">\n\t\t\t\t\t<dcsset:field>СтатусДокумента</dcsset:field>\n\t\t\t\t</dcsset:item>\n'
             '\t\t\t</dcsset:selection>\n'
             '\t\t\t<dcsset:filter>\n' + отборы + '\n' + отбор_статуса
             + f'\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n\t\t\t\t<dcsset:userSettingID>{ID_ОТБОР}</dcsset:userSettingID>\n'
             '\t\t\t</dcsset:filter>\n'
             '\t\t\t<dcsset:dataParameters>\n'
             + знач_пар("Период", '<dcscor:value xsi:type="xs:dateTime">0001-01-01T00:00:00</dcscor:value>', ID_ПЕРИОД, False)
             + знач_пар("Организация", '<dcscor:value xsi:nil="true"/>', ID_ОРГ, False)
             + знач_пар("Интервалы", '<dcscor:value xsi:type="xs:string">7, 15, 30</dcscor:value>', ID_ИНТЕРВАЛЫ, True)
             + знач_пар("ДатаНачалаБас", '<dcscor:value xsi:type="xs:dateTime">2023-07-01T00:00:00</dcscor:value>', ID_ДАТАБАС, True)
             + '\t\t\t</dcsset:dataParameters>\n'
             '\t\t\t<dcsset:order>\n\t\t\t\t<dcsset:item xsi:type="dcsset:OrderItemField">\n\t\t\t\t\t<dcsset:field>ОстатокДолга</dcsset:field>\n'
             '\t\t\t\t\t<dcsset:orderType>Desc</dcsset:orderType>\n\t\t\t\t</dcsset:item>\n\t\t\t</dcsset:order>\n'
             '\t\t\t<dcsset:conditionalAppearance>\n\t\t\t\t<dcsset:item>\n\t\t\t\t\t<dcsset:selection>\n\t\t\t\t\t\t<dcsset:item>\n'
             '\t\t\t\t\t\t\t<dcsset:field>Контрагент</dcsset:field>\n\t\t\t\t\t\t</dcsset:item>\n\t\t\t\t\t</dcsset:selection>\n'
             '\t\t\t\t\t<dcsset:filter/>\n\t\t\t\t\t<dcsset:appearance>\n'
             '\t\t\t\t\t\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n\t\t\t\t\t\t\t<dcscor:parameter>ЦветТекста</dcscor:parameter>\n'
             '\t\t\t\t\t\t\t<dcscor:value xsi:type="v8ui:Color">#B22222</dcscor:value>\n\t\t\t\t\t\t</dcscor:item>\n'
             '\t\t\t\t\t\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n\t\t\t\t\t\t\t<dcscor:parameter>Шрифт</dcscor:parameter>\n'
             '\t\t\t\t\t\t\t<dcscor:value xsi:type="v8ui:Font" ref="sys:DefaultGUIFont" height="10" bold="true" italic="false" underline="false" '
             'strikeout="false" kind="WindowsFont"/>\n\t\t\t\t\t\t</dcscor:item>\n'
             '\t\t\t\t\t</dcsset:appearance>\n\t\t\t\t</dcsset:item>\n\t\t\t</dcsset:conditionalAppearance>\n'
             '\t\t\t<dcsset:outputParameters>\n'
             '\t\t\t\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n\t\t\t\t\t<dcscor:parameter>МакетОформления</dcscor:parameter>\n'
             '\t\t\t\t\t<dcscor:value xsi:type="xs:string">Море</dcscor:value>\n\t\t\t\t</dcscor:item>\n'
             # «Вх. номер» — в отдельной колонке, как «Розміщення: В окремих колонках» у типового
             '\t\t\t\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n\t\t\t\t\t<dcscor:parameter>РасположениеРеквизитов</dcscor:parameter>\n'
             '\t\t\t\t\t<dcscor:value xsi:type="dcsset:DataCompositionAttributesPlacement">Separately</dcscor:value>\n\t\t\t\t</dcscor:item>\n'
             + "".join(f'\t\t\t\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n\t\t\t\t\t<dcscor:parameter>{п}</dcscor:parameter>\n'
                       '\t\t\t\t\t<dcscor:value xsi:type="dcsset:DataCompositionTextOutputType">DontOutput</dcscor:value>\n\t\t\t\t</dcscor:item>\n'
                       for п in ("ВыводитьЗаголовок", "ВыводитьПараметрыДанных", "ВыводитьОтбор"))
             + '\t\t\t</dcsset:outputParameters>\n' + структура + '\t\t</dcsset:settings>\n\t</settingsVariant>\n')
    x.append('</DataCompositionSchema>')
    return "".join(x)


def корневой_xml():
    return (f'<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject {NS_MD}>\n\t<ExternalReport uuid="{UUID_ОТЧЕТА}">\n'
            '\t\t<InternalInfo>\n\t\t\t<xr:ContainedObject>\n\t\t\t\t<xr:ClassId>e41aff26-25cf-4bb6-b6c1-3f478a75f374</xr:ClassId>\n'
            f'\t\t\t\t<xr:ObjectId>{OBJECT_ID}</xr:ObjectId>\n\t\t\t</xr:ContainedObject>\n'
            f'\t\t\t<xr:GeneratedType name="ExternalReportObject.{ИМЯ}" category="Object">\n'
            f'\t\t\t\t<xr:TypeId>{TYPE_ID}</xr:TypeId>\n\t\t\t\t<xr:ValueId>{VALUE_ID}</xr:ValueId>\n\t\t\t</xr:GeneratedType>\n'
            f'\t\t</InternalInfo>\n\t\t<Properties>\n\t\t\t<Name>{ИМЯ}</Name>\n\t\t\t<Synonym>\n{строка_ls(СИНОНИМ, chr(9) * 4)}\t\t\t</Synonym>\n'
            '\t\t\t<Comment/>\n\t\t\t<DefaultForm>CommonForm.ФормаОтчета</DefaultForm>\n\t\t\t<AuxiliaryForm/>\n'
            f'\t\t\t<MainDataCompositionSchema>ExternalReport.{ИМЯ}.Template.ОсновнаяСхемаКомпоновкиДанных</MainDataCompositionSchema>\n'
            '\t\t\t<DefaultSettingsForm>CommonForm.ФормаНастроекОтчета</DefaultSettingsForm>\n\t\t\t<AuxiliarySettingsForm/>\n'
            '\t\t\t<DefaultVariantForm>CommonForm.ФормаВариантаОтчета</DefaultVariantForm>\n\t\t\t<VariantsStorage/>\n'
            '\t\t\t<SettingsStorage/>\n\t\t</Properties>\n\t\t<ChildObjects>\n\t\t\t<Template>ОсновнаяСхемаКомпоновкиДанных</Template>\n'
            '\t\t</ChildObjects>\n\t</ExternalReport>\n</MetaDataObject>')


def метаданные_макета():
    return (f'<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject {NS_MD}>\n\t<Template uuid="{UUID_МАКЕТА}">\n\t\t<Properties>\n'
            '\t\t\t<Name>ОсновнаяСхемаКомпоновкиДанных</Name>\n\t\t\t<Synonym>\n'
            + строка_ls("Основна схема компонування даних", "\t\t\t\t")
            + '\t\t\t</Synonym>\n\t\t\t<Comment/>\n\t\t\t<TemplateType>DataCompositionSchema</TemplateType>\n'
            '\t\t</Properties>\n\t</Template>\n</MetaDataObject>')


if __name__ == "__main__":
    записать(f"{ИМЯ}.xml", корневой_xml())
    записать(f"{ИМЯ}/Templates/ОсновнаяСхемаКомпоновкиДанных.xml", метаданные_макета())
    записать(f"{ИМЯ}/Templates/ОсновнаяСхемаКомпоновкиДанных/Ext/Template.xml", макет_скд())
