# -*- coding: utf-8 -*-
"""Template.xml СКД отчёта А_ПрогнозВиконанняССЗаПотребою (канон «План-факт»: внешний набор Детали + totalField).

    python gen_schema.py <путь к Templates/ОсновнаяСхемаКомпоновкиДанных/Ext/Template.xml>
Один генератор для обоих носителей (встроенный Reports/ и внешний _Rarzrabotki/Отчеты/).
"""
import sys
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(sys.argv[1])

DIMS = [
    ("Подразделение", "Підрозділ", "cfg:CatalogRef.СтруктураПредприятия"),
    ("ДокументПотребы", "Документ потреби", "cfg:DocumentRef.А_ПотребаЗавершенняПроекту"),
    ("СписокПотребы", "Список потреби", "str"),
    ("ЕтапНазва", "Етап", "str"),
    ("НазваЗагальна", "Загальна назва", "str"),
    ("ОбщееНазвание", "Загальна назва (довідник)", "cfg:CatalogRef.А_ОбщиеНазванияНоменклатуры"),
    ("ОзнакаСС", "Ознака СС", "str"),
    ("ВидРядка", "Вид рядка", "str"),
    ("Аналитика", "Аналітика", "str"),
    ("Документ", "Документ", ["cfg:DocumentRef.ПриобретениеТоваровУслуг", "cfg:CatalogRef.А_СтруктураСебестоимости",
                              "cfg:DocumentRef.А_ПотребаЗавершенняПроекту"]),
    ("Дома", "Будинки", "str"),
]
S = "Сумма"
RES = [  # имя, заголовок, выражение итога, точность
    ("ПланКол", "План СС, к-сть", "Сумма(ПланКол)", 3),
    ("ПланГрн", "План СС, грн", "Сумма(ПланГрн)", 2),
    ("КупленоКол", "Куплено, к-сть", "Сумма(КупленоКол)", 3),
    ("КупленоГрн", "Куплено, грн", "Сумма(КупленоГрн)", 2),
    ("ПослеГрн", "Куплено після списку, грн", "Сумма(ПослеГрн)", 2),
    ("ЗалишокСС", "Залишок за СС, грн", "Сумма(ЗалишокСС)", 2),
    ("ПотребаКол", "Потреба, к-сть (од. СС)", "Сумма(ПотребаКол)", 3),
    ("ПотребаГрн", "Потреба за списком, грн", "Сумма(ПотребаГрн)", 2),
    ("ЗалишокПотребы", "Залишок потреби, грн", "Сумма(ЗалишокПотребы)", 2),
    ("ПрогнозПотреба", "Прогноз за потребою, грн", "Сумма(КупленоГрн)+Сумма(ЗалишокПотребы)", 2),
    ("ЭкономияПотреба", "Економія за потребою, грн", "Сумма(ПланГрн)-Сумма(КупленоГрн)-Сумма(ЗалишокПотребы)", 2),
    ("ЭкономияПроцент", "Економія, %",
     "ВЫБОР КОГДА Сумма(ПланГрн) = 0 ТОГДА 0 ИНАЧЕ (Сумма(ПланГрн)-Сумма(КупленоГрн)-Сумма(ЗалишокПотребы)) / Сумма(ПланГрн) * 100 КОНЕЦ", 2),
    ("ПрогнозСС", "Прогноз за СС, грн", "Сумма(КупленоГрн)+Сумма(ЗалишокСС)", 2),
    ("ЭкономияСС", "Економія за СС, грн", "Сумма(ПланГрн)-Сумма(КупленоГрн)-Сумма(ЗалишокСС)", 2),
    ("РазницаМетодов", "Різниця методів, грн", "Сумма(ЗалишокСС)-Сумма(ЗалишокПотребы)", 2),
    ("НеДокупаем", "Не докуповуємо (поза списком), грн", "Сумма(НеДокупаем)", 2),
    ("ЦенаПлан", "Ціна СС", "ВЫБОР КОГДА Сумма(ПланКол) = 0 ТОГДА 0 ИНАЧЕ Сумма(ПланГрн) / Сумма(ПланКол) КОНЕЦ", 2),
    ("ЦенаПотребы", "Ціна потреби", "ВЫБОР КОГДА Сумма(ПотребаКол) = 0 ТОГДА 0 ИНАЧЕ Сумма(ПотребаГрн) / Сумма(ПотребаКол) КОНЕЦ", 2),
    ("ЦенаФакт", "Ціна факту", "ВЫБОР КОГДА Сумма(КупленоКол) = 0 ТОГДА 0 ИНАЧЕ Сумма(КупленоГрн) / Сумма(КупленоКол) КОНЕЦ", 2),
    ("ОтношениеЦен", "Ціна потреби / ціна СС",
     "ВЫБОР КОГДА Сумма(ПотребаКол) = 0 ИЛИ Сумма(ПланКол) = 0 ИЛИ Сумма(ПланГрн) = 0 ТОГДА 0 "
     "ИНАЧЕ Сумма(ПотребаГрн) * Сумма(ПланКол) / (Сумма(ПотребаКол) * Сумма(ПланГрн)) КОНЕЦ", 2),
]


def lstr(text, ind):
    t = escape(text)
    return (f'{ind}<v8:item>\n{ind}\t<v8:lang>ru</v8:lang>\n{ind}\t<v8:content>{t}</v8:content>\n{ind}</v8:item>\n'
            f'{ind}<v8:item>\n{ind}\t<v8:lang>uk</v8:lang>\n{ind}\t<v8:content>{t}</v8:content>\n{ind}</v8:item>\n')


def vtype(t, ind):
    if t == "str":
        return (f"{ind}<v8:Type>xs:string</v8:Type>\n{ind}<v8:StringQualifiers>\n{ind}\t<v8:Length>250</v8:Length>\n"
                f"{ind}\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{ind}</v8:StringQualifiers>\n")
    if isinstance(t, list):
        return "".join(f"{ind}<v8:Type>{x}</v8:Type>\n" for x in t)
    if isinstance(t, int):
        return (f"{ind}<v8:Type>xs:decimal</v8:Type>\n{ind}<v8:NumberQualifiers>\n{ind}\t<v8:Digits>18</v8:Digits>\n"
                f"{ind}\t<v8:FractionDigits>{t}</v8:FractionDigits>\n{ind}\t<v8:AllowedSign>Any</v8:AllowedSign>\n{ind}</v8:NumberQualifiers>\n")
    return f"{ind}<v8:Type>{t}</v8:Type>\n"


def field(name, title, typ, dim):
    s = (f'\t\t<field xsi:type="DataSetFieldField">\n\t\t\t<dataPath>{name}</dataPath>\n\t\t\t<field>{name}</field>\n'
         f'\t\t\t<title xsi:type="v8:LocalStringType">\n' + lstr(title, "\t\t\t\t") + "\t\t\t</title>\n")
    if dim:
        s += "\t\t\t<role>\n\t\t\t\t<dcscom:dimension>true</dcscom:dimension>\n\t\t\t</role>\n"
    s += "\t\t\t<valueType>\n" + vtype(typ, "\t\t\t\t") + "\t\t\t</valueType>\n\t\t</field>\n"
    return s


def gi(fieldname, ind):
    return (f'{ind}<dcsset:item xsi:type="dcsset:GroupItemField">\n{ind}\t<dcsset:field>{fieldname}</dcsset:field>\n'
            f"{ind}\t<dcsset:groupType>Items</dcsset:groupType>\n{ind}\t<dcsset:periodAdditionType>None</dcsset:periodAdditionType>\n"
            f'{ind}\t<dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>\n'
            f'{ind}\t<dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>\n{ind}</dcsset:item>\n')


def group(fields, ind, inner="", order=None, flt=None, hide_filter=False):
    s = f'{ind}<dcsset:item xsi:type="dcsset:StructureItemGroup">\n{ind}\t<dcsset:groupItems>\n'
    s += "".join(gi(f, ind + "\t\t") for f in fields)
    s += f"{ind}\t</dcsset:groupItems>\n"
    if flt:
        s += f"{ind}\t<dcsset:filter>\n{flt(ind + chr(9) * 2)}{ind}\t</dcsset:filter>\n"
    s += f"{ind}\t<dcsset:order>\n"
    if order:
        for f, d in order:
            s += (f'{ind}\t\t<dcsset:item xsi:type="dcsset:OrderItemField">\n{ind}\t\t\t<dcsset:field>{f}</dcsset:field>\n'
                  f"{ind}\t\t\t<dcsset:orderType>{d}</dcsset:orderType>\n{ind}\t\t</dcsset:item>\n")
    else:
        s += f'{ind}\t\t<dcsset:item xsi:type="dcsset:OrderItemAuto"/>\n'
    s += (f"{ind}\t</dcsset:order>\n{ind}\t<dcsset:selection>\n{ind}\t\t<dcsset:item xsi:type=\"dcsset:SelectedItemAuto\"/>\n"
          f"{ind}\t</dcsset:selection>\n")
    if hide_filter:
        s += (f"{ind}\t<dcsset:outputParameters>\n{ind}\t\t<dcscor:item xsi:type=\"dcsset:SettingsParameterValue\">\n"
              f"{ind}\t\t\t<dcscor:parameter>ВыводитьОтбор</dcscor:parameter>\n"
              f"{ind}\t\t\t<dcscor:value xsi:type=\"dcsset:DataCompositionTextOutputType\">DontOutput</dcscor:value>\n"
              f"{ind}\t\t</dcscor:item>\n{ind}\t</dcsset:outputParameters>\n")
    s += inner + f"{ind}</dcsset:item>\n"
    return s


def selection(fields, ind):
    return (f"{ind}<dcsset:selection>\n" + "".join(
        f'{ind}\t<dcsset:item xsi:type="dcsset:SelectedItemField">\n{ind}\t\t<dcsset:field>{f}</dcsset:field>\n{ind}\t</dcsset:item>\n'
        for f in fields) + f"{ind}</dcsset:selection>\n")


def ca_item(fields, flt_xml, params_xml, ind):
    s = f"{ind}<dcsset:item>\n{ind}\t<dcsset:selection>\n"
    s += "".join(f"{ind}\t\t<dcsset:item>\n{ind}\t\t\t<dcsset:field>{f}</dcsset:field>\n{ind}\t\t</dcsset:item>\n" for f in fields)
    s += f"{ind}\t</dcsset:selection>\n"
    s += (f"{ind}\t<dcsset:filter>\n{flt_xml}{ind}\t</dcsset:filter>\n" if flt_xml else f"{ind}\t<dcsset:filter/>\n")
    s += f"{ind}\t<dcsset:appearance>\n{params_xml}{ind}\t</dcsset:appearance>\n{ind}</dcsset:item>\n"
    return s


def cmp_item(fieldname, comp, value, ind):
    return (f'{ind}<dcsset:item xsi:type="dcsset:FilterItemComparison">\n{ind}\t<dcsset:left xsi:type="dcscor:Field">{fieldname}</dcsset:left>\n'
            f"{ind}\t<dcsset:comparisonType>{comp}</dcsset:comparisonType>\n"
            f'{ind}\t<dcsset:right xsi:type="xs:decimal">{value}</dcsset:right>\n{ind}</dcsset:item>\n')


def param_fmt(text, ind):
    return (f'{ind}<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n{ind}\t<dcscor:parameter>Формат</dcscor:parameter>\n'
            f'{ind}\t<dcscor:value xsi:type="v8:LocalStringType">\n' + lstr(text, ind + "\t\t") + f"{ind}\t</dcscor:value>\n{ind}</dcscor:item>\n")


def param_color(name, color, ind):
    return (f'{ind}<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n{ind}\t<dcscor:parameter>{name}</dcscor:parameter>\n'
            f'{ind}\t<dcscor:value xsi:type="v8ui:Color">{color}</dcscor:value>\n{ind}</dcscor:item>\n')


DATE_USER_ID = "7a1c9e40-5b2d-4f3e-9a61-3c8d2e7f1b05"
PODR_USER_ID = "4b6e2d18-9c3a-4e57-8f02-61d7a5c3e9b4"


VAR_N = [0]


def variant(name, title, sel, structure_xml, ca_xml, report_title):
    VAR_N[0] += 1
    date_id = DATE_USER_ID[:-1] + str(VAR_N[0])
    podr_id = PODR_USER_ID[:-1] + str(VAR_N[0])
    ind = "\t\t\t"
    s = (f"\t<settingsVariant>\n\t\t<dcsset:name>{name}</dcsset:name>\n\t\t<dcsset:presentation xsi:type=\"v8:LocalStringType\">\n"
         + lstr(title, "\t\t\t") + "\t\t</dcsset:presentation>\n"
         '\t\t<dcsset:settings xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
         'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows">\n')
    s += selection(sel, ind)
    s += (f'{ind}<dcsset:filter>\n{ind}\t<dcsset:item xsi:type="dcsset:FilterItemComparison">\n{ind}\t\t<dcsset:use>false</dcsset:use>\n'
          f'{ind}\t\t<dcsset:left xsi:type="dcscor:Field">Подразделение</dcsset:left>\n{ind}\t\t<dcsset:comparisonType>Equal</dcsset:comparisonType>\n'
          f"{ind}\t\t<dcsset:userSettingID>{podr_id}</dcsset:userSettingID>\n{ind}\t</dcsset:item>\n{ind}</dcsset:filter>\n")
    s += (f"{ind}<dcsset:dataParameters>\n{ind}\t<dcscor:item xsi:type=\"dcsset:SettingsParameterValue\">\n"
          f"{ind}\t\t<dcscor:parameter>ДатаОтчета</dcscor:parameter>\n"
          f'{ind}\t\t<dcscor:value xsi:type="xs:dateTime">0001-01-01T00:00:00</dcscor:value>\n'
          f"{ind}\t\t<dcsset:userSettingID>{date_id}</dcsset:userSettingID>\n{ind}\t</dcscor:item>\n{ind}</dcsset:dataParameters>\n")
    s += f"{ind}<dcsset:conditionalAppearance>\n{ca_xml}{ind}</dcsset:conditionalAppearance>\n"
    s += (f"{ind}<dcsset:outputParameters>\n"
          f'{ind}\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n{ind}\t\t<dcscor:parameter>МакетОформления</dcscor:parameter>\n'
          f'{ind}\t\t<dcscor:value xsi:type="xs:string">Море</dcscor:value>\n{ind}\t</dcscor:item>\n'
          f'{ind}\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n{ind}\t\t<dcscor:parameter>Заголовок</dcscor:parameter>\n'
          f'{ind}\t\t<dcscor:value xsi:type="v8:LocalStringType">\n' + lstr(report_title, ind + "\t\t\t") + f"{ind}\t\t</dcscor:value>\n{ind}\t</dcscor:item>\n"
          f'{ind}\t<dcscor:item xsi:type="dcsset:SettingsParameterValue">\n{ind}\t\t<dcscor:parameter>ВыводитьПараметрыДанных</dcscor:parameter>\n'
          f'{ind}\t\t<dcscor:value xsi:type="dcsset:DataCompositionTextOutputType">Output</dcscor:value>\n{ind}\t</dcscor:item>\n'
          f"{ind}</dcsset:outputParameters>\n")
    s += structure_xml + "\t\t</dcsset:settings>\n\t</settingsVariant>\n"
    return s


def build():
    x = ('<?xml version="1.0" encoding="UTF-8"?>\n'
         '<DataCompositionSchema xmlns="http://v8.1c.ru/8.1/data-composition-system/schema" '
         'xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" '
         'xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" xmlns:v8="http://v8.1c.ru/8.1/data/core" '
         'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
         'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config">\n'
         "\t<dataSource>\n\t\t<name>ИсточникДанных1</name>\n\t\t<dataSourceType>Local</dataSourceType>\n\t</dataSource>\n"
         '\t<dataSet xsi:type="DataSetObject">\n\t\t<name>Детали</name>\n')
    for n, t, typ in DIMS:
        x += field(n, t, typ, True)
    for n, t, _e, prec in RES:
        x += field(n, t, prec, False)
    x += "\t\t<dataSource>ИсточникДанных1</dataSource>\n\t\t<objectName>Детали</objectName>\n\t</dataSet>\n"
    for n, _t, e, _p in RES:
        x += f"\t<totalField>\n\t\t<dataPath>{n}</dataPath>\n\t\t<expression>{escape(e)}</expression>\n\t</totalField>\n"
    x += ("\t<parameter>\n\t\t<name>ДатаОтчета</name>\n\t\t<title xsi:type=\"v8:LocalStringType\">\n" + lstr("Дата звіту", "\t\t\t") +
          "\t\t</title>\n\t\t<valueType>\n\t\t\t<v8:Type>xs:dateTime</v8:Type>\n\t\t\t<v8:DateQualifiers>\n"
          "\t\t\t\t<v8:DateFractions>Date</v8:DateFractions>\n\t\t\t</v8:DateQualifiers>\n\t\t</valueType>\n"
          '\t\t<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>\n\t\t<useRestriction>false</useRestriction>\n\t</parameter>\n')

    i3 = "\t\t\t"
    fmt1 = ca_item(["ЭкономияПроцент"], None, param_fmt("ЧДЦ=1", i3 + "\t\t"), i3 + "\t")
    red = ca_item(["ЭкономияПотреба", "ЭкономияПроцент"], cmp_item("ЭкономияПотреба", "Less", "0", i3 + "\t\t\t"),
                  param_color("ЦветТекста", "style:НедопустимоеЗначениеЦвет" if False else "#C00000", i3 + "\t\t"), i3 + "\t")
    ratio = ca_item(["ЦенаПотребы", "ОтношениеЦен"],
                    (f'{i3}\t\t\t<dcsset:item xsi:type="dcsset:FilterItemGroup">\n{i3}\t\t\t\t<dcsset:groupType>OrGroup</dcsset:groupType>\n'
                     + cmp_item("ОтношениеЦен", "Greater", "1.5", i3 + "\t\t\t\t")
                     + f'{i3}\t\t\t\t<dcsset:item xsi:type="dcsset:FilterItemGroup">\n{i3}\t\t\t\t\t<dcsset:groupType>AndGroup</dcsset:groupType>\n'
                     + cmp_item("ОтношениеЦен", "Less", "0.67", i3 + "\t\t\t\t\t")
                     + cmp_item("ОтношениеЦен", "NotEqual", "0", i3 + "\t\t\t\t\t")
                     + f"{i3}\t\t\t\t</dcsset:item>\n{i3}\t\t\t</dcsset:item>\n"),
                    param_color("ЦветФона", "#FFE0B0", i3 + "\t\t"), i3 + "\t")
    fmt2 = ca_item(["ОтношениеЦен"], None, param_fmt("ЧДЦ=2", i3 + "\t\t"), i3 + "\t")

    main_sel = ["ПланГрн", "КупленоГрн", "ПотребаГрн", "ПослеГрн", "ЗалишокПотребы", "ПрогнозПотреба", "ЭкономияПотреба",
                "ЭкономияПроцент", "ПрогнозСС", "ЭкономияСС", "РазницаМетодов", "НеДокупаем"]
    st_main = group(["Подразделение", "СписокПотребы"], i3,
                    group(["ЕтапНазва"], i3 + "\t", group(["НазваЗагальна"], i3 + "\t\t", order=[("НазваЗагальна", "Asc")])))
    x += variant("ПрогнозЗаПотребою", "Прогноз за потребою", main_sel, st_main, fmt1 + red, "Прогноз виконання СС за потребою")

    det_sel = ["ПланКол", "ПланГрн", "ЦенаПлан", "КупленоКол", "КупленоГрн", "ЦенаФакт", "ПотребаКол", "ПотребаГрн", "ЦенаПотребы",
               "ОтношениеЦен", "ПослеГрн", "ЗалишокСС", "ЗалишокПотребы", "ПрогнозПотреба", "ЭкономияПотреба", "ЭкономияПроцент",
               "ПрогнозСС", "ЭкономияСС", "НеДокупаем"]
    st_det = group(["Подразделение", "СписокПотребы"], i3,
                   group(["ЕтапНазва"], i3 + "\t",
                         group(["НазваЗагальна"], i3 + "\t\t",
                               group(["ОзнакаСС"], i3 + "\t\t\t",
                                     group(["ВидРядка"], i3 + "\t\t\t\t",
                                           group(["Аналитика", "Дома"], i3 + "\t\t\t\t\t",
                                                 group(["Документ"], i3 + "\t\t\t\t\t\t"))),
                                     ), order=[("НазваЗагальна", "Asc")])))
    x += variant("ПрогнозЗаПотребоюДетально", "Прогноз за потребою — детально", det_sel, st_det, fmt1 + red + ratio + fmt2,
                 "Прогноз виконання СС за потребою — детально")

    nd_sel = ["ПланГрн", "КупленоГрн", "ЗалишокСС", "ПотребаГрн", "НеДокупаем"]
    st_nd = group(["Подразделение", "СписокПотребы"], i3,
                  group(["НазваЗагальна"], i3 + "\t", order=[("НеДокупаем", "Desc")],
                        flt=lambda ind: cmp_item("НеДокупаем", "Greater", "0", ind), hide_filter=True))
    x += variant("НеДокуповуемо", "Не докуповуємо (поза списком)", nd_sel, st_nd, "", "Не докуповуємо: залишок за СС поза списком директора")
    x += "</DataCompositionSchema>"
    OUT.write_bytes(b"\xef\xbb\xbf" + x.encode("utf-8"))


build()
print("ok", OUT)
