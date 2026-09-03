# -*- coding: utf-8 -*-
"""Генератор СКД отчёта «Табели Казны и ЗУП», версия 2: один набор «Свод» (строка = сотрудник × организация ЗУП),
ресурсы со своим выражением на уровне сотрудника (Максимум(…Сотр)), организации (значение строки) и общего итога
(Сумма(…Итог) по первым строкам); стеки «дни / часы» строками; наборы СводДок (табели) и ПоДням."""
import sys, uuid
sys.stdout.reconfigure(encoding="utf-8")
OUT = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\timesheet-kazna-zup-integration-465382\_Rarzrabotki\Отчеты\ОтчетПоДаннымТабелейКазныиБазыЗУП\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"
NS = 'xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config"'


def T(s, n):
    return "\t" * n + s


def lstr(text, n):
    return "\n".join([T('<title xsi:type="v8:LocalStringType">', n), T("<v8:item>", n + 1), T("<v8:lang>ru</v8:lang>", n + 2),
                      T("<v8:content>%s</v8:content>" % text, n + 2), T("</v8:item>", n + 1), T("</title>", n)])


def vtype(kind, n):
    if kind == "num":
        return "\n".join([T("<valueType>", n), T("<v8:Type>xs:decimal</v8:Type>", n + 1), T("<v8:NumberQualifiers>", n + 1),
                          T("<v8:Digits>15</v8:Digits>", n + 2), T("<v8:FractionDigits>2</v8:FractionDigits>", n + 2), T("<v8:AllowedSign>Any</v8:AllowedSign>", n + 2),
                          T("</v8:NumberQualifiers>", n + 1), T("</valueType>", n)])
    if kind == "bool":
        return "\n".join([T("<valueType>", n), T("<v8:Type>xs:boolean</v8:Type>", n + 1), T("</valueType>", n)])
    if kind == "date":
        return "\n".join([T("<valueType>", n), T("<v8:Type>xs:dateTime</v8:Type>", n + 1), T("<v8:DateQualifiers>", n + 1),
                          T("<v8:DateFractions>Date</v8:DateFractions>", n + 2), T("</v8:DateQualifiers>", n + 1), T("</valueType>", n)])
    if kind == "ref":
        return "\n".join([T("<valueType>", n), T('<v8:Type %s>d5p1:CatalogRef.Сотрудники</v8:Type>' % NS, n + 1), T("</valueType>", n)])
    if kind == "docref":
        return "\n".join([T("<valueType>", n), T('<v8:Type %s>d5p1:DocumentRef.ТабельУчетаРабочегоВремени</v8:Type>' % NS, n + 1), T("</valueType>", n)])
    return "\n".join([T("<valueType>", n), T("<v8:Type>xs:string</v8:Type>", n + 1), T("<v8:StringQualifiers>", n + 1),
                      T("<v8:Length>500</v8:Length>", n + 2), T("<v8:AllowedLength>Variable</v8:AllowedLength>", n + 2),
                      T("</v8:StringQualifiers>", n + 1), T("</valueType>", n)])


def field(name, kind, title, n=2, src=None):
    out = [T('<field xsi:type="DataSetFieldField">', n), T("<dataPath>%s</dataPath>" % name, n + 1), T("<field>%s</field>" % (src or name), n + 1),
           lstr(title, n + 1), vtype(kind, n + 1)]
    if kind == "num":
        out += [T("<appearance>", n + 1), T('<dcscor:item xsi:type="dcsset:SettingsParameterValue">', n + 2), T("<dcscor:parameter>Формат</dcscor:parameter>", n + 3),
                T('<dcscor:value xsi:type="xs:string">ЧДЦ=0; ЧН=</dcscor:value>', n + 3), T("</dcscor:item>", n + 2), T("</appearance>", n + 1)]
    out.append(T("</field>", n))
    return "\n".join(out)


# ---------------- набор «Свод»: строка = сотрудник × организация ЗУП ----------------
SVOD = [
    # уровень сотрудника
    ("ФИО", "str", "Сотрудник"),
    ("Сотрудник", "ref", "Сотрудник (ссылка)"),
    ("ИНН", "str", "ИНН"),
    ("КлючСвода", "str", "Ключ сотрудника"),
    ("Трудоустроен", "bool", "Трудоустроен (Казна), отбор"),
    ("ТрудоустроенТекст", "str", "Трудоустроен (Казна)"),
    ("АктуаленВЗУП", "bool", "Актуален в ЗУП, отбор"),
    ("АктуаленВЗУПТекст", "str", "Актуален в ЗУП"),
    ("Контроль", "str", "Контроль"),
    ("РасхождениеДней", "num", "Расхождение дней (Казна − ЗУП)"),
    ("КодОкраски", "num", "Код окраски"),
    ("ПервыйОрг", "bool", "Первая строка сотрудника"),
    ("ЕстьКонтроль", "bool", "Есть контроль"),
    ("ЕстьРасхождениеДней", "bool", "Есть расхождение дней"),
    ("НетФлагаТрудоустроен", "bool", "Нет флага Трудоустроен"),
    ("УволенВЗУП", "bool", "Трудоустроен, но уволен в ЗУП"),
    ("НеПротабелированВЗУП", "bool", "В ЗУП нет рабочих дней"),
    ("НетНачисленийЗУП", "bool", "Нет начислений в ЗУП"),
    ("БольничныйНеВЗУП", "bool", "Больничный не в ЗУП"),
    ("ОтпускНеВЗУП", "bool", "Отпуск не в ЗУП"),
    ("ЗаСвойСчетНеВЗУП", "bool", "За свой счёт не в ЗУП"),
    ("НетИНН", "bool", "Нет ИНН"),
    ("ДубльИНН", "bool", "Дубль ИНН в Казне"),
    ("НетТабеляКазны", "bool", "Нет табеля Казны"),
    ("ЕстьТабельКазны", "bool", "Есть табель Казны"),
    ("ОрганизацийЗУП", "num", "Организаций (ЗУП)"),
    ("ФИОЗУП", "str", "ФИО (ЗУП)"),
    ("ОрганизацииКазны", "str", "Организации (Казна)"),
    ("ПодразделенияКазны", "str", "Подразделения (Казна)"),
    # значения сотрудника (для уровня сотрудника и варианта «Подробно»)
    ("ДатаПриемаСотр", "date", "Приём (ЗУП, сотрудник)"),
    ("ДатаУвольненияСотр", "date", "Увольнение (ЗУП, сотрудник)"),
    ("ДниКазнаСотр", "num", "Дни Казна (Р+М+К)"),
    ("ЧасыКазнаСотр", "num", "Часы Казна (Р+М)"),
    ("ДниЗУПСотр", "num", "Дни ЗУП (работа, уник. даты)"),
    ("ЧасыЗУПСотр", "num", "Часы ЗУП (работа)"),
    ("ДниОтпускСотр", "num", "Дни отпуск (ЗУП)"),
    ("ДниБольничныйСотр", "num", "Дни больничный (ЗУП)"),
    ("ДниБезОплатыСотр", "num", "Дни без оплаты (ЗУП)"),
    ("НачисленоСотр", "num", "Начислено (ЗУП)"),
    ("КазнаДниЧасыСотр", "str", "Казна дни/часы (сотрудник)"),
    ("ЗУПДниЧасыСотр", "str", "ЗУП дни/часы (сотрудник)"),
    ("ДниР", "num", "Дни Р (по ставке)"),
    ("ЧасыР", "num", "Часы Р (по ставке)"),
    ("ДниМ", "num", "Дни М (только буква М)"),
    ("ЧасыМ", "num", "Часы М (только буква М)"),
    ("ДниОфициальные", "num", "Дни официальные (Р+М)"),
    ("ДниО", "num", "Дни О (отпуск)"),
    ("ДниБ", "num", "Дни Б (больничный)"),
    ("ДниС", "num", "Дни С (за свой счёт)"),
    ("ДниК", "num", "Дни К (командировка)"),
    ("ДниН", "num", "Дни Н (неявка)"),
    ("ДниВ", "num", "Дни В (выходной)"),
    ("ДниОшибка", "num", "Дни ? (ошибка ввода)"),
    ("ЧасыПлан", "num", "Часы план (Казна)"),
    ("ДниКомандировкаЗУП", "num", "Дни командировка (ЗУП)"),
    ("ДниНеявкиЗУП", "num", "Дни неявки (ЗУП)"),
    ("ДниПрочиеЗУП", "num", "Дни прочие (ЗУП)"),
    ("ОтработаноДнейЗУП", "num", "Отработано дней (ЗУП, начисления)"),
    ("НормаДнейЗУП", "num", "Норма дней (ЗУП)"),
    # итоговые значения (только в первой строке сотрудника)
    ("ДниКазнаИтог", "num", "Дни Казна (для итога)"),
    ("ЧасыКазнаИтог", "num", "Часы Казна (для итога)"),
    ("ДниЗУПИтог", "num", "Дни ЗУП (для итога)"),
    ("ЧасыЗУПИтог", "num", "Часы ЗУП (для итога)"),
    ("НачисленоИтог", "num", "Начислено (для итога)"),
    # общие колонки: у сотрудника через ресурс Максимум(…Сотр), у организации значение строки
    ("Организация", "str", "Организация ЗУП"),
    ("Договор", "str", "Договор (ЗУП)"),
    ("ИсточникДатыПриема", "str", "Источник даты приёма"),
    ("ДатаПриема", "date", "Приём (ЗУП)"),
    ("ДатаУвольнения", "date", "Увольнение (ЗУП)"),
    ("ДниКазна", "num", "Дни Казна"),
    ("ЧасыКазна", "num", "Часы Казна"),
    ("ДниЗУП", "num", "Дни ЗУП (работа / норма орг.)"),
    ("ЧасыЗУП", "num", "Часы ЗУП (работа / норма орг.)"),
    ("КазнаДниЧасы", "str", "Казна: дни / часы"),
    ("ЗУПДниЧасы", "str", "ЗУП: дни / часы (по орг. — норма)"),
    ("ДниОтпуск", "num", "Отпуск ЗУП (дн)"),
    ("ДниБольничный", "num", "Больничный ЗУП (дн)"),
    ("ДниБезОплаты", "num", "Без оплаты ЗУП (дн)"),
    ("Начислено", "num", "Начислено (ЗУП)"),
]
SVODDOK = [
    ("ФИОДок", "str", "Сотрудник"),
    ("ИННДок", "str", "ИНН"),
    ("КонтрольДок", "str", "Контроль"),
    ("КазнаДниЧасыДок", "str", "Казна: дни / часы"),
    ("ЗУПДниЧасыДок", "str", "ЗУП: дни / часы"),
    ("ТабельКазныДок", "docref", "Табель Казны"),
    ("ОрганизацияДок", "str", "Организация табеля"),
    ("ПодразделениеДок", "str", "Подразделение табеля"),
    ("ДниРДок", "num", "Дни Р (табель)"),
    ("ЧасыРДок", "num", "Часы Р (табель)"),
    ("ДниМДок", "num", "Дни М (табель)"),
    ("ЧасыМДок", "num", "Часы М (табель)"),
    ("ДниОфицДок", "num", "Дни официальные (табель)"),
    ("ЧасыОфицДок", "num", "Часы официальные (табель)"),
    ("ДниОДок", "num", "Дни О (табель)"),
    ("ДниБДок", "num", "Дни Б (табель)"),
    ("ДниСДок", "num", "Дни С (табель)"),
    ("ДниКДок", "num", "Дни К (табель)"),
    ("ДниНДок", "num", "Дни Н (табель)"),
    ("ДниВДок", "num", "Дни В (табель)"),
]
PODNYAM = [
    ("СотрудникД", "ref", "Сотрудник (ссылка)"),
    ("ФИОД", "str", "Сотрудник"),
    ("ИННД", "str", "ИНН"),
    ("Дата", "date", "Дата"),
    ("ЗначениеКазны", "str", "Казна"),
    ("СостояниеКазны", "str", "Состояние (Казна)"),
    ("ЧасыКазны", "num", "Часы официальные (Казна)"),
    ("КодЗУП", "str", "Код (ЗУП)"),
    ("ОрганизацияЗУПД", "str", "Организация (ЗУП)"),
    ("ЧасыЗУПД", "num", "Часы (ЗУП)"),
    ("ДнейЗУП", "num", "Дней (ЗУП)"),
    ("Расхождение", "str", "Расхождение"),
    ("ТабельКазныДень", "docref", "Табель Казны"),
]
DOC_EMP = ["ФИОДок", "ИННДок", "КонтрольДок", "КазнаДниЧасыДок", "ЗУПДниЧасыДок"]
DOC_SHOW = ["ТабельКазныДок", "ОрганизацияДок", "ПодразделениеДок", "ДниОфицДок", "ЧасыОфицДок", "ДниРДок", "ДниМДок",
            "ДниОДок", "ДниБДок", "ДниСДок", "ДниКДок", "ДниНДок"]

# ресурсы: (dataPath, {группа: выражение})
EMP = "ФИО"
ORG = "Организация"
TOT = "ОбщийИтог"
TOTALS = []
for f in ["ИНН", "ТрудоустроенТекст", "АктуаленВЗУПТекст", "Контроль", "ФИОЗУП"]:
    TOTALS.append((f, {EMP: "Максимум(%s)" % f, ORG: '""'}))
for f in ["РасхождениеДней", "ОрганизацийЗУП"]:
    TOTALS.append((f, {EMP: "Максимум(%s)" % f, ORG: "0"}))
TOTALS.append(("КодОкраски", {EMP: "Максимум(КодОкраски)"}))
TOTALS += [
    ("ДатаПриема", {EMP: "Максимум(ДатаПриемаСотр)", ORG: "Максимум(ДатаПриема)"}),
    ("ДатаУвольнения", {EMP: "Максимум(ДатаУвольненияСотр)", ORG: "Максимум(ДатаУвольнения)"}),
    ("Договор", {ORG: "Максимум(Договор)"}),
    ("ИсточникДатыПриема", {ORG: "Максимум(ИсточникДатыПриема)"}),
    ("ДниКазна", {EMP: "Максимум(ДниКазнаСотр)", ORG: "0", TOT: "Сумма(ДниКазнаИтог)"}),
    ("ЧасыКазна", {EMP: "Максимум(ЧасыКазнаСотр)", ORG: "0", TOT: "Сумма(ЧасыКазнаИтог)"}),
    ("ДниЗУП", {EMP: "Максимум(ДниЗУПСотр)", ORG: "Сумма(ДниЗУП)", TOT: "Сумма(ДниЗУПИтог)"}),
    ("ЧасыЗУП", {EMP: "Максимум(ЧасыЗУПСотр)", ORG: "Сумма(ЧасыЗУП)", TOT: "Сумма(ЧасыЗУПИтог)"}),
    ("КазнаДниЧасы", {EMP: "Максимум(КазнаДниЧасыСотр)", ORG: '""', TOT: 'Формат(Сумма(ДниКазнаИтог), "ЧДЦ=0; ЧГ=0") + " / " + Формат(Сумма(ЧасыКазнаИтог), "ЧДЦ=0; ЧГ=0")'}),
    ("ЗУПДниЧасы", {EMP: "Максимум(ЗУПДниЧасыСотр)", ORG: "Максимум(ЗУПДниЧасы)", TOT: 'Формат(Сумма(ДниЗУПИтог), "ЧДЦ=0; ЧГ=0") + " / " + Формат(Сумма(ЧасыЗУПИтог), "ЧДЦ=0; ЧГ=0")'}),
    ("ДниОтпуск", {EMP: "Максимум(ДниОтпускСотр)", ORG: "Сумма(ДниОтпуск)"}),
    ("ДниБольничный", {EMP: "Максимум(ДниБольничныйСотр)", ORG: "Сумма(ДниБольничный)"}),
    ("ДниБезОплаты", {EMP: "Максимум(ДниБезОплатыСотр)", ORG: "Сумма(ДниБезОплаты)"}),
    ("Начислено", {EMP: "Максимум(НачисленоСотр)", ORG: "Сумма(Начислено)", TOT: "Сумма(НачисленоИтог)"}),
]
# итоги для плоского варианта «Подробно» (по первым строкам)
FLAT_TOTALS = [("ДниКазнаСотр", "Сумма(ДниКазнаИтог)"), ("ЧасыКазнаСотр", "Сумма(ЧасыКазнаИтог)"), ("ДниЗУПСотр", "Сумма(ДниЗУПИтог)"),
               ("ЧасыЗУПСотр", "Сумма(ЧасыЗУПИтог)"), ("НачисленоСотр", "Сумма(НачисленоИтог)")]
DOC = "ТабельКазныДок"
for f, k, _ in SVODDOK:
    if k == "num":
        TOTALS.append((f, {"ФИОДок": "Сумма(%s)" % f, DOC: "Сумма(%s)" % f, TOT: "Сумма(%s)" % f}))
for f in ["ОрганизацияДок", "ПодразделениеДок"]:
    TOTALS.append((f, {DOC: "Максимум(%s)" % f}))
for f in DOC_EMP[1:]:
    TOTALS.append((f, {"ФИОДок": "Максимум(%s)" % f, DOC: '""'}))
for f, k, _ in PODNYAM:
    if k == "num":
        TOTALS.append((f, {TOT: "Сумма(%s)" % f, "ФИОД": "Сумма(%s)" % f}))


def dataset(name, fields):
    out = [T('<dataSet xsi:type="DataSetObject">', 1), T("<name>%s</name>" % name, 2)]
    out += [field(f, k, t, 2, "ЧасыЗУП" if f == "ЧасыЗУПД" else None) for f, k, t in fields]
    out += [T("<dataSource>ИсточникДанных1</dataSource>", 2), T("<objectName>%s</objectName>" % name, 2), T("</dataSet>", 1)]
    return "\n".join(out)


def dataset_link(dst, src_expr, dst_expr):
    return "\n".join([T("<dataSetLink>", 1), T("<sourceDataSet>Свод</sourceDataSet>", 2), T("<destinationDataSet>%s</destinationDataSet>" % dst, 2),
                      T("<sourceExpression>%s</sourceExpression>" % src_expr, 2), T("<destinationExpression>%s</destinationExpression>" % dst_expr, 2),
                      T("<parameterListAllowed>false</parameterListAllowed>", 2), T("</dataSetLink>", 1)])


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def total(path, expr, groups):
    out = [T("<totalField>", 1), T("<dataPath>%s</dataPath>" % path, 2), T("<expression>%s</expression>" % esc(expr), 2)]
    out += [T("<group>%s</group>" % g, 2) for g in groups]
    out.append(T("</totalField>", 1))
    return "\n".join(out)


def all_totals():
    out = []
    for path, m in TOTALS:
        for g, expr in m.items():
            out.append(total(path, expr, [g]))
    for path, expr in FLAT_TOTALS:
        out.append(total(path, expr, [TOT]))
    return "\n".join(out)


def param_period():
    return "\n".join([T("<parameter>", 1), T("<name>Период</name>", 2), lstr("Период", 2),
                      T("<valueType>", 2), T("<v8:Type>v8:StandardPeriod</v8:Type>", 3), T("</valueType>", 2),
                      T('<value xsi:type="v8:StandardPeriod">', 2), T('<v8:variant xsi:type="v8:StandardPeriodVariant">LastMonth</v8:variant>', 3),
                      T("<v8:startDate>0001-01-01T00:00:00</v8:startDate>", 3), T("<v8:endDate>0001-01-01T00:00:00</v8:endDate>", 3), T("</value>", 2),
                      T("<useRestriction>false</useRestriction>", 2), T("</parameter>", 1)])


def sel(fields, n):
    return "\n".join([T("<dcsset:selection>", n)] + ["\n".join([T('<dcsset:item xsi:type="dcsset:SelectedItemField">', n + 1), T("<dcsset:field>%s</dcsset:field>" % f, n + 2), T("</dcsset:item>", n + 1)]) for f in fields] + [T("</dcsset:selection>", n)])


def filt(fieldname, use, cmp, right_xml, n, with_id=True):
    out = [T('<dcsset:item xsi:type="dcsset:FilterItemComparison">', n)]
    if not use:
        out.append(T("<dcsset:use>false</dcsset:use>", n + 1))
    out += [T('<dcsset:left xsi:type="dcscor:Field">%s</dcsset:left>' % fieldname, n + 1), T("<dcsset:comparisonType>%s</dcsset:comparisonType>" % cmp, n + 1)]
    if right_xml:
        out.append(T(right_xml, n + 1))
    if with_id:
        out.append(T("<dcsset:userSettingID>%s</dcsset:userSettingID>" % uuid.uuid4(), n + 1))
    out.append(T("</dcsset:item>", n))
    return "\n".join(out)


def filt_bool(fieldname, use, value, n, with_id=True):
    return filt(fieldname, use, "Equal", '<dcsset:right xsi:type="xs:boolean">%s</dcsset:right>' % ("true" if value else "false"), n, with_id)


def filt_num(fieldname, use, cmp, value, n, with_id=False):
    return filt(fieldname, use, cmp, '<dcsset:right xsi:type="xs:decimal">%s</dcsset:right>' % value, n, with_id)


def filt_str(fieldname, use, cmp, value, n, with_id=False):
    return filt(fieldname, use, cmp, '<dcsset:right xsi:type="xs:string">%s</dcsset:right>' % value, n, with_id)


def order(fields, n):
    return "\n".join([T("<dcsset:order>", n)] + ["\n".join([T('<dcsset:item xsi:type="dcsset:OrderItemField">', n + 1), T("<dcsset:field>%s</dcsset:field>" % f, n + 2), T("<dcsset:orderType>Asc</dcsset:orderType>", n + 2), T("</dcsset:item>", n + 1)]) for f in fields] + [T("</dcsset:order>", n)])


def cond_app(items, n):
    out = [T("<dcsset:conditionalAppearance>", n)]
    for filter_xml, color, pres in items:
        out += [T("<dcsset:item>", n + 1), T("<dcsset:selection/>", n + 2), T("<dcsset:filter>", n + 2), filter_xml, T("</dcsset:filter>", n + 2),
                T("<dcsset:appearance>", n + 2), T('<dcscor:item xsi:type="dcsset:SettingsParameterValue">', n + 3), T("<dcscor:parameter>ЦветФона</dcscor:parameter>", n + 4),
                T('<dcscor:value xsi:type="v8ui:Color">%s</dcscor:value>' % color, n + 4), T("</dcscor:item>", n + 3), T("</dcsset:appearance>", n + 2),
                T('<dcsset:presentation xsi:type="xs:string">%s</dcsset:presentation>' % pres, n + 2), T("</dcsset:item>", n + 1)]
    out.append(T("</dcsset:conditionalAppearance>", n))
    return "\n".join(out)


def data_params(n):
    return "\n".join([T("<dcsset:dataParameters>", n), T('<dcscor:item xsi:type="dcsset:SettingsParameterValue">', n + 1), T("<dcscor:parameter>Период</dcscor:parameter>", n + 2),
                      T('<dcscor:value xsi:type="v8:StandardPeriod">', n + 2), T('<v8:variant xsi:type="v8:StandardPeriodVariant">LastMonth</v8:variant>', n + 3),
                      T("<v8:startDate>0001-01-01T00:00:00</v8:startDate>", n + 3), T("<v8:endDate>0001-01-01T00:00:00</v8:endDate>", n + 3), T("</dcscor:value>", n + 2),
                      T("<dcsset:userSettingID>%s</dcsset:userSettingID>" % uuid.uuid4(), n + 2), T("</dcscor:item>", n + 1), T("</dcsset:dataParameters>", n)])


def group_detail(n):
    return "\n".join([T('<dcsset:item xsi:type="dcsset:StructureItemGroup">', n), T("<dcsset:order>", n + 1), T('<dcsset:item xsi:type="dcsset:OrderItemAuto"/>', n + 2), T("</dcsset:order>", n + 1),
                      T("<dcsset:selection>", n + 1), T('<dcsset:item xsi:type="dcsset:SelectedItemAuto"/>', n + 2), T("</dcsset:selection>", n + 1), T("</dcsset:item>", n)])


def out_params(n, hide_header=False):
    items = [("ВыводитьОтбор", "DontOutput")]
    out = [T("<dcsset:outputParameters>", n)]
    for name, val in items:
        out += [T('<dcscor:item xsi:type="dcsset:SettingsParameterValue">', n + 1), T("<dcscor:parameter>%s</dcscor:parameter>" % name, n + 2),
                T('<dcscor:value xsi:type="dcsset:DataCompositionTextOutputType">%s</dcscor:value>' % val, n + 2), T("</dcscor:item>", n + 1)]
    out.append(T("</dcsset:outputParameters>", n))
    return chr(10).join(out)


def group_fields(fields, n, inner=None, show=None, filters=None):
    items = []
    for f in fields:
        items += [T('<dcsset:item xsi:type="dcsset:GroupItemField">', n + 2), T("<dcsset:field>%s</dcsset:field>" % f, n + 3), T("<dcsset:groupType>Items</dcsset:groupType>", n + 3),
                  T("<dcsset:periodAdditionType>None</dcsset:periodAdditionType>", n + 3), T('<dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>', n + 3),
                  T('<dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>', n + 3), T("</dcsset:item>", n + 2)]
    out = [T('<dcsset:item xsi:type="dcsset:StructureItemGroup">', n), T("<dcsset:groupItems>", n + 1)] + items + [T("</dcsset:groupItems>", n + 1)]
    if filters:
        out += [T("<dcsset:filter>", n + 1)] + filters + [T("</dcsset:filter>", n + 1)]
    out += [T("<dcsset:order>", n + 1), T('<dcsset:item xsi:type="dcsset:OrderItemAuto"/>', n + 2), T("</dcsset:order>", n + 1)]
    out.append(sel(show, n + 1) if show else "\n".join([T("<dcsset:selection>", n + 1), T('<dcsset:item xsi:type="dcsset:SelectedItemAuto"/>', n + 2), T("</dcsset:selection>", n + 1)]))
    if filters:
        out.append(out_params(n + 1))
    if inner:
        out.append(inner)
    out.append(T("</dcsset:item>", n))
    return "\n".join(out)


def variant(name, pres, selection, filters, orders, cond, structure):
    return "\n".join([T("<settingsVariant>", 1), T("<dcsset:name>%s</dcsset:name>" % name, 2),
                      T('<dcsset:presentation xsi:type="v8:LocalStringType">', 2), T("<v8:item>", 3), T("<v8:lang>ru</v8:lang>", 4), T("<v8:content>%s</v8:content>" % pres, 4), T("</v8:item>", 3), T("</dcsset:presentation>", 2),
                      T("<dcsset:settings>", 2), sel(selection, 3), T("<dcsset:filter>", 3), "\n".join(filters), T("</dcsset:filter>", 3), order(orders, 3), cond_app(cond, 3), data_params(3), out_params(3), structure, T("</dcsset:settings>", 2), T("</settingsVariant>", 1)])


MAIN_SHOW = ["ФИО", "ИНН", "ТрудоустроенТекст", "АктуаленВЗУПТекст", "ДатаПриема", "ДатаУвольнения", "РасхождениеДней", "Контроль",
             "КазнаДниЧасы", "ЗУПДниЧасы", "Договор", "ДниОтпуск", "ДниБольничный", "ДниБезОплаты", "Начислено"]
DETAIL_SHOW = ["ФИО", "ИНН", "ТрудоустроенТекст", "АктуаленВЗУПТекст", "ДатаПриемаСотр", "ДатаУвольненияСотр", "РасхождениеДней", "Контроль",
               "ДниКазнаСотр", "ЧасыКазнаСотр", "ДниЗУПСотр", "ЧасыЗУПСотр", "ДниОфициальные", "ДниР", "ЧасыР", "ДниМ", "ЧасыМ",
               "ДниО", "ДниБ", "ДниС", "ДниК", "ДниН", "ДниВ", "ЧасыПлан", "ОрганизацийЗУП", "ОрганизацииКазны",
               "ДниОтпускСотр", "ДниБольничныйСотр", "ДниБезОплатыСотр", "ДниКомандировкаЗУП", "ДниНеявкиЗУП", "НачисленоСотр", "ОтработаноДнейЗУП", "НормаДнейЗУП"]
COND = [(filt_num("КодОкраски", True, "GreaterOrEqual", 1, 4, False), "#FFF2CC", "Есть контроль"),
        (filt_num("КодОкраски", True, "GreaterOrEqual", 2, 4, False), "#FFC7CE", "Нет флага Трудоустроен / уволен / нет рабочих дней в ЗУП")]
USER_FILTERS = [filt_bool("ЕстьКонтроль", False, True, 4), filt_bool("Трудоустроен", False, True, 4), filt_bool("НетТабеляКазны", False, False, 4), filt_bool("ЕстьРасхождениеДней", False, True, 4)]
ORG_LEVEL = group_fields([ORG], 4, None, None, [filt("Организация", True, "Filled", None, 5, False)])
MAIN_STRUCT = group_fields([EMP], 3, ORG_LEVEL)
v1 = variant("СводноПоСотрудникам", "Сводно по сотрудникам (с организациями ЗУП)", MAIN_SHOW, USER_FILTERS, ["ФИО"], COND, MAIN_STRUCT)
v2 = variant("Контроль", "Контроль (только с расхождениями)", MAIN_SHOW,
             [filt_bool("ЕстьКонтроль", True, True, 4), filt_bool("Трудоустроен", False, True, 4), filt_bool("НетТабеляКазны", False, False, 4)],
             ["ФИО"], COND, MAIN_STRUCT)
v3 = variant("ПоТабелямКазны", "Сотрудник → табели Казны", DOC_EMP + DOC_SHOW, [filt_str("ИННДок", False, "Equal", "", 4, True)],
             ["ФИОДок"], [], group_fields(["ФИОДок"], 3, group_fields([DOC], 4)))
v4 = variant("Подробно", "Подробно (все колонки)", DETAIL_SHOW, [filt_bool("ПервыйОрг", True, True, 4, False)] + USER_FILTERS, ["ФИО"], COND, group_detail(3))
v5 = variant("ПоДням", "По дням", ["ФИОД", "ИННД", "Дата", "ТабельКазныДень", "ЗначениеКазны", "СостояниеКазны", "ЧасыКазны", "КодЗУП", "ОрганизацияЗУПД", "ЧасыЗУПД", "ДнейЗУП", "Расхождение"],
             [filt_str("Расхождение", False, "NotEqual", "", 4, True)], ["ФИОД", "Дата"],
             [(filt_str("Расхождение", True, "NotEqual", "", 4, False), "#FFF2CC", "Есть расхождение")], group_detail(3))

xml = "\n".join([
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<DataCompositionSchema xmlns="http://v8.1c.ru/8.1/data-composition-system/schema" xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    T("<dataSource>", 1), T("<name>ИсточникДанных1</name>", 2), T("<dataSourceType>Local</dataSourceType>", 2), T("</dataSource>", 1),
    dataset("Свод", SVOD), dataset("СводДок", SVODDOK), dataset("ПоДням", PODNYAM),
    all_totals(),
    param_period(),
    v1, v2, v3, v4, v5,
    "</DataCompositionSchema>", ""])
open(OUT, "wb").write(b"\xef\xbb\xbf" + xml.replace("\n", "\r\n").encode("utf-8"))
print("written", OUT, "lines", xml.count("\n"))
