# -*- coding: utf-8 -*-
"""
Атомарные правки СКД А_ОтчетPL (оба файла: конфиг + внешний дамп):
  A) +2 поля набора данных (СуммаЕРПФ1, СуммаЕРПФ2) после поля СуммаЕРП
  B) +2 поля (РазницаФ1, РазницаФ2) после поля Разница
  C) +4 totalField (ресурсы) после totalField Разница
  D) в выборку вариантов: после СуммаЕРП добавить СуммаЕРПФ1/Ф2 (клон блока — точная отбивка)
  E) в выборку вариантов: после Разница добавить РазницаФ1/Ф2
Каждая структурная правка (A/B/C) — count==1 на файл; D/E — клонируют все selection-блоки.
Бэкап рядом (.bak_skd_formapl).
"""
import sys, io, os, re

if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

FILES = [
    r"C:\Configuration_downloads\BASERP25\Reports\А_ОтчетPL\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml",
    r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ОтчетPL\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml",
]


def field_xml(name, title):
    return (
        f'\t\t<field xsi:type="DataSetFieldField">\n'
        f'\t\t\t<dataPath>{name}</dataPath>\n'
        f'\t\t\t<field>{name}</field>\n'
        f'\t\t\t<title xsi:type="v8:LocalStringType">\n'
        f'\t\t\t\t<v8:item>\n'
        f'\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
        f'\t\t\t\t\t<v8:content>{title}</v8:content>\n'
        f'\t\t\t\t</v8:item>\n'
        f'\t\t\t</title>\n'
        f'\t\t\t<valueType>\n'
        f'\t\t\t\t<v8:Type>xs:decimal</v8:Type>\n'
        f'\t\t\t\t<v8:NumberQualifiers>\n'
        f'\t\t\t\t\t<v8:Digits>15</v8:Digits>\n'
        f'\t\t\t\t\t<v8:FractionDigits>2</v8:FractionDigits>\n'
        f'\t\t\t\t\t<v8:AllowedSign>Any</v8:AllowedSign>\n'
        f'\t\t\t\t</v8:NumberQualifiers>\n'
        f'\t\t\t</valueType>\n'
        f'\t\t</field>\n'
    )


def total_xml(name, expr):
    return (
        f'\t<totalField>\n'
        f'\t\t<dataPath>{name}</dataPath>\n'
        f'\t\t<expression>{expr}</expression>\n'
        f'\t</totalField>\n'
    )


def must_replace(text, anchor, repl, tag):
    cnt = text.count(anchor)
    if cnt != 1:
        raise RuntimeError(f"{tag}: anchor встречается {cnt} раз (ожидалось 1)")
    return text.replace(anchor, repl)


def clone_selection(text, fld, new_flds):
    pat = re.compile(
        r'([ \t]*)<dcsset:item xsi:type="dcsset:SelectedItemField">\n'
        r'[ \t]*<dcsset:field>' + re.escape(fld) + r'</dcsset:field>\n'
        r'[ \t]*</dcsset:item>')

    def repl(m):
        block = m.group(0)
        out = block
        for nf in new_flds:
            out += '\n' + block.replace('<dcsset:field>' + fld + '</dcsset:field>',
                                        '<dcsset:field>' + nf + '</dcsset:field>')
        return out
    return pat.subn(repl, text)


def process(path):
    with open(path, "r", encoding="utf-8") as f:
        src = f.read()
    orig = src
    enc_bom = src.startswith("﻿")

    # A) поля СуммаЕРПФ1/Ф2 перед полем СуммаДДСизКазныРасход
    ancA = '\t\t<field xsi:type="DataSetFieldField">\n\t\t\t<dataPath>СуммаДДСизКазныРасход</dataPath>'
    src = must_replace(src, ancA,
                       field_xml("СуммаЕРПФ1", "Сумма ЕРП Ф1 (факт)") +
                       field_xml("СуммаЕРПФ2", "Сумма ЕРП Ф2 (факт)") + ancA, "A")

    # B) поля РазницаФ1/Ф2 перед полем НаправлениеДеятельности
    ancB = '\t\t<field xsi:type="DataSetFieldField">\n\t\t\t<dataPath>НаправлениеДеятельности</dataPath>'
    src = must_replace(src, ancB,
                       field_xml("РазницаФ1", "Разница Ф1 (PL-ЕРП)") +
                       field_xml("РазницаФ2", "Разница Ф2 (PL-ЕРП)") + ancB, "B")

    # C) +4 totalField после totalField Разница
    ancC = ('\t<totalField>\n\t\t<dataPath>Разница</dataPath>\n'
            '\t\t<expression>Сумма(СуммаPL) - Сумма(СуммаЕРП)</expression>\n\t</totalField>\n')
    addC = (total_xml("СуммаЕРПФ1", "Сумма(СуммаЕРПФ1)") +
            total_xml("СуммаЕРПФ2", "Сумма(СуммаЕРПФ2)") +
            total_xml("РазницаФ1", "Сумма(СуммаФ1) - Сумма(СуммаЕРПФ1)") +
            total_xml("РазницаФ2", "Сумма(СуммаФ2) - Сумма(СуммаЕРПФ2)"))
    src = must_replace(src, ancC, ancC + addC, "C")

    # D/E) показ в вариантах
    src, nD = clone_selection(src, "СуммаЕРП", ["СуммаЕРПФ1", "СуммаЕРПФ2"])
    src, nE = clone_selection(src, "Разница", ["РазницаФ1", "РазницаФ2"])

    # проверки
    for need in ("<dataPath>СуммаЕРПФ1</dataPath>", "<dataPath>СуммаЕРПФ2</dataPath>",
                 "<dataPath>РазницаФ1</dataPath>", "<dataPath>РазницаФ2</dataPath>",
                 "<expression>Сумма(СуммаФ1) - Сумма(СуммаЕРПФ1)</expression>",
                 "<expression>Сумма(СуммаФ2) - Сумма(СуммаЕРПФ2)</expression>"):
        assert need in src, f"после правок нет {need}"

    with open(path + ".bak_skd_formapl", "w", encoding="utf-8") as f:
        f.write(orig)
    with open(path, "w", encoding="utf-8") as f:
        f.write(src)
    print(f"OK {os.path.basename(os.path.dirname(os.path.dirname(path)))} <- {path}")
    print(f"   selection-блоков клонировано: СуммаЕРП={nD}, Разница={nE}; BOM={enc_bom}; "
          f"размер {len(orig)} -> {len(src)}")


for p in FILES:
    process(p)
print("\nГотово. Бэкапы: *.bak_skd_formapl")
