# -*- coding: utf-8 -*-
"""Хирургические вставки П1 в СУЩЕСТВУЮЩИЕ объекты дампа WORKTREE (после sync_main_to_wt.py; Д1).
--etap meta: СС.ЭталоннаяСпецификация; РК: СпецификацияДома, ДатаНормы, ТЧ НормаДома и ЭкземплярыУзлов,
             колонки ТабличнаяЧастьОстатков.ПозицияСС/НеИзНормы, синоним Спецификация (§4.7).
--etap form: поле ЭталоннаяСпецификация (только просмотр) в Группа4 формы СС (Task 27).
Каждая вставка — по уникальному якорю с assert; повторный запуск ничего не дублирует. UUID детерминированные (uuid5).
Блоки XML — по образцу соседних элементов тех же файлов (E3 §1.3, §2; порядок свойств реквизита шапки/ТЧ — E1 §2.3–2.4)."""
import argparse
import os
import re
import sys
import uuid
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, read_text, write_text  # noqa: E402

NS = uuid.UUID("a6f1d1e0-5b8c-4f43-9d0e-2c7b6a91f3d4")


def uid(key):
    return str(uuid.uuid5(NS, key))


def syn(ru, uk, t):
    return (f"{t}<Synonym>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{ru}</v8:content>\n{t}\t</v8:item>\n"
            f"{t}\t<v8:item>\n{t}\t\t<v8:lang>uk</v8:lang>\n{t}\t\t<v8:content>{uk}</v8:content>\n{t}\t</v8:item>\n{t}</Synonym>\n")


def type_xml(spec, t):
    if spec.startswith("cfg:"):
        return f"{t}<v8:Type>{spec}</v8:Type>\n"
    if spec == "uuid":
        return f"{t}<v8:Type>v8:UUID</v8:Type>\n"
    if spec == "bool":
        return f"{t}<v8:Type>xs:boolean</v8:Type>\n"
    if spec == "datetime":
        return (f"{t}<v8:Type>xs:dateTime</v8:Type>\n{t}<v8:DateQualifiers>\n{t}\t<v8:DateFractions>DateTime</v8:DateFractions>\n"
                f"{t}</v8:DateQualifiers>\n")
    if spec.startswith("str:"):
        return (f"{t}<v8:Type>xs:string</v8:Type>\n{t}<v8:StringQualifiers>\n{t}\t<v8:Length>{spec[4:]}</v8:Length>\n"
                f"{t}\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{t}</v8:StringQualifiers>\n")
    if spec.startswith("num:"):
        d, f = spec[4:].split(",")
        return (f"{t}<v8:Type>xs:decimal</v8:Type>\n{t}<v8:NumberQualifiers>\n{t}\t<v8:Digits>{d}</v8:Digits>\n"
                f"{t}\t<v8:FractionDigits>{f}</v8:FractionDigits>\n{t}\t<v8:AllowedSign>Any</v8:AllowedSign>\n{t}</v8:NumberQualifiers>\n")
    raise ValueError(spec)


def fill_value(spec):
    if spec == "bool":
        return '<FillValue xsi:type="xs:boolean">false</FillValue>'
    if spec.startswith("str:"):
        return '<FillValue xsi:type="xs:string"/>'
    return '<FillValue xsi:nil="true"/>'


def attr(key, name, ru, uk, spec, t, kind):
    """kind: header_catalog | header_document | ts (реквизит ТЧ — без FillFromFillingValue/FillValue/Use)."""
    p = t + "\t\t"
    s = f'{t}<Attribute uuid="{uid(key)}">\n{t}\t<Properties>\n{p}<Name>{name}</Name>\n' + syn(ru, uk, p)
    s += f"{p}<Comment/>\n{p}<Type>\n" + type_xml(spec, p + "\t") + f"{p}</Type>\n"
    s += (f"{p}<PasswordMode>false</PasswordMode>\n{p}<Format/>\n{p}<EditFormat/>\n{p}<ToolTip/>\n"
          f"{p}<MarkNegatives>false</MarkNegatives>\n{p}<Mask/>\n{p}<MultiLine>false</MultiLine>\n"
          f"{p}<ExtendedEdit>false</ExtendedEdit>\n{p}<MinValue xsi:nil=\"true\"/>\n{p}<MaxValue xsi:nil=\"true\"/>\n")
    if kind != "ts":
        s += f"{p}<FillFromFillingValue>false</FillFromFillingValue>\n{p}{fill_value(spec)}\n"
    s += (f"{p}<FillChecking>DontCheck</FillChecking>\n{p}<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>\n"
          f"{p}<ChoiceParameterLinks/>\n{p}<ChoiceParameters/>\n{p}<QuickChoice>Auto</QuickChoice>\n"
          f"{p}<CreateOnInput>Auto</CreateOnInput>\n{p}<ChoiceForm/>\n{p}<LinkByType/>\n"
          f"{p}<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>\n")
    if kind == "header_catalog":
        s += f"{p}<Use>ForItem</Use>\n"
    s += f"{p}<Indexing>DontIndex</Indexing>\n{p}<FullTextSearch>Use</FullTextSearch>\n{p}<DataHistory>Use</DataHistory>\n"
    return s + f"{t}\t</Properties>\n{t}</Attribute>\n"


LINE_NUMBER = "".join(f"\t\t\t\t\t\t\t{x}\n" for x in (
    "<xr:LinkByType/>", "<xr:FillChecking>DontCheck</xr:FillChecking>", "<xr:MultiLine>false</xr:MultiLine>",
    "<xr:FillFromFillingValue>false</xr:FillFromFillingValue>", "<xr:CreateOnInput>Auto</xr:CreateOnInput>",
    '<xr:MaxValue xsi:nil="true"/>', "<xr:ToolTip/>", "<xr:ExtendedEdit>false</xr:ExtendedEdit>", "<xr:Format/>",
    "<xr:ChoiceForm/>", "<xr:QuickChoice>Auto</xr:QuickChoice>", "<xr:ChoiceHistoryOnInput>Auto</xr:ChoiceHistoryOnInput>",
    "<xr:EditFormat/>", "<xr:PasswordMode>false</xr:PasswordMode>", "<xr:DataHistory>Use</xr:DataHistory>",
    "<xr:MarkNegatives>false</xr:MarkNegatives>", '<xr:MinValue xsi:nil="true"/>', "<xr:Synonym/>", "<xr:Comment/>",
    "<xr:FullTextSearch>Use</xr:FullTextSearch>", "<xr:ChoiceParameterLinks/>", '<xr:FillValue xsi:nil="true"/>',
    "<xr:Mask/>", "<xr:ChoiceParameters/>"))


def doc_ts(doc, ts, ru, uk, cols):
    s = (f'\t\t\t<TabularSection uuid="{uid(doc + "." + ts)}">\n\t\t\t\t<InternalInfo>\n'
         f'\t\t\t\t\t<xr:GeneratedType name="DocumentTabularSection.{doc}.{ts}" category="TabularSection">\n'
         f'\t\t\t\t\t\t<xr:TypeId>{uid(doc + "." + ts + ".t1")}</xr:TypeId>\n\t\t\t\t\t\t<xr:ValueId>{uid(doc + "." + ts + ".v1")}</xr:ValueId>\n'
         f'\t\t\t\t\t</xr:GeneratedType>\n'
         f'\t\t\t\t\t<xr:GeneratedType name="DocumentTabularSectionRow.{doc}.{ts}" category="TabularSectionRow">\n'
         f'\t\t\t\t\t\t<xr:TypeId>{uid(doc + "." + ts + ".t2")}</xr:TypeId>\n\t\t\t\t\t\t<xr:ValueId>{uid(doc + "." + ts + ".v2")}</xr:ValueId>\n'
         f"\t\t\t\t\t</xr:GeneratedType>\n\t\t\t\t</InternalInfo>\n\t\t\t\t<Properties>\n\t\t\t\t\t<Name>{ts}</Name>\n"
         + syn(ru, uk, "\t\t\t\t\t") +
         "\t\t\t\t\t<Comment/>\n\t\t\t\t\t<ToolTip/>\n\t\t\t\t\t<FillChecking>DontCheck</FillChecking>\n"
         '\t\t\t\t\t<StandardAttributes>\n\t\t\t\t\t\t<xr:StandardAttribute name="LineNumber">\n' + LINE_NUMBER +
         "\t\t\t\t\t\t</xr:StandardAttribute>\n\t\t\t\t\t</StandardAttributes>\n\t\t\t\t</Properties>\n\t\t\t\t<ChildObjects>\n")
    for name, cru, cuk, spec in cols:
        s += attr(f"{doc}.{ts}.{name}", name, cru, cuk, spec, "\t\t\t\t\t", "ts")
    return s + "\t\t\t\t</ChildObjects>\n\t\t\t</TabularSection>\n"


REF_SPEC = "cfg:CatalogRef.А_СпецификацииДомов"
NORMA = [
    ("КлючЛиста", "Ключ листа", "Ключ листа", "str:32"), ("КлючСтроки", "Ключ строки", "Ключ рядка", "uuid"),
    ("КлючЭкземпляра", "Ключ экземпляра", "Ключ екземпляра", "str:32"),
    ("Спецификация", "Спецификация", "Специфікація", REF_SPEC),
    ("ВидСтроки", "Вид строки", "Вид рядка", "cfg:EnumRef.А_ВидыСтрокСпецификацииДомов"),
    ("Номенклатура", "Карточка", "Картка", "cfg:CatalogRef.Номенклатура"),
    ("ОбщееНазвание", "Общее название", "Загальна назва", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры"),
    ("ЭтапСС", "Этап СС", "Етап СС", "cfg:CatalogRef.ЭтапыРабот"), ("ПозицияСС", "Позиция СС", "Позиція СС", "str:200"),
    ("ЕдиницаСтроки", "Единица строки", "Одиниця рядка", "cfg:CatalogRef.КлассификаторЕдиницИзмерения"),
    ("КоличествоЕдСтроки", "Количество в ед. строки", "Кількість в од. рядка", "num:15,6"),
    ("Количество", "Количество", "Кількість", "num:15,6"), ("ЦенаНормы", "Цена нормы, с НДС", "Ціна норми, з ПДВ", "num:16,6"),
    ("СуммаНормы", "Сумма нормы, с НДС", "Сума норми, з ПДВ", "num:15,2"), ("ПутьУзла", "Путь узлов", "Шлях вузлів", "str:200"),
    ("Уровень", "Уровень", "Рівень", "num:2,0"),
]
EKZ = [
    ("КлючЭкземпляра", "Ключ экземпляра", "Ключ екземпляра", "str:32"),
    ("КлючЭкземпляраРодителя", "Ключ экземпляра родителя", "Ключ екземпляра батька", "str:32"),
    ("Узел", "Узел", "Вузол", REF_SPEC), ("НоменклатураВузла", "Карточка узла", "Картка вузла", "cfg:CatalogRef.Номенклатура"),
    ("ПокупаетсяГотовым", "Покупается готовым", "Купується готовим", "bool"), ("ПутьУзла", "Путь узлов", "Шлях вузлів", "str:200"),
    ("Уровень", "Уровень", "Рівень", "num:2,0"), ("Количество", "Узлов на дом", "Вузлів на будинок", "num:15,6"),
    ("СуммаНормы", "Сумма нормы узла, с НДС", "Сума норми вузла, з ПДВ", "num:15,2"),
]


def rep_once(text, old, new, what):
    assert text.count(old) == 1, f"якорь «{what}» найден {text.count(old)} раз"
    return text.replace(old, new, 1)


def patch_ss():
    rel = "Catalogs/СтруктураСебестоимости.xml"
    path = os.path.join(BB_WT, rel)
    text, _, _ = read_text(path)
    if "<Name>ЭталоннаяСпецификация</Name>" in text:
        print("  СС: уже есть")
        return
    block = attr("СС.ЭталоннаяСпецификация", "ЭталоннаяСпецификация", "Эталонная спецификация", "Еталонна специфікація",
                 REF_SPEC, "\t\t\t", "header_catalog")
    first_ts = "\t\t\t<TabularSection uuid="
    i = text.index(first_ts)
    assert text[:i].rstrip("\n").endswith("</Attribute>"), "перед первой ТЧ СС не реквизит"
    assert "<Name>ЗапретитьРедактированиеСС</Name>" in text[:i]
    text = text[:i] + block + text[i:]
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
    print("  СС: + ЭталоннаяСпецификация")


def patch_rk():
    rel = "Documents/РасчетКомплектаций.xml"
    path = os.path.join(BB_WT, rel)
    text, _, _ = read_text(path)
    assert "<Name>ПричиниЦеху</Name>" in text, "РК не из основного каталога (нет ТЧ ПричиниЦеху) — сначала sync_main_to_wt.py"
    doc = "РасчетКомплектаций"
    if "<Name>СпецификацияДома</Name>" not in text:
        hdr = (attr(doc + ".СпецификацияДома", "СпецификацияДома", "Спецификация дома", "Специфікація будинку", REF_SPEC,
                    "\t\t\t", "header_document")
               + attr(doc + ".ДатаНормы", "ДатаНормы", "Норма от", "Норма від", "datetime", "\t\t\t", "header_document"))
        anchor = "\t\t\t<Form>ФормаДокумента</Form>\n"
        text = rep_once(text, anchor, hdr + anchor, "Form ФормаДокумента")
    if "<Name>НормаДома</Name>" not in text:
        ts = (doc_ts(doc, "НормаДома", "Норма дома (снимок)", "Норма будинку (знімок)", NORMA)
              + doc_ts(doc, "ЭкземплярыУзлов", "Узлы дома (снимок)", "Вузли будинку (знімок)", EKZ))
        anchor = "\t\t\t<Template>МакетПланФакт</Template>\n"
        text = rep_once(text, anchor, ts + anchor, "Template МакетПланФакт")
    start = text.index("<Name>ТабличнаяЧастьОстатков</Name>")
    end_ts = text.index("\t\t\t\t</ChildObjects>\n\t\t\t</TabularSection>\n", start)
    if "<Name>НеИзНормы</Name>" not in text[start:end_ts]:
        cols = (attr(doc + ".ТабличнаяЧастьОстатков.ПозицияСС", "ПозицияСС", "Позиция СС", "Позиція СС", "str:200", "\t\t\t\t\t", "ts")
                + attr(doc + ".ТабличнаяЧастьОстатков.НеИзНормы", "НеИзНормы", "Не из нормы", "Не з норми", "bool", "\t\t\t\t\t", "ts"))
        text = text[:end_ts] + cols + text[end_ts:]
    old_syn = ("<Name>Спецификация</Name>\n" + syn("Спецификация", "Специфікація (норми)", "\t\t\t\t\t"))
    if old_syn in text:
        text = rep_once(text, old_syn, "<Name>Спецификация</Name>\n"
                        + syn("Структура себестоимости (СС)", "Структура собівартості (СС)", "\t\t\t\t\t"), "синоним Спецификация")
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
    print("  РК: реквизиты, ТЧ, колонки, синоним — на месте")


def patch_ss_form():
    rel = "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml"
    path = os.path.join(BB_WT, rel)
    text, _, _ = read_text(path)
    if 'name="ЭталоннаяСпецификация"' in text:
        print("  форма СС: уже есть")
        return
    start = rep_start = text.index('<InputField name="КоличествоДомов"')
    end = text.index("\t\t\t\t</InputField>\n", rep_start) + len("\t\t\t\t</InputField>\n")
    assert '<UsualGroup name="Группа4"' in text[:start], "поле КоличествоДомов не в Группа4"
    n = max(int(x) for x in re.findall(r' id="(\d+)"', text.split("<Attributes>")[0])) + 1
    block = (f'\t\t\t\t<InputField name="ЭталоннаяСпецификация" id="{n}">\n'
             "\t\t\t\t\t<DataPath>Объект.ЭталоннаяСпецификация</DataPath>\n\t\t\t\t\t<ReadOnly>true</ReadOnly>\n"
             f'\t\t\t\t\t<ContextMenu name="ЭталоннаяСпецификацияКонтекстноеМеню" id="{n + 1}"/>\n'
             f'\t\t\t\t\t<ExtendedTooltip name="ЭталоннаяСпецификацияРасширеннаяПодсказка" id="{n + 2}"/>\n'
             "\t\t\t\t</InputField>\n")
    text = text[:end] + block + text[end:]
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)
    print(f"  форма СС: + ЭталоннаяСпецификация (id {n}–{n + 2}, только просмотр)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--etap", choices=["meta", "form"], required=True)
    a = ap.parse_args()
    if a.etap == "meta":
        patch_ss()
        patch_rk()
    else:
        patch_ss_form()
