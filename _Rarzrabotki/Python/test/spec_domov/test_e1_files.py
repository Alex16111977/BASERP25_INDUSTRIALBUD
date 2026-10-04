# -*- coding: utf-8 -*-
"""Статическая приёмка файлов этапа 1A в дампе WORKTREE (до загрузки). Разделы: enums, catalog, existing, register.
Запуск: python test_e1_files.py [--only enums|catalog|existing|register|module|ssform|forms]"""
import argparse
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT  # noqa: E402
from test_e1_metadata import HEADER, RK_TS, ROLE_RIGHTS, TS, UK_ENUM  # noqa: E402

RE_ATTR_NAME = re.compile(r"<Attribute uuid=\"[^\"]+\">\s*<Properties>\s*<Name>([^<]+)</Name>")


def load(rel):
    p = os.path.join(BB_WT, rel)
    assert os.path.exists(p), f"нет файла {rel}"
    raw = open(p, "rb").read()
    assert raw.startswith(b"\xef\xbb\xbf"), f"{rel}: нет BOM"
    assert raw.replace(b"\r\n", b"").count(b"\n") == 0, f"{rel}: не CRLF"
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")
    xml.dom.minidom.parseString(text.encode("utf-8"))
    return text


def clean_new(rel, text):
    assert " />" not in text and "d5p1" not in text, f"{rel}: шум генератора"
    ru, uk = text.count("<v8:lang>ru</v8:lang>"), text.count("<v8:lang>uk</v8:lang>")
    assert ru == uk, f"{rel}: ru {ru} ≠ uk {uk}"


def ts_block(text, ts):
    for m in re.finditer(r"\t\t\t<TabularSection uuid=.*?\n\t\t\t</TabularSection>\n", text, re.S):
        if f"<Name>{ts}</Name>" in m.group(0).split("<ChildObjects>")[0]:
            return m.group(0)
    raise AssertionError(f"нет ТЧ {ts}")


def attr_block(text, name):
    for m in re.finditer(r"<Attribute uuid=\"[^\"]+\">.*?</Attribute>", text, re.S):
        if f"<Name>{name}</Name>" in m.group(0):
            return m.group(0)
    raise AssertionError(f"нет реквизита {name}")


def check_enums():
    for name, values in UK_ENUM.items():
        rel = f"Enums/{name}.xml"
        t = load(rel)
        clean_new(rel, t)
        got = re.findall(r"<EnumValue uuid=\"[^\"]+\">\s*<Properties>\s*<Name>([^<]+)</Name>", t)
        assert got == list(values), (name, got)
        for uk in values.values():
            assert f"<v8:content>{uk}</v8:content>" in t, (name, uk)
    print("enums OK")


def check_catalog():
    rel = "Catalogs/А_СпецификацииДомов.xml"
    t = load(rel)
    clean_new(rel, t)
    for tag, val in (("Hierarchical", "true"), ("HierarchyType", "HierarchyFoldersAndItems"), ("CodeLength", "9"),
                     ("DescriptionLength", "150"), ("CheckUnique", "true"), ("Autonumbering", "true"),
                     ("DataLockControlMode", "Managed")):
        assert f"<{tag}>{val}</{tag}>" in t, (tag, val)
    assert "<Owners/>" in t
    head = t.split("<TabularSection", 1)[0]
    assert RE_ATTR_NAME.findall(head) == HEADER, RE_ATTR_NAME.findall(head)
    for ts, cols in TS.items():
        assert RE_ATTR_NAME.findall(ts_block(t, ts)) == cols, ts
    assert t.count("Enum.А_ВидыСпецификацийДомов.EnumValue.Узел") == 3
    assert t.count("<v8:Type>v8:UUID</v8:Type>") == 3
    assert "<v8:content>Специфікації будинків і вузлів</v8:content>" in t
    assert "<v8:Length>200</v8:Length>" in attr_block(ts_block(t, "Состав"), "ПозицияСС")
    q = attr_block(ts_block(t, "Состав"), "Количество")
    assert "<v8:Digits>15</v8:Digits>" in q and "<v8:FractionDigits>6</v8:FractionDigits>" in q
    k = attr_block(ts_block(t, "Состав"), "Коэффициент")
    assert "<v8:FractionDigits>9</v8:FractionDigits>" in k
    c = attr_block(ts_block(t, "Состав"), "ЦенаНормы")
    assert "<v8:Digits>16</v8:Digits>" in c and "<v8:FractionDigits>6</v8:FractionDigits>" in c
    print("catalog OK")


def check_existing():
    ss = load("Catalogs/СтруктураСебестоимости.xml")
    a = attr_block(ss, "ЭталоннаяСпецификация")
    assert "<v8:Type>cfg:CatalogRef.А_СпецификацииДомов</v8:Type>" in a and "Еталонна специфікація" in a
    assert ss.index("<Name>ЭталоннаяСпецификация</Name>") < ss.index("\t\t\t<TabularSection")
    assert ss.count("<Name>ЭталоннаяСпецификация</Name>") == 1
    rk = load("Documents/РасчетКомплектаций.xml")
    for n in ("СпецификацияДома", "ДатаНормы"):
        assert rk.count(f"<Name>{n}</Name>") == 1 and rk.index(f"<Name>{n}</Name>") < rk.index("<Form>ФормаДокумента</Form>")
    for ts, cols in RK_TS.items():
        assert RE_ATTR_NAME.findall(ts_block(rk, ts)) == cols, ts
    assert rk.index("<Name>НормаДома</Name>") < rk.index("<Template>МакетПланФакт</Template>")
    assert RE_ATTR_NAME.findall(ts_block(rk, "ТабличнаяЧастьОстатков"))[-2:] == ["ПозицияСС", "НеИзНормы"]
    sp = attr_block(rk, "Спецификация")
    assert "Структура собівартості (СС)" in sp and "Структура себестоимости (СС)" in sp and "(норми)" not in sp
    assert "<Posting>Deny</Posting>" in rk
    assert rk.count("\t\t\t<TabularSection uuid=") == 15 and rk.count("<Name>ПричиниЦеху</Name>") == 1
    print("existing OK")


def check_register():
    cfg = load("Configuration.xml")
    for tag, name in [("Catalog", "А_СпецификацииДомов")] + [("Enum", e) for e in UK_ENUM]:
        assert cfg.count(f"\t\t\t<{tag}>{name}</{tag}>\n") == 1, (tag, name)
    for role, rights in ROLE_RIGHTS.items():
        t = load(f"Roles/{role}/Ext/Rights.xml")
        m = re.search(r"<object>\s*<name>Catalog\.А_СпецификацииДомов</name>(.*?)</object>", t, re.S)
        assert m, role
        assert re.findall(r"<name>(\w+)</name>\s*<value>true</value>", m.group(1)) == rights, role
    sub = load("Subsystems/Склад/Subsystems/А_СписанняБудинків.xml")
    assert sub.count(">Catalog.А_СпецификацииДомов</xr:Item>") == 1
    ci = load("Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml")
    assert ci.count('"Catalog.А_СпецификацииДомов.StandardCommand.OpenList"') == 1
    print("register OK")


def check_module():
    t = load("CommonModules/А_СпецификацииДомовСервер.xml")
    clean_new("CommonModules/А_СпецификацииДомовСервер.xml", t)
    for tag, val in (("Server", "true"), ("ExternalConnection", "true"), ("ServerCall", "false"), ("Privileged", "false"),
                     ("ClientManagedApplication", "false"), ("Global", "false")):
        assert f"<{tag}>{val}</{tag}>" in t, (tag, val)
    m = open(os.path.join(BB_WT, "CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"), encoding="utf-8-sig").read()
    for fn in ("Функция ХешMD5(", "Функция РазвернутьДоКарточек(", "Функция НайтиЦикл(", "Процедура ПередЗаписьюСпецификации(",
               "Функция СделатьЭталономСС(", "Функция Входимость(", "Функция ПроверитьЦиклы("):
        assert fn in m, fn
    o = open(os.path.join(BB_WT, "Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl"), encoding="utf-8-sig").read()
    assert "А_СпецификацииДомовСервер.ПередЗаписьюСпецификации(ЭтотОбъект, Отказ)" in o
    print("module OK")


def check_forms():
    cat = load("Catalogs/А_СпецификацииДомов.xml")
    for tag, form in (("DefaultObjectForm", "ФормаЭлемента"), ("DefaultListForm", "ФормаСписка"),
                      ("DefaultFolderForm", "ФормаГруппы")):
        assert f"<{tag}>Catalog.А_СпецификацииДомов.Form.{form}</{tag}>" in cat, tag
        assert cat.count(f"<Form>{form}</Form>") == 1, form
        md = load(f"Catalogs/А_СпецификацииДомов/Forms/{form}.xml")
        clean_new(form, md)
        fx = load(f"Catalogs/А_СпецификацииДомов/Forms/{form}/Ext/Form.xml")
        clean_new(form + "/Form.xml", fx)
        assert 'version="2.13"' in fx and "<Action xsi:type" not in fx
    el = load("Catalogs/А_СпецификацииДомов/Forms/ФормаЭлемента/Ext/Form.xml")
    for name in ("Состав", "Дома", "ДеревоУзлов", "РазвернутыйСостав", "СравнениеСС", "ЗакупочныйСостав", "ДомаФакта", "ФактПоДомам"):
        assert f'<Table name="{name}"' in el, name
    for cmd in ("СделатьЭталономСС", "ПоказатьВходимость", "ОбновитьДерево", "РазвернутьДоДеталей"):
        assert f'<Command name="{cmd}"' in el and f"<Action>{cmd}</Action>" in el, cmd
    sp = load("Catalogs/А_СпецификацииДомов/Forms/ФормаСписка/Ext/Form.xml")
    assert "<ManualQuery>true</ManualQuery>" in sp and "Список.ЭталонДляСС" in sp and '<Command name="ПроверитьЦиклы"' in sp
    for mod in ("ФормаЭлемента", "ФормаСписка"):
        m = open(os.path.join(BB_WT, f"Catalogs/А_СпецификацииДомов/Forms/{mod}/Ext/Form/Module.bsl"), "rb").read()
        assert m.startswith(b"\xef\xbb\xbf") and b"\r\n" in m, mod
    print("forms OK")


def check_ssform():
    f = load("Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml")
    m = re.search(r'<InputField name="ЭталоннаяСпецификация" id="(\d+)">(.*?)</InputField>', f, re.S)
    assert m and "<DataPath>Объект.ЭталоннаяСпецификация</DataPath>" in m.group(2) and "<ReadOnly>true</ReadOnly>" in m.group(2)
    g4 = f[f.index('<UsualGroup name="Группа4"'):]
    assert g4.index('name="КоличествоДомов"') < g4.index('name="ЭталоннаяСпецификация"') < g4.index("</UsualGroup>")
    ids = re.findall(r' id="(\d+)"', f.split("<Attributes>")[0])
    assert len(ids) == len(set(ids)), "повтор id в форме СС"
    cfg = load("Configuration.xml")
    assert cfg.count("\t\t\t<CommonModule>А_СпецификацииДомовСервер</CommonModule>\n") == 1
    print("ssform OK")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--only")
    a = ap.parse_args()
    parts = {"enums": check_enums, "catalog": check_catalog, "existing": check_existing, "register": check_register,
             "module": check_module, "forms": check_forms, "ssform": check_ssform}
    for key, fn in parts.items():
        if a.only in (None, key):
            fn()
    print("OK test_e1_files", a.only or "all")
