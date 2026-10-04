# -*- coding: utf-8 -*-
"""Регистрация объектов П1 в дампе WORKTREE хирургически (role-compile и meta-compile переписывают файлы целиком — не
используются, память role_compile_configuration_xml_gotcha). Повторный запуск ничего не дублирует.
--etap 1a: Configuration.xml (справочник + 6 перечислений), роли (Д2), подсистема А_СписанняБудинків.
--etap 1b: Configuration.xml (общий модуль А_СпецификацииДомовСервер)."""
import argparse
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, read_text, write_text  # noqa: E402
from test_e1_metadata import ROLE_RIGHTS, UK_ENUM  # noqa: E402

NEW_1A = [("Catalog", "А_СпецификацииДомов")] + [("Enum", e) for e in UK_ENUM]
NEW_1B = [("CommonModule", "А_СпецификацииДомовСервер")]
SUB = "Subsystems/Склад/Subsystems/А_СписанняБудинків.xml"
SUB_CI = "Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml"
OBJ = "Catalog.А_СпецификацииДомов"


def save(path, text):
    xml.dom.minidom.parseString(text.encode("utf-8"))
    write_text(path, text)


def register_config(objects):
    path = os.path.join(BB_WT, "Configuration.xml")
    text, _, _ = read_text(path)
    added = 0
    for tag, name in objects:
        line = f"\t\t\t<{tag}>{name}</{tag}>\n"
        if line in text:
            continue
        last = list(re.finditer(rf"^\t\t\t<{tag}>[^<]+</{tag}>\n", text, re.M))[-1]      # конец своей группы (E1 §1.1)
        text = text[:last.end()] + line + text[last.end():]
        added += 1
    save(path, text)
    print(f"  Configuration.xml: +{added}")


def rights_block(name, rights):
    lines = ["\t<object>", f"\t\t<name>{name}</name>"]
    for r in rights:
        lines += ["\t\t<right>", f"\t\t\t<name>{r}</name>", "\t\t\t<value>true</value>", "\t\t</right>"]
    return "\n".join(lines + ["\t</object>"]) + "\n"


def register_roles():
    for role, rights in ROLE_RIGHTS.items():
        path = os.path.join(BB_WT, "Roles", role, "Ext", "Rights.xml")
        text, _, _ = read_text(path)
        if f"<name>{OBJ}</name>" in text:
            print(f"  роль {role}: уже есть")
            continue
        anchor = "\t<restrictionTemplate>" if "\t<restrictionTemplate>" in text else "</Rights>"
        i = text.index(anchor)
        text = text[:i] + rights_block(OBJ, rights) + text[i:]
        save(path, text)
        print(f"  роль {role}: + {OBJ} ({', '.join(rights)})")


def register_subsystem():
    path = os.path.join(BB_WT, SUB)
    text, _, _ = read_text(path)
    item = f'\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">{OBJ}</xr:Item>\n'
    anchor = '\t\t\t\t<xr:Item xsi:type="xr:MDObjectRef">Catalog.СтруктураСебестоимости</xr:Item>\n'
    if item not in text:
        assert text.count(anchor) == 1, "якорь СС в Content"
        save(path, text.replace(anchor, anchor + item, 1))
    path = os.path.join(BB_WT, SUB_CI)
    text, _, _ = read_text(path)
    cmd = (f'\t\t<Command name="{OBJ}.StandardCommand.OpenList">\n\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n'
           "\t\t</Command>\n")
    anchor = ('\t\t<Command name="Catalog.СтруктураСебестоимости.StandardCommand.OpenList">\n'
              "\t\t\t<CommandGroup>NavigationPanelOrdinary</CommandGroup>\n\t\t</Command>\n")
    if cmd not in text:
        assert text.count(anchor) == 1, "якорь СС в CommandInterface"
        save(path, text.replace(anchor, anchor + cmd, 1))
    print("  подсистема А_СписанняБудинків: справочник в составе и в панели навигации")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--etap", choices=["1a", "1b"], required=True)
    a = ap.parse_args()
    if a.etap == "1a":
        register_config(NEW_1A)
        register_roles()
        register_subsystem()
    else:
        register_config(NEW_1B)
