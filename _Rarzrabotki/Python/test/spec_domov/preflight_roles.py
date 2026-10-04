# -*- coding: utf-8 -*-
"""Гейт 3: права роли в Rights.xml ОСНОВНОГО каталога = ПравоДоступа в базе (ловит чужие незакоммиченные правки:
наша загрузка роли иначе протащит их в базу или отнимет права). Новые объекты П1 до загрузки пропускаются."""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, connect  # noqa: E402

ROLES = ["ДобавлениеИзменениеДанныхБухгалтерии", "ЧтениеДанныхБухгалтерии", "ЗапускТонкогоКлиента"]
NEW_OBJECTS = {"Catalog.А_СпецификацииДомов"}
COLLECTION = {"Catalog": "Catalogs", "Document": "Documents", "InformationRegister": "InformationRegisters",
              "AccumulationRegister": "AccumulationRegisters", "AccountingRegister": "AccountingRegisters",
              "DataProcessor": "DataProcessors", "Report": "Reports", "Constant": "Constants",
              "DocumentJournal": "DocumentJournals", "ChartOfAccounts": "ChartsOfAccounts",
              "ChartOfCharacteristicTypes": "ChartsOfCharacteristicTypes", "ExchangePlan": "ExchangePlans",
              "BusinessProcess": "BusinessProcesses", "Task": "Tasks", "CalculationRegister": "CalculationRegisters",
              "ChartOfCalculationTypes": "ChartsOfCalculationTypes"}
CHECKED = {"Read", "Insert", "Update", "Delete", "View", "Edit", "InteractiveInsert", "InteractiveDelete",
           "InteractiveSetDeletionMark", "InteractiveClearDeletionMark", "InteractiveDeleteMarked", "InputByString",
           "Posting", "UndoPosting", "InteractivePosting", "InteractiveUndoPosting", "Use", "TotalsControl"}
RE_OBJ = re.compile(r"<object>\s*<name>([^<]+)</name>(.*?)</object>", re.S)
RE_RIGHT = re.compile(r"<right>\s*<name>(\w+)</name>\s*<value>(\w+)</value>")


def file_rights(role):
    text = io.open(os.path.join(BB_MAIN, "Roles", role, "Ext", "Rights.xml"), encoding="utf-8-sig").read()
    out = {}
    for name, body in RE_OBJ.findall(text):
        if name.count(".") == 1 and name.split(".")[0] in COLLECTION:
            out[name] = {r: v == "true" for r, v in RE_RIGHT.findall(body) if r in CHECKED}
    return out


def main():
    B = connect()
    md = B.Metadata
    fails = []
    for role in ROLES:
        md_role = md.Roles.Find(role)
        checked = skipped = 0
        for name, flags in file_rights(role).items():
            kind, obj = name.split(".")
            md_obj = getattr(md, COLLECTION[kind]).Find(obj)
            if md_obj is None:
                if name in NEW_OBJECTS:
                    skipped += 1
                    continue
                fails.append(f"{role}: {name} есть в Rights.xml, объекта нет в базе")
                continue
            for right, want in flags.items():
                try:
                    have = bool(B.ПравоДоступа(right, md_obj, md_role))
                except Exception:
                    continue
                checked += 1
                if have != want:
                    fails.append(f"{role}: {name}.{right} — файл {want}, база {have}")
        print(f"{role}: сверено прав {checked}, новых П1 пропущено {skipped}")
    print(f"РАСХОЖДЕНИЙ: {len(fails)}")
    for f in fails[:40]:
        print("   ", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
