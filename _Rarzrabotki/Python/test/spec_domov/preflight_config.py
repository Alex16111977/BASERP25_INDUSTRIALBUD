# -*- coding: utf-8 -*-
"""Гейт 2: Partial-загрузка с Configuration.xml, где нет объекта базы, УДАЛЯЕТ этот объект (память
config_xml_partial_load_deletes_objects). Правило: база ⊆ файл и файл − база ⊆ EXPECTED_NEW. Все коллекции Metadata.
Запуск: python preflight_config.py [путь к Configuration.xml]  (по умолчанию — основной каталог). Код 0 — грузить можно."""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, connect  # noqa: E402

EXPECTED_NEW = {
    ("Catalog", "А_СпецификацииДомов"),
    ("Enum", "А_ВидыСпецификацийДомов"),
    ("Enum", "А_СтатусыСпецификацийДомов"),
    ("Enum", "А_ВидыСтрокСпецификацииДомов"),
    ("Enum", "А_ИсточникиКарточкиСпецификации"),
    ("Enum", "А_ГруппыЗакупочнойСпецификации"),
    ("Enum", "А_ИсточникиЦеныЗакупки"),
    ("CommonModule", "А_СпецификацииДомовСервер"),
}
EXPECTED_REMOVED = set()

COLL = {"Catalog": "Catalogs", "Document": "Documents", "Enum": "Enums", "InformationRegister": "InformationRegisters",
        "AccumulationRegister": "AccumulationRegisters", "AccountingRegister": "AccountingRegisters",
        "DataProcessor": "DataProcessors", "Report": "Reports", "CommonModule": "CommonModules", "Role": "Roles",
        "EventSubscription": "EventSubscriptions", "CommandGroup": "CommandGroups", "Subsystem": "Subsystems",
        "Constant": "Constants", "CommonForm": "CommonForms", "DocumentJournal": "DocumentJournals",
        "ScheduledJob": "ScheduledJobs", "ExchangePlan": "ExchangePlans", "ChartOfAccounts": "ChartsOfAccounts",
        "ChartOfCharacteristicTypes": "ChartsOfCharacteristicTypes", "BusinessProcess": "BusinessProcesses",
        "Task": "Tasks", "CommonTemplate": "CommonTemplates", "CommonPicture": "CommonPictures",
        "CommonAttribute": "CommonAttributes", "SessionParameter": "SessionParameters",
        "FunctionalOption": "FunctionalOptions", "DefinedType": "DefinedTypes", "FilterCriterion": "FilterCriteria",
        "SettingsStorage": "SettingsStorages", "CommonCommand": "CommonCommands", "StyleItem": "StyleItems",
        "Language": "Languages", "WebService": "WebServices", "HTTPService": "HTTPServices", "XDTOPackage": "XDTOPackages",
        "FunctionalOptionsParameter": "FunctionalOptionsParameters", "CalculationRegister": "CalculationRegisters",
        "ChartOfCalculationTypes": "ChartsOfCalculationTypes", "Sequence": "Sequences", "Style": "Styles",
        "WSReference": "WSReferences", "ExternalDataSource": "ExternalDataSources",
        "DocumentNumerator": "DocumentNumerators"}


def file_objects(path):
    text = io.open(path, encoding="utf-8-sig").read()
    body = text[text.find("<ChildObjects>"):text.find("</ChildObjects>")]
    return {(k, n) for k, n in re.findall(r"<(\w+)>([^<]+)</\1>", body) if k in COLL}


def db_objects(B):
    md, found = B.Metadata, set()
    for kind, prop in COLL.items():
        try:
            coll = getattr(md, prop)
        except Exception:
            continue
        for i in range(coll.Count()):
            found.add((kind, coll.Get(i).Name))
    return found


def main(path):
    B = connect()
    db = db_objects(B)
    kinds = {k for k, _ in db}
    cfg = {x for x in file_objects(path) if x[0] in kinds or x in EXPECTED_NEW}
    lost = sorted(db - cfg - EXPECTED_REMOVED)
    extra = sorted(cfg - db)
    unexpected = sorted(set(extra) - EXPECTED_NEW)
    print(f"Configuration.xml: {path}")
    print(f"объектов в базе {len(db)}, в файле {len(cfg)}")
    print(f"новых П1 в файле, но не в базе: {[n for _, n in extra if (_, n) in EXPECTED_NEW]}")
    print(f"новых П1 уже в базе: {[n for k, n in sorted(EXPECTED_NEW & db)]}")
    fails = [f"УДАЛИТ из базы: {k}.{n}" for k, n in lost] + [f"посторонний в файле: {k}.{n}" for k, n in unexpected]
    print(f"ПРОВАЛОВ: {len(fails)}")
    for f in fails[:50]:
        print("   ", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(BB_MAIN, "Configuration.xml")))
