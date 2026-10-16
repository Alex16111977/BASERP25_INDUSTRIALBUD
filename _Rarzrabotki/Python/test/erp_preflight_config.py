# -*- coding: utf-8 -*-
"""Гейт перед частичной загрузкой конфигурации в рабочую базу ERP (BaseERP).

Частичная загрузка с Configuration.xml, в котором НЕТ объекта, существующего в базе, УДАЛЯЕТ этот объект
(память config_xml_partial_load_deletes_objects). Копия mk_preflight_config.py (BuhBud) для BaseERP.

Правило: объекты_базы ⊆ объекты_Configuration.xml, а лишние в файле - только EXPECTED_NEW.

    python kazna_preflight_config.py                  # основной каталог (корень репозитория)
    python kazna_preflight_config.py <путь к Configuration.xml>

Код возврата 0 - загружать можно, 1 - НЕЛЬЗЯ.
"""
import io
import re
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

import win32com.client

MAIN_CONFIG = r"C:\Configuration_downloads\BASERP25\Configuration.xml"
CONNECTION = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

# Новые объекты текущей доработки: есть в файле, в базе появятся после загрузки.
EXPECTED_NEW = {
    ("Enum", "А_СпособыРасчетаПроцентов"),  # 16.10.2026 нарахування відсотків
    ("InformationRegister", "А_ПроцентныеСтавкиКредитов"),  # 16.10.2026
    ("Document", "А_НачислениеПроцентовПоКредитам"),  # 16.10.2026
    ("CommonModule", "А_НачислениеПроцентовПоКредитамСервер"),  # 16.10.2026
    ("InformationRegister", "А_ТоварыСчетовПоставщиков"),  # 15.10.2026 товары счетов Казны
    ("Enum", "А_ВидыКредитныхДоговоров"),  # 14.10.2026 кредитные договоры ОУМП
    ("Constant", "А_НеПредлагатьПерезапускПриОбновленииКонфигурации"),  # 07.10.2026
    ("Document", "А_ПотребаЗавершенняПроекту"),  # 08.10.2026
    ("Role", "А_ПотребаЗавершенняПроектуРедактирование"),  # 08.10.2026
    ("Report", "А_ПрогнозВиконанняССЗаПотребою"),  # 08.10.2026, этап встраивания отчёта
}

# Объекты, которые намеренно удаляются из базы загрузкой.
EXPECTED_REMOVED = set()

# Колекции, которые сверяются. Имя тега в Configuration.xml -> свойство Metadata.
COLLECTIONS = {
    "Catalog": "Catalogs", "Document": "Documents", "Enum": "Enums",
    "InformationRegister": "InformationRegisters", "AccumulationRegister": "AccumulationRegisters",
    "AccountingRegister": "AccountingRegisters", "DataProcessor": "DataProcessors",
    "Report": "Reports", "CommonModule": "CommonModules", "Role": "Roles",
    "EventSubscription": "EventSubscriptions", "CommandGroup": "CommandGroups",
    "Subsystem": "Subsystems", "Constant": "Constants", "CommonForm": "CommonForms",
    "DocumentJournal": "DocumentJournals", "ScheduledJob": "ScheduledJobs",
}


def config_objects(path):
    text = io.open(path, encoding="utf-8-sig").read()
    body = text[text.find("<ChildObjects>"):text.find("</ChildObjects>")]
    return {(kind, name) for kind, name in re.findall(r"<(\w+)>([^<]+)</\1>", body)
            if kind in COLLECTIONS}


def database_objects():
    buh = win32com.client.Dispatch("V83.COMConnector").Connect(CONNECTION)
    md = buh.Metadata
    found = set()
    for kind, prop in COLLECTIONS.items():
        coll = getattr(md, prop)
        for i in range(coll.Count()):
            found.add((kind, coll.Get(i).Name))
    return found


def main(config_path):
    cfg = config_objects(config_path)
    db = database_objects()

    removed = sorted((db - cfg) & EXPECTED_REMOVED)
    lost = sorted(db - cfg - EXPECTED_REMOVED)
    extra = sorted(cfg - db)
    unexpected = sorted(set(extra) - EXPECTED_NEW)
    new_present = sorted(set(extra) & EXPECTED_NEW)
    already_in_db = sorted(EXPECTED_NEW & db)

    print(f"Configuration.xml : {config_path}")
    print(f"объектов в базе   : {len(db)}")
    print(f"объектов в файле  : {len(cfg)}")
    print(f"новых объектов в файле: {len(new_present)} из {len(EXPECTED_NEW)}")
    if already_in_db:
        print(f"  (уже загружены в базу: {[n for _, n in already_in_db]})")
    if removed:
        print(f"НАМЕРЕННО будут удалены из базы: {[f'{k}.{n}' for k, n in removed]}")

    failures = []
    if lost:
        failures.append(f"{len(lost)} объектов есть в БАЗЕ, но НЕТ в файле — загрузка их УДАЛИТ:")
        failures.extend(f"    {k}.{n}" for k, n in lost)
    if unexpected:
        failures.append(f"{len(unexpected)} посторонних объектов в файле, которых нет в базе и нет в EXPECTED_NEW:")
        failures.extend(f"    {k}.{n}" for k, n in unexpected)

    print()
    print(f"ПРОВАЛОВ: {len(failures) and len(lost) + len(unexpected)}")
    if failures:
        for line in failures:
            print(line)
        print("\nЗАГРУЖАТЬ НЕЛЬЗЯ.")
        return 1
    if removed:
        print("Загружать можно: удалятся ТОЛЬКО объекты из EXPECTED_REMOVED, остальные объекты базы целы.")
    else:
        print("Загружать можно: ни один объект базы не будет удалён.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else MAIN_CONFIG))
