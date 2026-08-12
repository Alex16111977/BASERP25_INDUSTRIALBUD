# -*- coding: utf-8 -*-
"""Аудит: які об'єкти лежать файлами на диску, але НЕ зареєстровані в Configuration.xml.

Причина аварії 05.08.2026: часткове завантаження з Configuration.xml, у якому
не було реєстрації 7 об'єктів, призвело до їх видалення з бази.
"""
import io, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

ROOT = r"C:\Configuration_downloads\BASERP25"
CFG = os.path.join(ROOT, "Configuration.xml")
cfg = io.open(CFG, encoding="utf-8-sig").read()

# каталог на диску -> тег у Configuration.xml
DIRS = {
    "Catalogs": "Catalog", "Documents": "Document", "Enums": "Enum",
    "InformationRegisters": "InformationRegister",
    "AccumulationRegisters": "AccumulationRegister",
    "Reports": "Report", "DataProcessors": "DataProcessor",
    "CommonModules": "CommonModule", "ChartsOfCharacteristicTypes": "ChartOfCharacteristicTypes",
    "DocumentJournals": "DocumentJournal", "ExchangePlans": "ExchangePlan",
    "Constants": "Constant", "Roles": "Role", "Subsystems": "Subsystem",
    "AccountingRegisters": "AccountingRegister", "CalculationRegisters": "CalculationRegister",
    "BusinessProcesses": "BusinessProcess", "Tasks": "Task",
    "ChartsOfAccounts": "ChartOfAccounts", "ChartsOfCalculationTypes": "ChartOfCalculationTypes",
    "DefinedTypes": "DefinedType", "CommonForms": "CommonForm",
    "CommonTemplates": "CommonTemplate", "CommonCommands": "CommonCommand",
    "EventSubscriptions": "EventSubscription", "ScheduledJobs": "ScheduledJob",
    "FunctionalOptions": "FunctionalOption",
    "FunctionalOptionsParameters": "FunctionalOptionsParameter",
    "SettingsStorages": "SettingsStorage", "FilterCriteria": "FilterCriterion",
    "CommonPictures": "CommonPicture", "SessionParameters": "SessionParameter",
    "WSReferences": "WSReference", "WebServices": "WebService",
    "HTTPServices": "HTTPService", "Sequences": "Sequence",
    "CommonAttributes": "CommonAttribute", "XDTOPackages": "XDTOPackage",
    "StyleItems": "StyleItem", "Languages": "Language",
    "ExternalDataSources": "ExternalDataSource", "Styles": "Style",
    "Interfaces": "Interface", "Bots": "Bot", "IntegrationServices": "IntegrationService",
}

registered = {}
for tag in set(DIRS.values()):
    registered[tag] = set(re.findall(r"<%s>([^<]+)</%s>" % (tag, tag), cfg))

missing_total = []
for d, tag in sorted(DIRS.items()):
    path = os.path.join(ROOT, d)
    if not os.path.isdir(path):
        continue
    on_disk = {f[:-4] for f in os.listdir(path) if f.endswith(".xml")}
    miss = sorted(on_disk - registered.get(tag, set()))
    if miss:
        print("%-30s на диску %4d, не зареєстровано %d:" % (d, len(on_disk), len(miss)))
        for m in miss:
            fp = os.path.join(path, m + ".xml")
            sz = os.path.getsize(fp)
            import datetime
            mt = datetime.datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M")
            print("      %-46s %8d байт  %s" % (m, sz, mt))
            missing_total.append((d, tag, m))

print("\nВСЬОГО не зареєстровано: %d" % len(missing_total))
if missing_total:
    print("\nРядки для вставки в Configuration.xml:")
    for d, tag, m in missing_total:
        print("\t\t\t<%s>%s</%s>" % (tag, m, tag))
