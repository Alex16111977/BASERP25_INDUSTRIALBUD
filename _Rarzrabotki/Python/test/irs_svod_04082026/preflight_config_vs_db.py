# -*- coding: utf-8 -*-
"""PRE-FLIGHT перед завантаженням Configuration.xml.

Головна перевірка: що є в БАЗІ, але НЕ зареєстровано в Configuration.xml.
Саме такі об'єкти платформа ВИДАЛЯЄ разом із даними (аварія 05.08.2026).

Друга перевірка: що зареєстровано у файлі, але файла об'єкта на диску немає
(завантаження впаде).
"""
import io, os, re, sys
import win32com.client
sys.stdout.reconfigure(encoding='utf-8')

ROOT = r"C:\Configuration_downloads\BASERP25"
CFG = os.path.join(ROOT, "Configuration.xml")

# колекція метаданих 1С -> (тег Configuration.xml, каталог на диску)
MAP = [
    ("Справочники",            "Catalog",              "Catalogs"),
    ("Документы",              "Document",             "Documents"),
    ("Перечисления",           "Enum",                 "Enums"),
    ("РегистрыСведений",       "InformationRegister",  "InformationRegisters"),
    ("РегистрыНакопления",     "AccumulationRegister", "AccumulationRegisters"),
    ("РегистрыБухгалтерии",    "AccountingRegister",   "AccountingRegisters"),
    ("РегистрыРасчета",        "CalculationRegister",  "CalculationRegisters"),
    ("Отчеты",                 "Report",               "Reports"),
    ("Обработки",              "DataProcessor",        "DataProcessors"),
    ("ОбщиеМодули",            "CommonModule",         "CommonModules"),
    ("Константы",              "Constant",             "Constants"),
    ("ПланыВидовХарактеристик","ChartOfCharacteristicTypes", "ChartsOfCharacteristicTypes"),
    ("ПланыСчетов",            "ChartOfAccounts",      "ChartsOfAccounts"),
    ("ПланыВидовРасчета",      "ChartOfCalculationTypes", "ChartsOfCalculationTypes"),
    ("ПланыОбмена",            "ExchangePlan",         "ExchangePlans"),
    ("ЖурналыДокументов",      "DocumentJournal",      "DocumentJournals"),
    ("БизнесПроцессы",         "BusinessProcess",      "BusinessProcesses"),
    ("Задачи",                 "Task",                 "Tasks"),
    ("Роли",                   "Role",                 "Roles"),
    ("Подсистемы",             "Subsystem",            "Subsystems"),
    ("ОпределяемыеТипы",       "DefinedType",          "DefinedTypes"),
    ("ОбщиеФормы",             "CommonForm",           "CommonForms"),
    ("ОбщиеМакеты",            "CommonTemplate",       "CommonTemplates"),
    ("ОбщиeКоманды",           "CommonCommand",        "CommonCommands"),
    ("ПодпискиНаСобытия",      "EventSubscription",    "EventSubscriptions"),
    ("РегламентныеЗадания",    "ScheduledJob",         "ScheduledJobs"),
    ("ФункциональныеОпции",    "FunctionalOption",     "FunctionalOptions"),
    ("КритерииОтбора",         "FilterCriterion",      "FilterCriteria"),
    ("Последовательности",     "Sequence",             "Sequences"),
    ("ОбщиеРеквизиты",         "CommonAttribute",      "CommonAttributes"),
    ("ПакетыXDTO",             "XDTOPackage",          "XDTOPackages"),
    ("WebСервисы",             "WebService",           "WebServices"),
    ("HTTPСервисы",            "HTTPService",          "HTTPServices"),
    ("ХранилищаНастроек",      "SettingsStorage",      "SettingsStorages"),
    ("ПараметрыСеанса",        "SessionParameter",     "SessionParameters"),
]

cfg = io.open(CFG, encoding="utf-8-sig").read()
reg = {}
for _, tag, _ in MAP:
    reg[tag] = set(re.findall(r"<%s>([^<]+)</%s>" % (tag, tag), cfg))

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
md = erp.Метаданные

in_db_not_registered = []
registered_no_file = []
in_db_total = 0

for coll, tag, folder in MAP:
    try:
        c = getattr(md, coll)
    except Exception:
        continue
    names = set()
    for i in range(c.Количество()):
        names.add(S(c.Получить(i).Имя))
    in_db_total += len(names)

    miss = sorted(names - reg.get(tag, set()))
    for m in miss:
        in_db_not_registered.append((tag, m, coll))

    # зареєстровано, але файла немає
    path = os.path.join(ROOT, folder)
    for r in sorted(reg.get(tag, set())):
        if not os.path.exists(os.path.join(path, r + ".xml")):
            registered_no_file.append((tag, r))

print("об'єктів у базі: %d" % in_db_total)
print("\n" + "=" * 78)
print("1. Є В БАЗІ, АЛЕ НЕ ЗАРЕЄСТРОВАНО -> БУДЕ ВИДАЛЕНО ПРИ ЗАВАНТАЖЕННІ")
print("=" * 78)
if in_db_not_registered:
    for tag, name, coll in in_db_not_registered:
        print("   !!! %-24s %s" % (tag, name))
    print("\n   РАЗОМ: %d  -> ЗАВАНТАЖУВАТИ НЕ МОЖНА" % len(in_db_not_registered))
else:
    print("   немає — завантаження безпечне")

print("\n" + "=" * 78)
print("2. ЗАРЕЄСТРОВАНО, АЛЕ ФАЙЛА НЕМАЄ -> завантаження впаде")
print("=" * 78)
if registered_no_file:
    for tag, name in registered_no_file:
        print("   !!! %-24s %s" % (tag, name))
    print("\n   РАЗОМ: %d" % len(registered_no_file))
else:
    print("   немає")

ok = not in_db_not_registered and not registered_no_file
print("\n" + ("PRE-FLIGHT ПРОЙДЕНО — можна завантажувати"
              if ok else "!!! PRE-FLIGHT ПРОВАЛЕНО — НЕ ЗАВАНТАЖУВАТИ !!!"))
sys.exit(0 if ok else 1)
