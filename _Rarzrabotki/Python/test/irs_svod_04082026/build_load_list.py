# -*- coding: utf-8 -*-
"""Список файлів для часткового завантаження 16 відновлюваних об'єктів.

Правила гранулярності (з бази знань проєкту):
  - корінь об'єкта: <Каталог>/<Имя>.xml
  - модулі: .../Ext/*.bsl прямим шляхом
  - форми: ТІЛЬКИ кореневий .../Forms/<Форма>.xml (платформа підтягне Ext сама)
  - макети: ТІЛЬКИ кореневий .../Templates/<Макет>.xml
"""
import io, os, sys, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

ROOT = r"C:\Configuration_downloads\BASERP25"

OBJECTS = [
    ("Catalogs", "А_КонтурыСверкиБаз"),
    ("CommonModules", "А_ВыравниваниеБазСервер"),
    ("DataProcessors", "А_СинхронизироватьВзаиморасчеты"),
    ("DataProcessors", "А_СинхронизироватьДеньги"),
    ("DataProcessors", "А_СинхронизироватьДеньгиКасса"),
    ("DataProcessors", "А_СинхронизироватьТоварыТолькоТовары"),
    ("DataProcessors", "А_ЦентрВыравниванияБаз"),
    ("DataProcessors", "СозданиеСотрудниковИзФизическихЛиц"),
    ("Documents", "А_Отчет_ПланФактныйПроизводство"),
    ("Enums", "А_СтатусыВыравниванияБаз"),
    ("Enums", "А_ТипыСтрокПланФактПроизводство"),
    ("InformationRegisters", "А_ЖурналВыравниванияБаз"),
    ("InformationRegisters", "А_ИсключенияСверкиБаз"),
    ("InformationRegisters", "А_ПланФактПроизводство_Свод"),
    ("Reports", "А_ОтчетПоВыполнениюРаботвПроизводстве"),
    ("Reports", "А_ПланФактныйПроизводствоСвод"),
]

files, problems = [], []

# перевірка валідності XML усіх коренів + Configuration.xml
for path in [os.path.join(ROOT, "Configuration.xml")] + \
            [os.path.join(ROOT, d, n + ".xml") for d, n in OBJECTS]:
    try:
        ET.parse(path)
    except Exception as e:
        problems.append("XML НЕВАЛІДНИЙ: %s -> %s" % (path, e))

for d, n in OBJECTS:
    root_xml = os.path.join(ROOT, d, n + ".xml")
    if not os.path.exists(root_xml):
        problems.append("немає кореня: %s" % root_xml)
        continue
    files.append("%s/%s.xml" % (d, n))

    objdir = os.path.join(ROOT, d, n)
    if not os.path.isdir(objdir):
        continue

    ext = os.path.join(objdir, "Ext")
    if os.path.isdir(ext):
        for f in sorted(os.listdir(ext)):
            if f.lower().endswith(".bsl"):
                files.append("%s/%s/Ext/%s" % (d, n, f))

    forms = os.path.join(objdir, "Forms")
    if os.path.isdir(forms):
        for f in sorted(os.listdir(forms)):
            if f.lower().endswith(".xml"):
                files.append("%s/%s/Forms/%s" % (d, n, f))

    tmpl = os.path.join(objdir, "Templates")
    if os.path.isdir(tmpl):
        for f in sorted(os.listdir(tmpl)):
            if f.lower().endswith(".xml"):
                files.append("%s/%s/Templates/%s" % (d, n, f))

files.append("Configuration.xml")

print("файлів для завантаження: %d" % len(files))
for f in files:
    print("   " + f)
print("\nпроблеми: %s" % (problems if problems else "немає"))

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "load_list_restore.txt")
io.open(out, "w", encoding="utf-8-sig", newline="\r\n").write("\n".join(files) + "\n")
print("\n-> %s" % out)
