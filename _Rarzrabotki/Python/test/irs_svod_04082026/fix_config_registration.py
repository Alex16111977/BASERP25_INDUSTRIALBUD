# -*- coding: utf-8 -*-
"""Відновлення реєстрації об'єктів у Configuration.xml.

Вставляє <Tag>Имя</Tag> одразу після ОСТАННЬОГО наявного запису того самого тега,
щоб зберегти групування. Робить резервну копію перед правкою.
"""
import io, os, re, sys, shutil, datetime
sys.stdout.reconfigure(encoding='utf-8')

ROOT = r"C:\Configuration_downloads\BASERP25"
CFG = os.path.join(ROOT, "Configuration.xml")

MISSING = [
    ("Catalog", "А_КонтурыСверкиБаз"),
    ("CommonModule", "А_ВыравниваниеБазСервер"),
    ("DataProcessor", "А_СинхронизироватьВзаиморасчеты"),
    ("DataProcessor", "А_СинхронизироватьДеньги"),
    ("DataProcessor", "А_СинхронизироватьДеньгиКасса"),
    ("DataProcessor", "А_СинхронизироватьТоварыТолькоТовары"),
    ("DataProcessor", "А_ЦентрВыравниванияБаз"),
    ("DataProcessor", "СозданиеСотрудниковИзФизическихЛиц"),
    ("Document", "А_Отчет_ПланФактныйПроизводство"),
    ("Enum", "А_СтатусыВыравниванияБаз"),
    ("Enum", "А_ТипыСтрокПланФактПроизводство"),
    ("InformationRegister", "А_ЖурналВыравниванияБаз"),
    ("InformationRegister", "А_ИсключенияСверкиБаз"),
    ("InformationRegister", "А_ПланФактПроизводство_Свод"),
    ("Report", "А_ОтчетПоВыполнениюРаботвПроизводстве"),
    ("Report", "А_ПланФактныйПроизводствоСвод"),
]

raw = io.open(CFG, encoding="utf-8-sig").read()
bak = CFG + ".bak_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
shutil.copy2(CFG, bak)
print("резервна копія: %s" % os.path.basename(bak))

lines = raw.split("\n")
added, skipped = 0, 0

for tag, name in MISSING:
    entry = "<%s>%s</%s>" % (tag, name, tag)
    if any(entry in ln for ln in lines):
        print("  вже є: %s" % entry)
        skipped += 1
        continue
    # останній рядок із цим тегом
    idx = None
    for i, ln in enumerate(lines):
        if re.match(r"\s*<%s>[^<]+</%s>\s*$" % (tag, tag), ln):
            idx = i
    if idx is None:
        print("  !! немає жодного <%s> — пропущено %s" % (tag, name))
        continue
    indent = re.match(r"(\s*)", lines[idx]).group(1)
    lines.insert(idx + 1, indent + entry)
    added += 1
    print("  + %-24s після рядка %d" % (entry, idx + 1))

io.open(CFG, "w", encoding="utf-8-sig", newline="").write("\n".join(lines))
print("\nдодано: %d, пропущено (вже було): %d" % (added, skipped))

# контроль
chk = io.open(CFG, encoding="utf-8-sig").read()
bad = [(t, n) for t, n in MISSING if "<%s>%s</%s>" % (t, n, t) not in chk]
print("не знайдено після правки: %s" % (bad if bad else "немає — усі 16 зареєстровані"))
for t, n in (("Document", "А_ГрафикОплатМатериалов"),
             ("AccumulationRegister", "А_ПланОплатМатериалов")):
    print("  наш %-22s : %s" % (n, "є" if "<%s>%s</%s>" % (t, n, t) in chk else "НЕМАЄ"))
