# -*- coding: utf-8 -*-
# Клонирование обработки СинхронизироватьДеньги -> СинхронизироватьДеньгиКасса
# с переименованием полей/синонимов и регенерацией GUID. ObjectModule.bsl НЕ трогаем (перезапишем отдельно).
import sys, os, re, uuid, shutil
sys.stdout.reconfigure(encoding="utf-8")

WT = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\trusting-noyce-fbeddb\_Rarzrabotki\Обработки"
SRC_DIR = os.path.join(WT, "СинхронизироватьДеньги")
DST_DIR = os.path.join(WT, "СинхронизироватьДеньгиКасса")
SRC_XML = os.path.join(WT, "СинхронизироватьДеньги.xml")
DST_XML = os.path.join(WT, "СинхронизироватьДеньгиКасса.xml")

if os.path.exists(DST_DIR):
    shutil.rmtree(DST_DIR)
shutil.copytree(SRC_DIR, DST_DIR)
shutil.copyfile(SRC_XML, DST_XML)

CLASSID = "c3831ec8-d8d5-4f93-8a22-f9bfae07327f"  # платформенный ClassId внешней обработки — НЕ менять
GUID_RE = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")

def regen_guids(text):
    mapping = {}
    def repl(m):
        g = m.group(0)
        if g.lower() == CLASSID:
            return g
        if g not in mapping:
            mapping[g] = str(uuid.uuid4())
        return mapping[g]
    return GUID_RE.sub(repl, text), len(mapping)

# Порядок важен: длинные токены раньше коротких
SUBS = [
    ("БанковскиеСчетаОрганизаций", "Кассы"),
    ("ФильтрБанковскийСчет", "ФильтрКасса"),
    ("БанковскийСчетUID", "КассаUID"),
    ("БанковскийСчет", "Касса"),
    # идентификаторы-колонки Бух -> Казна
    ("НачОстатокБух", "НачОстатокКазна"),
    ("ПриходБух", "ПриходКазна"),
    ("РасходБух", "РасходКазна"),
    ("ОстатокБух", "ОстатокКазна"),
    ("СуммаБух", "СуммаКазна"),
    ("ДокументБух", "ДокументКазна"),
    ("UIDДокументаБух", "UIDДокументаКазна"),
    # видимые заголовки колонок
    ("остаток Бух", "остаток Казна"),
    ("залишок Бух", "залишок Казна"),
    ("Приход Бух", "Приход Казна"),
    ("Прихід Бух", "Прихід Казна"),
    ("Расход Бух", "Расход Казна"),
    ("Видаток Бух", "Видаток Казна"),
    ("Остаток Бух", "Остаток Казна"),
    ("Залишок Бух", "Залишок Казна"),
    ("Сумма Бух", "Сумма Казна"),
    ("Сума Бух", "Сума Казна"),
    ("Документ Бух", "Документ Казна"),
    ("UID документа Бух", "UID документа Казна"),
    ("Поч.Бух", "Поч.Казна"),
    ("Прих.Бух", "Прих.Казна"),
    ("Расх.Бух", "Расх.Казна"),
    ("Кін.Бух", "Кін.Казна"),
    # заголовки счёта -> касса
    ("UID банковского счета", "UID кассы"),
    ("UID банківського рахунку", "UID каси"),
    ("Банковский счет", "Касса"),
    ("Банківський рахунок", "Каса"),
    # резюме/тултипы база-источник
    ("ERP та BuhBud", "ERP та Казни"),
    ("ERP та Бухгалтерії", "ERP та Казни"),
    ("Бухгалтерії", "Казни"),
    ("BuhBud", "Казна"),
    # имя объекта (camelCase, без пробела)
    ("СинхронизироватьДеньги", "СинхронизироватьДеньгиКасса"),
    # синонимы (с пробелом)
    ("Синхронизировать деньги", "Синхронизировать деньги (касса)"),
    ("Синхронізувати гроші", "Синхронізувати гроші (каса)"),
]

# Файлы для текстовой трансформации (ObjectModule.bsl исключён — будет перезаписан)
TARGETS = [
    DST_XML,
    os.path.join(DST_DIR, "Forms", "Форма.xml"),
    os.path.join(DST_DIR, "Forms", "Форма", "Ext", "Form.xml"),
    os.path.join(DST_DIR, "Forms", "Форма", "Ext", "Form", "Module.bsl"),
]

for path in TARGETS:
    with open(path, "r", encoding="utf-8-sig") as f:
        text = f.read()
    before = text
    for a, b in SUBS:
        text = text.replace(a, b)
    guids_changed = 0
    if path.endswith(".xml"):
        text, guids_changed = regen_guids(text)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"OK {os.path.relpath(path, WT)} | subs_applied={'yes' if text!=before or guids_changed else 'no'} | guids_regen={guids_changed}")

# Контроль остаточных «опасных» токенов в трансформированных файлах
print("\n--- контроль остаточных токенов ---")
for path in TARGETS:
    with open(path, "r", encoding="utf-8") as f:
        t = f.read()
    leftovers = []
    for tok in ["БанковскийСчет", "БанковскиеСчетаОрганизаций", "СинхронизироватьДеньгиКассаКасса", "BuhBud"]:
        c = t.count(tok)
        if c:
            leftovers.append(f"{tok}={c}")
    print(f"{os.path.relpath(path, WT)}: {leftovers if leftovers else 'чисто'}")
print("\nГОТОВО. ObjectModule.bsl скопирован как есть — будет перезаписан отдельно.")
