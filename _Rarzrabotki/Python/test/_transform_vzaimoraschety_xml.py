# -*- coding: utf-8 -*-
# Детерминированный трансформер XML обработки СинхронизироватьВзаиморасчеты:
#  A. Внутр. переименование клона Товары→Взаиморасчеты (Name/DefaultForm/GeneratedType/Synonym)
#  B. Регенерация внутр. GUID (ObjectId + все TypeId/ValueId + внешний uuid) — анти-коллизия с сиблингом
#  C. Трансформация полей товары→взаиморасчёты (имена/типы/синонимы) + новые Валюта/ВалютаКод
import re, uuid, io, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PATH = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты.xml"
with io.open(PATH, encoding="utf-8") as f:
    t = f.read()
orig = t

def ng():
    return str(uuid.uuid4())

# ---------- A. внутреннее имя + объектный синоним ----------
assert t.count("СинхронизироватьТоварыТолькоТовары") >= 9
t = t.replace("СинхронизироватьТоварыТолькоТовары", "СинхронизироватьВзаиморасчеты")
t = t.replace("Синхронизировать ERP и BuhBud только товары",
              "Сверка взаиморасчётов ERP и BuhBud по контрагентам")
t = t.replace("Синхронізувати товари ERP і BuhBud",
              "Звірка взаєморозрахунків ERP і BuhBud по контрагентах")

# ---------- B. регенерация внутренних GUID (анти-коллизия с Товары) ----------
GUIDS = [
    "178f1b6f-5313-4d05-b6c4-70eba2a5aba9",  # внешний uuid обработки
    "043bfd56-aee6-4312-ad03-0eab596a9249",  # ObjectId
    "1c3624fa-e71c-473c-b294-f2c0ec03b4ea", "b118f123-aa98-4e08-ac44-e424462a0c0a",  # Object type
    "2b7520eb-1651-4bff-930f-7ad67ff7b244", "f0c0abd4-4352-4095-b356-04662b192fcb",  # Расх TS
    "2fac7e7c-98f8-47f1-afca-983d39f5e793", "3d59911e-917f-4c9a-8e88-c43d76dbefb9",  # Расх Row
    "091b01e4-c088-4626-ba61-2a01aa3f6631", "6534f21b-3226-4fc8-b044-6ca700027a3b",  # Док TS
    "88381444-56d6-4b7b-85f0-faca732903b3", "e98d9125-7c6b-4d5d-a1b9-91eb5460b104",  # Док Row
    "1981b681-8cf0-4e87-ab72-10e9369fb204", "cc960689-8f89-4b04-9acf-4fb3a3b7dae8",  # Скл TS
    "b54d59a7-9bfc-457d-a370-48077e88e0e8", "6780941f-138e-4f11-b37c-76923a768bde",  # Скл Row
]
for old in GUIDS:
    assert old in t, "GUID не найден: " + old
    t = t.replace(old, ng())

# ---------- C1. типы объектных реквизитов-фильтров ----------
assert "cfg:CatalogRef.Номенклатура" in t and "cfg:CatalogRef.Склады" in t
t = t.replace("cfg:CatalogRef.Номенклатура", "cfg:CatalogRef.Контрагенты")
t = t.replace("cfg:CatalogRef.Склады", "cfg:CatalogRef.ДоговорыКонтрагентов")

# ---------- C2. деньги: FractionDigits 3 -> 2 ----------
t = t.replace("<v8:FractionDigits>3</v8:FractionDigits>", "<v8:FractionDigits>2</v8:FractionDigits>")

# ---------- C3. имя ТЧ ТаблицаСкладов -> ТаблицаКонтрагентов ----------
assert t.count("ТаблицаСкладов") >= 3
t = t.replace("ТаблицаСкладов", "ТаблицаКонтрагентов")
t = t.replace("Таблица складов", "Таблица контрагентов").replace("Таблиця складів", "Таблиця контрагентів")

# ---------- C4. переименование реквизитов (Name+Synonym) по uuid-блокам ----------
NAME_MAP = {
    "b1a2c3d4-1004": ("Контрагент", "Контрагент", "Контрагент"),
    "b1a2c3d4-1005": ("Договор", "Договор", "Договір"),
    "c1a00001-0002": ("Контрагент", "Контрагент", "Контрагент"),
    "c1a00001-0003": ("Договор", "Договор", "Договір"),
    "c1a00001-000e": ("КонтрагентКлюч", "Контрагент ключ", "Контрагент ключ"),
    "c1a00001-000f": ("ДоговорКлюч", "Договор ключ", "Договір ключ"),
    "c2a00001-0002": ("Контрагент", "Контрагент", "Контрагент"),
    "c2a00001-0003": ("Договор", "Договор", "Договір"),
    "c2a00001-0007": ("СуммаЕРП", "Сумма ЕРП", "Сума ЕРП"),
    "c2a00001-0008": ("СуммаБух", "Сумма Бух", "Сума Бух"),
    "f2f19761": ("Контрагент", "Контрагент", "Контрагент"),
    "c43ae0f9": ("КонтрагентКлюч", "Контрагент ключ", "Контрагент ключ"),
    "51a0eefc": ("Договоров", "Договоров", "Договорів"),
}

def repl_attr(m):
    uid = m.group(1)
    block = m.group(0)
    for pref, (nn, ru, uk) in NAME_MAP.items():
        if uid.startswith(pref):
            block = re.sub(r"<Name>.*?</Name>", "<Name>%s</Name>" % nn, block, count=1, flags=re.S)
            conts = re.findall(r"<v8:content>.*?</v8:content>", block, flags=re.S)
            newc = ["<v8:content>%s</v8:content>" % ru, "<v8:content>%s</v8:content>" % uk]
            for i, c in enumerate(conts[:2]):
                block = block.replace(c, newc[i], 1)
            return block
    return block

t = re.sub(r'<Attribute uuid="([^"]+)">.*?</Attribute>', repl_attr, t, flags=re.S)

# ---------- C5. новые реквизиты Валюта (расх.) и ВалютаКод (расх.) клонированием ----------
def find_block(text, uuid_prefix):
    m = re.search(r'<Attribute uuid="' + re.escape(uuid_prefix) + r'[^"]*">.*?</Attribute>', text, flags=re.S)
    assert m, "блок не найден: " + uuid_prefix
    return m.group(0), m.end()

# Валюта — клон блока Договор (c1a00001-0003, составной string+AnyRef)
dog_block, dog_end = find_block(t, "c1a00001-0003")
val_block = dog_block
val_block = re.sub(r'uuid="[^"]+"', 'uuid="%s"' % ng(), val_block, count=1)
val_block = re.sub(r"<Name>.*?</Name>", "<Name>Валюта</Name>", val_block, count=1, flags=re.S)
for c in re.findall(r"<v8:content>.*?</v8:content>", val_block, flags=re.S)[:2]:
    val_block = val_block.replace(c, "<v8:content>Валюта</v8:content>", 1)
t = t[:dog_end] + "\n\t\t\t\t\t" + val_block + t[dog_end:]

# ВалютаКод — клон блока ДоговорКлюч (c1a00001-000f, строка 36)
dk_block, dk_end = find_block(t, "c1a00001-000f")
vk_block = dk_block
vk_block = re.sub(r'uuid="[^"]+"', 'uuid="%s"' % ng(), vk_block, count=1)
vk_block = re.sub(r"<Name>.*?</Name>", "<Name>ВалютаКод</Name>", vk_block, count=1, flags=re.S)
for c in re.findall(r"<v8:content>.*?</v8:content>", vk_block, flags=re.S)[:2]:
    vk_block = vk_block.replace(c, "<v8:content>Валюта код</v8:content>", 1)
vk_block = vk_block.replace("<v8:Length>36</v8:Length>", "<v8:Length>10</v8:Length>")
t = t[:dk_end] + "\n\t\t\t\t\t" + vk_block + t[dk_end:]

# ---------- проверки ----------
assert "СинхронизироватьТоварыТолькоТовары" not in t
assert "ТаблицаСкладов" not in t
assert "ТаблицаКонтрагентов" in t
assert "<Name>Валюта</Name>" in t and "<Name>ВалютаКод</Name>" in t
assert "<Name>Контрагент</Name>" in t and "<Name>Договор</Name>" in t
assert "<Name>СуммаЕРП</Name>" in t and "<Name>СуммаБух</Name>" in t
assert "<Name>Договоров</Name>" in t
assert "<Name>Номенклатура</Name>" not in t and "<Name>Склад</Name>" not in t
assert "<Name>КоличествоЕРП</Name>" not in t
assert "cfg:CatalogRef.Контрагенты" in t and "cfg:CatalogRef.ДоговорыКонтрагентов" in t
for old in GUIDS:
    assert old not in t, "GUID не заменён: " + old
assert "<v8:FractionDigits>3</v8:FractionDigits>" not in t

with io.open(PATH, "w", encoding="utf-8", newline="\n") as f:
    f.write(t)
print("OK transform: было %d симв -> стало %d симв" % (len(orig), len(t)))
print("   Контрагент x", t.count("<Name>Контрагент</Name>"),
      " Договор x", t.count("<Name>Договор</Name>"),
      " Валюта x", t.count("<Name>Валюта</Name>"),
      " ВалютаКод x", t.count("<Name>ВалютаКод</Name>"))
