# -*- coding: utf-8 -*-
"""
ОДНОРАЗОВЫЙ скаффолд (v1, 2026-06-10): генерация XML обработки СинхронизироватьДеньги
из исходника СинхронизироватьТовары.
НЕ ПЕРЕЗАПУСКАТЬ: начиная с v2 описатель и Form.xml дорабатываются напрямую
(фильтр по счёту, порядок/видимость колонок, итоги, обработчики) — перегенерация
затрёт эти правки. Оставлен для истории происхождения файлов.
"""
import re
import sys
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

SRC_ROOT = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки")
OUT_ROOT = Path(sys.argv[1]) if len(sys.argv) > 1 else SRC_ROOT

SRC_DESC = SRC_ROOT / "СинхронизироватьТовары.xml"
SRC_FORM_DESC = SRC_ROOT / "СинхронизироватьТовары" / "Forms" / "Форма.xml"
SRC_FORM = SRC_ROOT / "СинхронизироватьТовары" / "Forms" / "Форма" / "Ext" / "Form.xml"

OUT_DESC = OUT_ROOT / "СинхронизироватьДеньги.xml"
OUT_DIR = OUT_ROOT / "СинхронизироватьДеньги"
OUT_FORM_DESC = OUT_DIR / "Forms" / "Форма.xml"
OUT_FORM = OUT_DIR / "Forms" / "Форма" / "Ext" / "Form.xml"


def read(p):
    return p.read_text(encoding='utf-8-sig')


def write(p, text):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding='utf-8-sig', newline='\r\n')
    ET.parse(p)  # well-formed или исключение
    print(f"OK: {p}")


def sub_count(text, pattern, repl, expected, flags=0, label=""):
    new, n = re.subn(pattern, repl, text, flags=flags)
    assert n == expected, f"{label}: замен {n}, ожидалось {expected}"
    return new


def synonym_block(ru, uk):
    return (
        "<Synonym>\n"
        f"\t\t\t\t\t\t\t\t<v8:item>\n"
        f"\t\t\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n"
        f"\t\t\t\t\t\t\t\t\t<v8:content>{ru}</v8:content>\n"
        f"\t\t\t\t\t\t\t\t</v8:item>\n"
        f"\t\t\t\t\t\t\t\t<v8:item>\n"
        f"\t\t\t\t\t\t\t\t\t<v8:lang>uk</v8:lang>\n"
        f"\t\t\t\t\t\t\t\t\t<v8:content>{uk}</v8:content>\n"
        f"\t\t\t\t\t\t\t\t</v8:item>\n"
        f"\t\t\t\t\t\t\t</Synonym>"
    )


# ============================ ОПИСАТЕЛЬ ============================

DELETE_ATTRS = {
    "РазделТовары", "РазделДеньги", "РазделВзаиморасчеты", "Номенклатура", "Склад",
    "Раздел", "Измерение2", "Измерение2Ссылка",
}
# имя -> (новое имя, синоним ru, синоним uk, патч типа AnyRef->БанковскиеСчетаОрганизаций)
RENAME_ATTRS = {
    "Измерение1": ("БанковскийСчет", "Банковский счет", "Банківський рахунок", True),
    "Измерение1Ссылка": ("БанковскийСчетUID", "UID банковского счета", "UID банківського рахунку", False),
    "КоличествоЕРП": ("СуммаЕРП", "Сумма ЕРП", "Сума ЕРП", False),
    "КоличествоБух": ("СуммаБух", "Сумма Бух", "Сума Бух", False),
}

text = read(SRC_DESC)
deleted, renamed = [], []


def attr_handler(m):
    block = m.group(0)
    name_m = re.search(r"<Name>([^<]+)</Name>", block)
    name = name_m.group(1)
    if name in DELETE_ATTRS:
        deleted.append(name)
        return ""
    if name in RENAME_ATTRS:
        new_name, ru, uk, patch_type = RENAME_ATTRS[name]
        block = block.replace(f"<Name>{name}</Name>", f"<Name>{new_name}</Name>", 1)
        block = re.sub(r"<Synonym>.*?</Synonym>", lambda _: synonym_block(ru, uk), block, count=1, flags=re.S)
        if patch_type:
            block, n = re.subn(
                r"<v8:TypeSet>cfg:AnyRef</v8:TypeSet>",
                "<v8:Type>cfg:CatalogRef.БанковскиеСчетаОрганизаций</v8:Type>",
                block)
            assert n == 1, f"{name}: патч типа AnyRef: замен {n}"
        if name in ("КоличествоЕРП", "КоличествоБух"):
            # деньги — 2 знака после запятой (у Количество* было 3)
            block, n = re.subn(
                r"<v8:FractionDigits>3</v8:FractionDigits>",
                "<v8:FractionDigits>2</v8:FractionDigits>",
                block)
            assert n == 1, f"{name}: патч FractionDigits: замен {n}"
        renamed.append(f"{name}->{new_name}")
    return block


text = re.sub(r"[ \t]*<Attribute uuid=\"[0-9a-f-]+\">.*?</Attribute>\r?\n", attr_handler, text, flags=re.S)
# Раздел и КоличествоЕРП/Бух встречаются по разу на ТЧ: Раздел x2, остальные согласно структуре
assert sorted(deleted) == sorted(
    ["РазделТовары", "РазделДеньги", "РазделВзаиморасчеты", "Номенклатура", "Склад",
     "Раздел", "Раздел", "Измерение2", "Измерение2", "Измерение2Ссылка"]), f"удалено: {deleted}"
assert sorted(renamed) == sorted(
    ["Измерение1->БанковскийСчет", "Измерение1->БанковскийСчет",
     "Измерение1Ссылка->БанковскийСчетUID",
     "КоличествоЕРП->СуммаЕРП", "КоличествоБух->СуммаБух"]), f"переименовано: {renamed}"

n_renames = text.count("СинхронизироватьТовары")
assert n_renames > 0
text = text.replace("СинхронизироватьТовары", "СинхронизироватьДеньги")

# Корневой синоним (первый Synonym после корневого Name)
text = sub_count(
    text,
    r"(<Name>СинхронизироватьДеньги</Name>\s*)<Synonym>.*?</Synonym>",
    lambda m: m.group(1) + synonym_block("Синхронизировать деньги", "Синхронізувати гроші"),
    1, flags=re.S, label="корневой синоним")

# Новые UUID: ObjectId, TypeId, ValueId, uuid="..." (ClassId НЕ трогаем — общий для формата)
old_uuids = set()
for pat in (r"<xr:ObjectId>([0-9a-f-]{36})</xr:ObjectId>",
            r"<xr:TypeId>([0-9a-f-]{36})</xr:TypeId>",
            r"<xr:ValueId>([0-9a-f-]{36})</xr:ValueId>",
            r"uuid=\"([0-9a-f-]{36})\""):
    old_uuids.update(re.findall(pat, text))
for old in sorted(old_uuids):
    text = text.replace(old, str(uuid.uuid4()))
print(f"UUID заменено: {len(old_uuids)}")

write(OUT_DESC, text)
names = re.findall(r"<Name>([^<]+)</Name>", text)
print("Имена в описателе:", ", ".join(names))

# ============================ Forms/Форма.xml ============================

text = read(SRC_FORM_DESC)
old = re.search(r"<Form uuid=\"([0-9a-f-]{36})\">", text).group(1)
text = text.replace(old, str(uuid.uuid4()))
write(OUT_FORM_DESC, text)

# ============================ Form.xml ============================

text = read(SRC_FORM)

ops = [
    # (pattern, repl, expected, flags)
    (r"[ \t]*<CheckBoxField name=\"РазделТовары\".*?</CheckBoxField>\r?\n", "", 1, re.S),
    (r"[ \t]*<CheckBoxField name=\"РазделДеньги\".*?</CheckBoxField>\r?\n", "", 1, re.S),
    (r"[ \t]*<CheckBoxField name=\"РазделВзаиморасчеты\".*?</CheckBoxField>\r?\n", "", 1, re.S),
    (r"<v8:content>Розділи перевірки</v8:content>", "<v8:content>Налаштування</v8:content>", 1, 0),
    (r"[ \t]*<UsualGroup name=\"ГруппаФильтрыТоваров\".*?</UsualGroup>\r?\n", "", 1, re.S),
    (r"[ \t]*<InputField name=\"ТРРаздел\".*?</InputField>\r?\n", "", 1, re.S),
    (r"[ \t]*<InputField name=\"ТРИзмерение2\".*?</InputField>\r?\n", "", 1, re.S),
    (r"[ \t]*<InputField name=\"ТДРаздел\".*?</InputField>\r?\n", "", 1, re.S),
    (r"ТРИзмерение1", "ТРБанковскийСчет", 3, 0),
    (r"Объект\.ТаблицаРасхождений\.Измерение1<", "Объект.ТаблицаРасхождений.БанковскийСчет<", 1, 0),
    (r"ТДИзмерение1", "ТДБанковскийСчет", 3, 0),
    (r"Объект\.ТаблицаДокументов\.Измерение1<", "Объект.ТаблицаДокументов.БанковскийСчет<", 1, 0),
    (r"ТДКоличествоЕРП", "ТДСуммаЕРП", 3, 0),
    (r"Объект\.ТаблицаДокументов\.КоличествоЕРП<", "Объект.ТаблицаДокументов.СуммаЕРП<", 1, 0),
    (r"<v8:content>К-ть ЕРП</v8:content>", "<v8:content>Сума ЕРП</v8:content>", 1, 0),
    (r"ТДКоличествоБух", "ТДСуммаБух", 3, 0),
    (r"Объект\.ТаблицаДокументов\.КоличествоБух<", "Объект.ТаблицаДокументов.СуммаБух<", 1, 0),
    (r"<v8:content>К-ть Бух</v8:content>", "<v8:content>Сума Бух</v8:content>", 1, 0),
    (r"<v8:content>Вимір 1</v8:content>", "<v8:content>Банківський рахунок</v8:content>", 2, 0),
    (r"[ \t]*<Field>Объект\.Номенклатура</Field>\r?\n", "", 1, 0),
    (r"[ \t]*<Field>Объект\.РазделВзаиморасчеты</Field>\r?\n", "", 1, 0),
    (r"[ \t]*<Field>Объект\.РазделДеньги</Field>\r?\n", "", 1, 0),
    (r"[ \t]*<Field>Объект\.РазделТовары</Field>\r?\n", "", 1, 0),
    (r"[ \t]*<Field>Объект\.Склад</Field>\r?\n", "", 1, 0),
    (r"СинхронизироватьТовары", "СинхронизироватьДеньги", 1, 0),
]
for pat, repl, expected, flags in ops:
    text = sub_count(text, pat, repl, expected, flags=flags, label=pat[:40])

write(OUT_FORM, text)
print("ГОТОВО")
