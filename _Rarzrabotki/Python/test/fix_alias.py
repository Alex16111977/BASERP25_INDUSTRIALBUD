# -*- coding: utf-8 -*-
"""
Fix: rename alias Количество → КоличествоИтого to avoid COM method name collision.
V83_Рез.Количество conflicts with ВыборкаИзРезультатаЗапроса.Количество() method via COM.
"""
FILE_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьТовары\Ext\ObjectModule.bsl"

with open(FILE_PATH, 'r', encoding='utf-8-sig') as f:
    content = f.read()

replacements = [
    # === Main UNION ALL query ===
    # Outer SELECT alias
    ("ЕСТЬNULL(Т.Количество, 0)) КАК Количество", "ЕСТЬNULL(Т.Количество, 0)) КАК КоличествоИтого"),
    # Reading results - all 3 places where V83_Рез.Количество is used
    ("V83_Рез.Количество", "V83_Рез.КоличествоИтого"),

    # === Sub-query in ПроверитьДокументВBuhBud ===
    # Alias in query
    ("КоличествоКт, 0)) КАК Количество", "КоличествоКт, 0)) КАК КоличествоИтого"),
    # Reading sub-query result
    ("V83_РезКол.Количество", "V83_РезКол.КоличествоИтого"),
]

changes = []
for old, new in replacements:
    count = content.count(old)
    if count > 0:
        content = content.replace(old, new)
        changes.append(f"  '{old}' → '{new}' ({count} замен)")
    else:
        changes.append(f"  [!] НЕ НАЙДЕНО: '{old}'")

with open(FILE_PATH, 'w', encoding='utf-8-sig') as f:
    f.write(content)

print(f"Файл обновлён: {FILE_PATH}")
print(f"\nИзменения:")
for c in changes:
    print(c)
