# -*- coding: utf-8 -*-
"""
Fix ObjectModule.bsl - remove non-existent BuhBud fields:
1. СчетНачисленияАмортизации.Код - not in ПараметрыАмортизацииОСБухгалтерскийУчет
2. НалоговаяГруппаОС.Наименование - not in ПервоначальныеСведенияОСБухгалтерскийУчет
"""

filepath = r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Перенос остатков основных средств\Ext\ObjectModule.bsl'

with open(filepath, 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()

original_count = len(lines)
print(f'Original file: {original_count} lines')

lines_to_remove = set()

for i, line in enumerate(lines):
    stripped = line.strip()

    # CHANGE 1: Outer SELECT - remove fields
    if 'ВЗ.СчетАмортизацииКод КАК СчетАмортизацииКод' in stripped:
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove outer SELECT СчетАмортизацииКод')

    if 'ВЗ.НалоговаяГруппаОСНаим КАК НалоговаяГруппаОСНаим' in stripped:
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove outer SELECT НалоговаяГруппаОСНаим')

    # CHANGE 2: Inner SELECT - remove fields
    if 'ПараметрыАморт.СчетНачисленияАмортизации.Код КАК СчетАмортизацииКод' in stripped:
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove inner SELECT СчетНачисленияАмортизации')

    if 'ПервоначСведения.НалоговаяГруппаОС.Наименование КАК НалоговаяГруппаОСНаим' in stripped:
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove inner SELECT НалоговаяГруппаОС')

    # CHANGE 3: Inner GROUP BY - remove fields
    if stripped == '|\t\tПараметрыАморт.СчетНачисленияАмортизации.Код,' or \
       stripped == 'ПараметрыАморт.СчетНачисленияАмортизации.Код,':
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove GROUP BY СчетНачисленияАмортизации')

    if stripped == '|\t\tПервоначСведения.НалоговаяГруппаОС.Наименование,' or \
       stripped == 'ПервоначСведения.НалоговаяГруппаОС.Наименование,':
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove GROUP BY НалоговаяГруппаОС')

    # CHANGE 4: Loading code - remove both lines
    if 'НоваяСтрока.СчетАмортизацииКод = СокрЛП(СтрокаТЗ.СчетАмортизацииКод)' in stripped:
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove loading СчетАмортизацииКод')

    if 'НоваяСтрока.НалоговаяГруппаОСНаим = СокрЛП(СтрокаТЗ.НалоговаяГруппаОСНаим)' in stripped:
        lines_to_remove.add(i)
        print(f'  Line {i+1}: Remove loading НалоговаяГруппаОСНаим')

print(f'\nLines to remove (changes 1-4): {len(lines_to_remove)}')

# CHANGE 5: Replace СчетАмортизации block with simplified version
change5_start = None
change5_end = None
for i, line in enumerate(lines):
    stripped = line.strip()
    if 'Рахунок амортизації' in stripped and 'BuhBud' in stripped:
        change5_start = i
    if change5_start is not None and change5_end is None and i > change5_start:
        if stripped == 'КонецЕсли;':
            change5_end = i
            break

replacement5 = []
if change5_start is not None and change5_end is not None:
    print(f'\nCHANGE 5: Replace lines {change5_start+1}-{change5_end+1} (СчетАмортизации block)')
    indent = lines[change5_start][:len(lines[change5_start]) - len(lines[change5_start].lstrip())]
    for j in range(change5_start, change5_end + 1):
        lines_to_remove.add(j)
    replacement5 = [
        indent + '// Рахунок амортизації — 131 (в BuhBud немає поля СчетНачисленияАмортизации)\n',
        indent + 'СчетАмортизацииЕРП = ПланыСчетов.Хозрасчетный.НайтиПоКоду("131");\n',
    ]
else:
    print(f'\nCHANGE 5 FAILED: block not found (start={change5_start}, end={change5_end})')

# CHANGE 6: Remove НалоговаяГруппаОС lookup block
change6_start = None
change6_end = None
nesting = 0
for i, line in enumerate(lines):
    stripped = line.strip()
    if 'НалоговаяГруппаОС' in stripped and 'пошук за найменуванням' in stripped:
        change6_start = i
        nesting = 0
        continue
    if change6_start is not None and change6_end is None and i > change6_start:
        if stripped.startswith('Если ') or stripped.startswith('Если\t'):
            nesting += 1
        if stripped == 'КонецЕсли;':
            nesting -= 1
            if nesting <= 0:
                change6_end = i
                break

if change6_start is not None and change6_end is not None:
    print(f'CHANGE 6: Remove lines {change6_start+1}-{change6_end+1} (НалоговаяГруппаОС lookup)')
    # Also remove preceding empty line if exists
    if change6_start > 0 and lines[change6_start - 1].strip() == '':
        lines_to_remove.add(change6_start - 1)
    for j in range(change6_start, change6_end + 1):
        lines_to_remove.add(j)
else:
    print(f'CHANGE 6 FAILED: block not found (start={change6_start}, end={change6_end})')

# Also remove СчетДооценокОСЕРП variable since we no longer need it? No, keep it - it's still used

# Build new content
new_lines = []
for i, line in enumerate(lines):
    if i in lines_to_remove:
        if change5_start is not None and i == change5_start:
            new_lines.extend(replacement5)
        continue
    new_lines.append(line)

# Write
with open(filepath, 'w', encoding='utf-8-sig') as f:
    f.writelines(new_lines)

new_count = len(new_lines)
print(f'\nNew file: {new_count} lines (was {original_count}, removed {original_count - new_count})')
print('File saved successfully!')
