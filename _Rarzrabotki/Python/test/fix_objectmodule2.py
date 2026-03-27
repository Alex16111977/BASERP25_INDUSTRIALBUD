# -*- coding: utf-8 -*-
"""
Fix ObjectModule.bsl - add correct BuhBud sources for:
1. СчетНачисленияАмортизации → from РегистрСведений.СчетаБухгалтерскогоУчетаОС
2. НалоговаяГруппаОС → from РегистрСведений.ПервоначальныеСведенияОСНалоговыйУчет
"""

filepath = r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Перенос остатков основных средств\Ext\ObjectModule.bsl'

with open(filepath, 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()

original_count = len(lines)
print(f'Original file: {original_count} lines')

new_lines = []
changes = 0

i = 0
while i < len(lines):
    line = lines[i]
    stripped = line.strip()

    # ====== CHANGE 1: Inner SELECT - add 2 fields after СрокИспользования ======
    if 'ПараметрыАморт.СрокПолезногоИспользования КАК СрокИспользования,' in stripped:
        new_lines.append(line)
        indent = line[:len(line) - len(line.lstrip())]
        # Add 2 new fields
        new_lines.append(indent + 'СчетаУчетаОС.СчетНачисленияАмортизации.Код КАК СчетАмортизацииКод,\n')
        new_lines.append(indent + 'ПервоначСведенияНУ.НалоговаяГруппаОС.Наименование КАК НалоговаяГруппаОСНаим,\n')
        changes += 1
        print(f'  CHANGE 1: Added 2 fields to inner SELECT after line {i+1}')
        i += 1
        continue

    # ====== CHANGE 2: Add 2 LEFT JOINs after ПараметрыАморт JOIN ======
    # Find the line with the second condition of ПараметрыАморт JOIN
    if 'ОстаткиОС.Организация = ПараметрыАморт.Организация' in stripped:
        new_lines.append(line)
        indent = line[:len(line) - len(line.lstrip())]
        # Determine the indent for JOIN keywords (2 tabs)
        join_indent = '\t\t'
        on_indent = '\t\t'
        # Add LEFT JOIN for СчетаБухгалтерскогоУчетаОС
        new_lines.append(join_indent + '|\t\tЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.СчетаБухгалтерскогоУчетаОС.СрезПоследних(\n')
        new_lines.append(join_indent + '|\t\t\t&Дата,\n')
        new_lines.append(join_indent + '|\t\t) КАК СчетаУчетаОС\n')
        new_lines.append(join_indent + '|\t\tПО ВЫРАЗИТЬ(ОстаткиОС.Субконто1 КАК Справочник.ОсновныеСредства) = СчетаУчетаОС.ОсновноеСредство\n')
        new_lines.append(join_indent + '|\t\t\tИ ОстаткиОС.Организация = СчетаУчетаОС.Организация\n')
        # Add LEFT JOIN for ПервоначальныеСведенияОСНалоговыйУчет
        new_lines.append(join_indent + '|\t\tЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ПервоначальныеСведенияОСНалоговыйУчет.СрезПоследних(\n')
        new_lines.append(join_indent + '|\t\t\t&Дата,\n')
        new_lines.append(join_indent + '|\t\t) КАК ПервоначСведенияНУ\n')
        new_lines.append(join_indent + '|\t\tПО ВЫРАЗИТЬ(ОстаткиОС.Субконто1 КАК Справочник.ОсновныеСредства) = ПервоначСведенияНУ.ОсновноеСредство\n')
        changes += 1
        print(f'  CHANGE 2: Added 2 LEFT JOINs after line {i+1}')
        i += 1
        continue

    # ====== CHANGE 3: Inner GROUP BY - add 2 fields after СрокПолезногоИспользования ======
    if stripped == '|\t\tПараметрыАморт.СрокПолезногоИспользования,':
        new_lines.append(line)
        indent = line[:len(line) - len(line.lstrip())]
        new_lines.append(indent + 'СчетаУчетаОС.СчетНачисленияАмортизации.Код,\n')
        new_lines.append(indent + 'ПервоначСведенияНУ.НалоговаяГруппаОС.Наименование,\n')
        changes += 1
        print(f'  CHANGE 3: Added 2 fields to inner GROUP BY after line {i+1}')
        i += 1
        continue

    # ====== CHANGE 4: Outer SELECT - add 2 fields after МОЛ_ИдКод ======
    if 'ВЗ.МОЛ_ИдКод КАК МОЛ_ИдКод,' in stripped:
        new_lines.append(line)
        indent = line[:len(line) - len(line.lstrip())]
        new_lines.append(indent + 'ВЗ.СчетАмортизацииКод КАК СчетАмортизацииКод,\n')
        new_lines.append(indent + 'ВЗ.НалоговаяГруппаОСНаим КАК НалоговаяГруппаОСНаим,\n')
        changes += 1
        print(f'  CHANGE 4: Added 2 fields to outer SELECT after line {i+1}')
        i += 1
        continue

    # ====== CHANGE 5: Restore loading code after Подразделение_Наим ======
    if 'НоваяСтрока.Подразделение_Наим = СокрЛП(СтрокаТЗ.Подразделение_Наим)' in stripped:
        new_lines.append(line)
        indent = line[:len(line) - len(line.lstrip())]
        new_lines.append(indent + 'НоваяСтрока.СчетАмортизацииКод = СокрЛП(СтрокаТЗ.СчетАмортизацииКод);\n')
        new_lines.append(indent + 'НоваяСтрока.НалоговаяГруппаОСНаим = СокрЛП(СтрокаТЗ.НалоговаяГруппаОСНаим);\n')
        changes += 1
        print(f'  CHANGE 5: Restored loading code after line {i+1}')
        i += 1
        continue

    # ====== CHANGE 6: Replace hardcoded СчетАмортизации with BuhBud lookup ======
    if 'Рахунок амортизації' in stripped and 'немає поля' in stripped:
        indent = line[:len(line) - len(line.lstrip())]
        # Replace this line and next line with BuhBud lookup block
        new_lines.append(indent + '// Рахунок амортизації — з BuhBud (СчетаБухгалтерскогоУчетаОС) або 131 за замовчуванням\n')
        new_lines.append(indent + 'Если НЕ ПустаяСтрока(СокрЛП(СтрокаОС.СчетАмортизацииКод)) Тогда\n')
        new_lines.append(indent + '\tСчетАмортизацииЕРП = ПланыСчетов.Хозрасчетный.НайтиПоКоду(СокрЛП(СтрокаОС.СчетАмортизацииКод));\n')
        new_lines.append(indent + 'Иначе\n')
        new_lines.append(indent + '\tСчетАмортизацииЕРП = ПланыСчетов.Хозрасчетный.НайтиПоКоду("131");\n')
        new_lines.append(indent + 'КонецЕсли;\n')
        # Skip the next line (the old hardcoded assignment)
        i += 2  # skip comment + assignment
        changes += 1
        print(f'  CHANGE 6: Replaced hardcoded СчетАмортизации with BuhBud lookup at line {i+1}')
        continue

    # ====== CHANGE 7: Add НалоговаяГруппаОС lookup after СчетУчетаДооценокОС block ======
    if 'НоваяСтрокаДок.СчетУчетаДооценокОС = СчетДооценокОСЕРП' in stripped:
        new_lines.append(line)
        # Find the КонецЕсли for this block
        i += 1
        while i < len(lines) and 'КонецЕсли;' not in lines[i].strip():
            new_lines.append(lines[i])
            i += 1
        # Append the КонецЕсли
        if i < len(lines):
            new_lines.append(lines[i])
            indent = lines[i][:len(lines[i]) - len(lines[i].lstrip())]
            # Add empty line and НалоговаяГруппаОС lookup block
            new_lines.append('\n')
            new_lines.append(indent + '// НалоговаяГруппаОС — пошук за найменуванням з BuhBud (ПервоначальныеСведенияОСНалоговыйУчет)\n')
            new_lines.append(indent + 'Если НЕ ПустаяСтрока(СокрЛП(СтрокаОС.НалоговаяГруппаОСНаим)) Тогда\n')
            new_lines.append(indent + '\tНалГруппаОС = Справочники.НалоговыеГруппыОсновныхСредств.НайтиПоНаименованию(\n')
            new_lines.append(indent + '\t\tСокрЛП(СтрокаОС.НалоговаяГруппаОСНаим), Истина);\n')
            new_lines.append(indent + '\tЕсли ЗначениеЗаполнено(НалГруппаОС) Тогда\n')
            new_lines.append(indent + '\t\tНоваяСтрокаДок.НалоговаяГруппаОС = НалГруппаОС;\n')
            new_lines.append(indent + '\tКонецЕсли;\n')
            new_lines.append(indent + 'КонецЕсли;\n')
            changes += 1
            print(f'  CHANGE 7: Added НалоговаяГруппаОС lookup after line {i+1}')
        i += 1
        continue

    new_lines.append(line)
    i += 1

# Write result
with open(filepath, 'w', encoding='utf-8-sig') as f:
    f.writelines(new_lines)

new_count = len(new_lines)
print(f'\nChanges applied: {changes}')
print(f'New file: {new_count} lines (was {original_count}, added {new_count - original_count})')
print('File saved successfully!')
