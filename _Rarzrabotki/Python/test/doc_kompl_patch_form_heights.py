# -*- coding: utf-8 -*-
"""РасчетКомплектаций: HeightInTableRows=15 трём большим таблицам —
итоги под таблицами перестают вылезать за экран (растяжение по вертикали
остаётся, на больших экранах таблицы займут свободное место)."""
import re, sys, io

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml"

with io.open(PATH, 'r', encoding='utf-8-sig', newline='') as f:
    text = f.read()

errors = 0

def insert_height(table_name, table_id, first_child):
    global text, errors
    pattern = (r'(<Table name="' + table_name + r'" id="' + table_id + r'">\r\n)([ \t]+)(<' + first_child + r'>)')
    repl = '\\1\\2<HeightInTableRows>15</HeightInTableRows>\r\n\\2\\3'
    new, n = re.subn(pattern, repl, text)
    print(f"{table_name}: замен = {n}")
    if n != 1:
        errors += 1
    else:
        text = new

insert_height("ТаблицаАнализа", "45", "ReadOnly")
insert_height("СписаниеПоНормам", "428", "ReadOnly")
insert_height("СписаниеСверхНормы", "475", "DataPath")

if errors:
    print("ABORT: файл НЕ записан")
    sys.exit(1)

with io.open(PATH, 'w', encoding='utf-8-sig', newline='') as f:
    f.write(text)
print("Файл записан")
