# -*- coding: utf-8 -*-
"""Check BuhBud register structures for СчетаБухгалтерскогоУчетаОС and ПервоначальныеСведенияОСНалоговыйУчет"""
import win32com.client
import pythoncom

pythoncom.CoInitialize()
conn_str = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'
v8 = win32com.client.Dispatch('V83.COMConnector')
v83 = v8.Connect(conn_str)

regs = v83.Metadata.InformationRegisters

with open('reg_structure2.txt', 'w', encoding='utf-8') as f:
    for reg_name in ['СчетаБухгалтерскогоУчетаОС', 'ПервоначальныеСведенияОСНалоговыйУчет']:
        r = regs.Find(reg_name)
        if r is None:
            f.write(f'Register {reg_name}: NOT FOUND!\n\n')
            continue

        f.write(f'=== {r.Name} ===\n')
        f.write(f'  InformationRegisterPeriodicity: {r.InformationRegisterPeriodicity}\n')

        f.write('  Dimensions:\n')
        for i in range(r.Dimensions.Count()):
            d = r.Dimensions.Get(i)
            f.write(f'    {d.Name}\n')

        f.write('  Resources:\n')
        for i in range(r.Resources.Count()):
            res = r.Resources.Get(i)
            f.write(f'    {res.Name}\n')

        f.write('  Attributes:\n')
        if r.Attributes.Count() == 0:
            f.write('    (none)\n')
        for i in range(r.Attributes.Count()):
            a = r.Attributes.Get(i)
            f.write(f'    {a.Name}\n')
        f.write('\n')

v83 = None
v8 = None
pythoncom.CoUninitialize()
print('Done - check reg_structure2.txt')
