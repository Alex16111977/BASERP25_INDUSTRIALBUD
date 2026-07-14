# -*- coding: utf-8 -*-
"""Smoke пересобранных .epf/.erf: обработка/отчёт создаётся в контексте BaseERP
(= модуль объекта компилируется) и видит общий модуль А_ПодключенияБазСервер.

БезопасныйРежим = Ложь обязателен: обработки создают COMОбъект("V83.COMConnector").
"""
import os, sys
import win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ROOT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki"

PROCESSORS = [
    r"Обработки\А_НачальнаяЗадолженностьПоЗарплатеФорма2.epf",
    r"Обработки\Перенос остатков исполнительные листы и социальные фонды из бухгалтерии.epf",
    r"Обработки\Перенос остатков налогов из бухгалтерии.epf",
    r"Обработки\Перенос остатков сотрудников.epf",
    r"Обработки\СинхронизироватьВзаиморасчеты.epf",
    r"Обработки\СинхронизироватьДеньги.epf",
    r"Обработки\СинхронизироватьДеньгиКасса.epf",
    r"Обработки\СинхронизироватьТоварыТолькоТовары.epf",
]
REPORTS = [
    r"Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсBASБухгалтерия.erf",
    r"Отчеты\А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсЗУП.erf",
    r"Отчеты\А_СравнитьОстаткиНалоговЕРПсBASБухгалтерия.erf",
    r"Отчеты\А_СравнитьОстаткиНалоговПосчетамЕРПсBASБухгалтерия.erf",
]

def err_text(e):
    return (e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)).strip()[:180]

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
fails = 0

for rel in PROCESSORS:
    path = os.path.join(ROOT, rel)
    try:
        obj = erp.ВнешниеОбработки.Создать(path, False)  # БезопасныйРежим = Ложь
        obj = None
        print(f"OK  обработка {os.path.basename(rel)}")
    except Exception as e:
        print(f"FAIL обработка {os.path.basename(rel)}: {err_text(e)}")
        fails += 1

for rel in REPORTS:
    path = os.path.join(ROOT, rel)
    try:
        obj = erp.ВнешниеОтчеты.Создать(path, False)
        obj = None
        print(f"OK  отчёт {os.path.basename(rel)}")
    except Exception as e:
        print(f"FAIL отчёт {os.path.basename(rel)}: {err_text(e)}")
        fails += 1

print("ИТОГ:", f"OK ({len(PROCESSORS) + len(REPORTS)}/{len(PROCESSORS) + len(REPORTS)})" if fails == 0 else f"FAIL ({fails})")
