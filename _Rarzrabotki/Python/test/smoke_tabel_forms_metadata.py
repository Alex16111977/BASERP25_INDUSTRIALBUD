# -*- coding: utf-8 -*-
"""Smoke: после загрузки форм А_ТабельУчетаРабочегоВремени база живая, формы в метаданных.

Контекст: доработка ФормаДокумента/ФормаСписка (единое окно БСП, 2026-06-10).
COM не открывает управляемые формы — проверяем доступность базы и состав форм,
визуальная проверка меню выполняется пользователем в 1С Enterprise.
"""
import sys

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("OK: соединение с BaseERP установлено")

md = erp.Метаданные.Документы.А_ТабельУчетаРабочегоВремени
print(f"Документ: {erp.String(md.Имя)}")

forms = md.Формы
names = [erp.String(forms.Получить(i).Имя) for i in range(forms.Количество())]
print(f"Форм: {len(names)}: {', '.join(names)}")

expected = {"ФормаДокумента", "ФормаСписка", "ФормаВводаОборудования"}
missing = expected - set(names)
if missing:
    print(f"FAIL: отсутствуют формы: {missing}")
    sys.exit(1)

print("OK: все ожидаемые формы на месте")
