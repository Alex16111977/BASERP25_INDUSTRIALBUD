# -*- coding: utf-8 -*-
"""
Перевірка що реквізит А_ИсключатьИзОтчетаCashflow доданий до
Справочник.СтатьиДвиженияДенежныхСредств у конфігурації 1С BAS ERP.
"""
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)
S = erp.String

mdo = erp.Метаданные.Справочники.Найти("СтатьиДвиженияДенежныхСредств")
assert mdo is not None, "Справочник.СтатьиДвиженияДенежныхСредств не знайдено у конфігурації"

ATTR_NAME = "А_ИсключатьИзОтчетаCashflow"
found = None
for i in range(mdo.Реквизиты.Количество()):
    rekv = mdo.Реквизиты.Получить(i)
    if S(rekv.Имя) == ATTR_NAME:
        found = rekv
        break

assert found, (
    f"Реквізит {ATTR_NAME} ще НЕ доданий у конфігурацію "
    "— потрібно зробити в 1С Designer"
)
print(f"OK: реквізит знайдено, тип={S(found.Тип)}")
print(f"     Синонім={S(found.Синоним)}")
