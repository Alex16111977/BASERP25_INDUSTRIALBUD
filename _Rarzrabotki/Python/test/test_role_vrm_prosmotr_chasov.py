# -*- coding: utf-8 -*-
"""Smoke: роль А_ВыполнениеРаботПросмотрЧасов в метаданных BaseERP + РольДоступна не падает."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

ROLE = "А_ВыполнениеРаботПросмотрЧасов"

md = erp.Метаданные.Роли.Найти(ROLE)
if md is None:
    print("FAIL: роль не найдена в метаданных")
    sys.exit(1)
print(f"OK: роль в метаданных, синоним = {md.Синоним}")

try:
    dostupna = erp.РольДоступна(ROLE)
    print(f"OK: РольДоступна('{ROLE}') = {dostupna} (без исключения)")
except Exception as e:
    print(f"FAIL: РольДоступна кинула исключение: {e}")
    sys.exit(1)

print("SMOKE PASSED")
