# -*- coding: utf-8 -*-
"""Проверка функции СведенияОВнешнейОбработке() в скомпилированной .epf."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\ЗакрытиеОтрицательныхОстатков.epf"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

обработка = erp.ВнешниеОбработки.Создать(EPF)
рез = обработка.СведенияОВнешнейОбработке()

print("=" * 60)
print("СведенияОВнешнейОбработке:")
print(f"  Вид:             {erp.String(рез.Вид)}")
print(f"  Версия:          {erp.String(рез.Версия)}")
print(f"  БезопасныйРежим: {рез.БезопасныйРежим}")
print(f"  Наименование:    {erp.String(рез.Наименование)}")
print(f"  Информация:      {erp.String(рез.Информация)}")
print(f"  Команды (кол-во): {рез.Команды.Количество()}")
for i in range(рез.Команды.Количество()):
    к = рез.Команды.Получить(i)
    print(f"    [{i}] Представление: {erp.String(к.Представление)}")
    print(f"        Идентификатор:   {erp.String(к.Идентификатор)}")
    print(f"        Использование:   {erp.String(к.Использование)}")
    print(f"        ПоказыватьОповещение: {к.ПоказыватьОповещение}")
print("=" * 60)
print("OK")
