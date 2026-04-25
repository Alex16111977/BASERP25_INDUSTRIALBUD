# -*- coding: utf-8 -*-
# Дивимось структуру регістра ВзаиморасчетыПоНДФЛ у ZUP_1 (БАС ЗУП 2.1)
# і знаходимо поле з сумою доходу.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
zup = v8.Connect('Srvr="SQLSERVER";Ref="zup_1";Usr="cfo";Pwd="2442"')

S = zup.String

# Перелічуємо реквізити і ресурси регістру
print("=" * 70)
print("РегистрНакопления.ВзаиморасчетыПоНДФЛ")
print("=" * 70)
mo = zup.Метаданные.РегистрыНакопления.Найти("ВзаиморасчетыПоНДФЛ")
if mo is None:
    print("НЕ ЗНАЙДЕНО")
    sys.exit(1)

print(f"\nИзмерения ({mo.Измерения.Количество()}):")
for i in range(mo.Измерения.Количество()):
    d = mo.Измерения.Получить(i)
    print(f"  {S(d.Имя)}: {S(d.Тип)[:60]}")

print(f"\nРесурсы ({mo.Ресурсы.Количество()}):")
for i in range(mo.Ресурсы.Количество()):
    d = mo.Ресурсы.Получить(i)
    print(f"  {S(d.Имя)}: {S(d.Тип)[:60]}")

print(f"\nРеквизиты ({mo.Реквизиты.Количество()}):")
for i in range(mo.Реквизиты.Количество()):
    d = mo.Реквизиты.Получить(i)
    print(f"  {S(d.Имя)}: {S(d.Тип)[:60]}")
