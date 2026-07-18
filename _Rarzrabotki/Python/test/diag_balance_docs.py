# -*- coding: utf-8 -*-
"""Сколько А_ФинРез_Баланс документов у нас? Каждый покрывает свой месяц?"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Ссылка ИЗ Справочник.Организации ГДЕ КодПоЕДРПОУ = "40645273" """
own_org = q.Execute().Выгрузить().Получить(0).Ссылка

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", own_org)
q.Text = """
ВЫБРАТЬ Номер, Дата, Месяц, Проведен, ПометкаУдаления
ИЗ Документ.А_ФинРез_Баланс
ГДЕ Организация = &Орг
УПОРЯДОЧИТЬ ПО Месяц
"""
r = q.Execute().Выгрузить()
print(f"Всего А_ФинРез_Баланс для ТОВ ИНДАСТРИАЛБУД: {r.Количество()}")
for i in range(r.Количество()):
    row = r.Получить(i)
    print(f"  Номер={row.Номер} | Дата={row.Дата} | Месяц={row.Месяц} | Проведен={row.Проведен} | ПометкаУдаления={row.ПометкаУдаления}")
