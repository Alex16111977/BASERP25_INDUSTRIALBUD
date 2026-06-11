# -*- coding: utf-8 -*-
"""Сверка количества строк: группировка по ФЛ vs по ФЛ+ДРФО (А_ВзСС, дек 2025)."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def count(text):
    q = erp.NewObject("Запрос")
    q.Text = text
    return q.Execute().Выгрузить().Количество()

BASE = """ВЫБРАТЬ
	Ост.ФизическоеЛицо КАК ФЛ{EXTRA_SEL},
	СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК К
ИЗ
	РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , ) КАК Ост
СГРУППИРОВАТЬ ПО
	Ост.ФизическоеЛицо{EXTRA_GRP}"""

n_fl = count(BASE.format(EXTRA_SEL="", EXTRA_GRP=""))
n_fl_drfo = count(BASE.format(
    EXTRA_SEL=",\n\tОст.ФизическоеЛицо.КодПоДРФО КАК ДРФО",
    EXTRA_GRP=",\n\tОст.ФизическоеЛицо.КодПоДРФО"))
print(f"группировка по ФЛ:       {n_fl}")
print(f"группировка по ФЛ+ДРФО:  {n_fl_drfo}")
