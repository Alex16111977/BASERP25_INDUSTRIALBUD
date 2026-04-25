"""Diag: выручка Астарта декабрь."""
import sys, datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

c = connect_erp()
podr = c.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Астарта. Тищенки", True)
q = c.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
	СУММА(ТБ.СуммаВыручкиРеглОборот) КАК ВыручкаРегл,
	СУММА(ТБ.СтоимостьРеглБезНДСОборот) КАК СебестРегл
ИЗ РегистрНакопления.ВыручкаИСебестоимостьПродаж.Обороты(&С, &ПО, , Подразделение = &П) КАК ТБ
"""
q.УстановитьПараметр("С", datetime.datetime(2025, 12, 1))
q.УстановитьПараметр("ПО", datetime.datetime(2025, 12, 31, 23, 59, 59))
q.УстановитьПараметр("П", podr)
tz = q.Выполнить().Выгрузить()
r = tz.Получить(0)
vyr = 0.0 if r.ВыручкаРегл is None else float(r.ВыручкаРегл)
seb = 0.0 if r.СебестРегл is None else float(r.СебестРегл)
print(f"Выручка регл: {vyr:.2f}")
print(f"Себест регл:  {seb:.2f}")
