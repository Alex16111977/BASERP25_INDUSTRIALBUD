"""ФинРезультаты Астарта."""
import sys, datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

c = connect_erp()
podr = c.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Астарта. Тищенки", True)
q = c.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
	СУММА(ТБ.ДоходыРеглБезНДСОборот) КАК ДохРегл,
	СУММА(ТБ.РасходыРеглБезНДСОборот) КАК РасРегл
ИЗ РегистрНакопления.ФинансовыеРезультаты.Обороты(&С, &ПО, , Подразделение = &П) КАК ТБ
"""
q.УстановитьПараметр("С", datetime.datetime(2025, 12, 1))
q.УстановитьПараметр("ПО", datetime.datetime(2025, 12, 31, 23, 59, 59))
q.УстановитьПараметр("П", podr)
r = q.Выполнить().Выгрузить().Получить(0)
for f in ["ДохРегл", "РасРегл"]:
    v = getattr(r, f)
    print(f"{f}: {0 if v is None else float(v):.2f}")
