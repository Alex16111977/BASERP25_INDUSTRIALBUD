"""Run the external report А_ОтчетPL.erf via COM and save result as .mxl."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp
import datetime

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\А_ОтчетPL.erf"
MXL_OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\А_ОтчетPL_результат.mxl"

conn = connect_erp()
rep = conn.ВнешниеОтчеты.Создать(ERF)
print("Loaded:", rep.Метаданные().Имя)

# Параметры: декабрь 2025
period = conn.NewObject("СтандартныйПериод")
period.ДатаНачала = datetime.datetime(2025, 12, 1, 0, 0, 0)
period.ДатаОкончания = datetime.datetime(2025, 12, 31, 23, 59, 59)

# Прямой доступ через КомпоновщикНастроек
setts = rep.КомпоновщикНастроек.Настройки
params = setts.ПараметрыДанных.Элементы
for i in range(params.Количество()):
    p = params.Получить(i)
    name = str(p.Параметр)
    print(f"  Param: {name}")
    if name == "Период":
        p.Значение = period
        p.Использование = True

tab_doc = conn.NewObject("ТабличныйДокумент")

try:
    rep.Сформировать(tab_doc)
    print("Сформировать: OK")
except Exception as e:
    print("Сформировать error:", e)
    # Пробую напрямую через объект отчёта
    try:
        rep.ПриКомпоновкеРезультата(tab_doc, None, False)
        print("ПриКомпоновкеРезультата: OK")
    except Exception as e2:
        print("Object direct call error:", e2)

tab_doc.Записать(MXL_OUT)
print(f"Saved: {MXL_OUT}")
print(f"Table: {tab_doc.ВысотаТаблицы} rows x {tab_doc.ШиринаТаблицы} cols")
