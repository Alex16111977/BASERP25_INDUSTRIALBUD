# -*- coding: utf-8 -*-
"""Тест зовнішнього звіту А_ОтчетPL_v2.erf через COM (русские методы через getattr)."""
import sys, io, datetime
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\А_ОтчетPL_v2.erf"
MXL = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\А_ОтчетPL_v2_янв2026.mxl"

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
)
print("[OK] Connected")

R = lambda obj, name: getattr(obj, name)

ВнешниеОтчеты = R(erp, "ВнешниеОтчеты")
otchet = R(ВнешниеОтчеты, "Создать")(ERF)
print("[OK] ERF loaded")

nastr = R(R(otchet, "КомпоновщикНастроек"), "Настройки")
PCD = erp.NewObject("ПараметрКомпоновкиДанных", "Период")
param = R(R(nastr, "Параметры"), "НайтиЗначениеПараметра")(PCD)
if param is None:
    print("[ERR] Параметр Период не найден")
    sys.exit(1)

sp = erp.NewObject("СтандартныйПериод")
setattr(sp, "ДатаНачала",  datetime.datetime(2026, 1, 1))
setattr(sp, "ДатаОкончания", datetime.datetime(2026, 1, 31, 23, 59, 59))
setattr(param, "Значение", sp)
setattr(param, "Использование", True)
print("[OK] Период = 2026-01-01 .. 2026-01-31")

result = erp.NewObject("ТабличныйДокумент")
try:
    R(otchet, "КомпоноватьРезультат")(result)
except Exception as e:
    print(f"[ERR formирование]: {e}")
    sys.exit(2)

h = R(result, "ВысотаТаблицы")
w = R(result, "ШиринаТаблицы")
print(f"[OK] Сформирован: высота={h}, ширина={w}")

# Сохранение
TYPE_MXL = 4  # ТипФайлаТабличногоДокумента.MXL
R(result, "Записать")(MXL, TYPE_MXL)
print(f"[OK] Сохранено: {MXL}")

# Проверка наличия колонок
расх, прих = False, False
for row in range(1, min(15, h + 1)):
    for col in range(1, min(60, w + 1)):
        try:
            cell = R(result, "Область")(row, col, row, col)
            t = str(R(cell, "Текст") or "")
            if "Касса (расход)" in t: расх = True; print(f"  '{t}' @ ({row},{col})")
            if "Касса (приход)" in t: прих = True; print(f"  '{t}' @ ({row},{col})")
        except: pass

print(f"\n[РЕЗУЛЬТАТ] Колонки в шапке: расход={расх}, приход={прих}")
if not (расх and прих):
    print("[WARN] Колонки в шапке не найдены — нужно добавить их в варианты СКД")
else:
    print("[OK] Колонки на месте, отчёт работает")
