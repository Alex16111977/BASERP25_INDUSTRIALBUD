# -*- coding: utf-8 -*-
"""Test проведения А_ФинРез_Баланс (Task 12 финализирует: rows>0 + идемпотентно).
На стадии каркаса (Task 4): достаточно Проведен=Истина (регистр пуст — ОК).
Параметр ОЖИДАТЬ_СТРОКИ управляет жёсткостью (False до Task 5+, True с Task 5)."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom
from datetime import datetime
pythoncom.CoInitialize()
ERP = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

ОЖИДАТЬ_СТРОКИ = os.environ.get("BALANS_EXPECT_ROWS", "0") == "1"

q0 = ERP.NewObject("Запрос")
q0.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации '
            'ГДЕ КодПоЕДРПОУ = "40645273"')
s = q0.Выполнить().Выбрать(); s.Следующий(); ORG = s.С

q = ERP.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ссылка КАК С ИЗ Документ.А_ФинРез_Баланс"
tz = q.Выполнить().Выгрузить()
for i in range(tz.Количество()):
    tz.Получить(i).С.ПолучитьОбъект().Удалить()

doc = ERP.Документы.А_ФинРез_Баланс.СоздатьДокумент()
doc.Дата = datetime(2026, 1, 31, 14, 0, 0)
doc.Организация = ORG
doc.Месяц = datetime(2026, 1, 1, 12, 0, 0)
doc.ВключатьДочерние = False
mode = ERP.PredefinedValue("РежимЗаписиДокумента.Проведение")
try:
    doc.Записать(mode)
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"FAIL: проведение упало: {msg}")
    sys.exit(1)
assert doc.Проведен, "FAIL: не проведено"
ref = doc.Ссылка

qr = ERP.NewObject("Запрос")
qr.Текст = ("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, СУММА(СуммаКонечныйОстаток) КАК S "
            "ИЗ РегистрСведений.А_ОтчетБаланс_Свод ГДЕ Регистратор=&Д")
qr.УстановитьПараметр("Д", ref)
r = qr.Выполнить().Выбрать(); r.Следующий()
К = int(r.К)
print(f"[OK] проведено №{doc.Номер}, строк={К}, ΣКонОст={float(r.S or 0):,.2f}")

if ОЖИДАТЬ_СТРОКИ:
    assert К > 0, "FAIL: регистр пуст (ожидались строки)"
    doc.Записать(mode)
    qr2 = ERP.NewObject("Запрос"); qr2.Текст = qr.Текст
    qr2.УстановитьПараметр("Д", ref)
    r2 = qr2.Выполнить().Выбрать(); r2.Следующий()
    assert int(r2.К) == К, f"FAIL: не идемпотентно {К}->{int(r2.К)}"
    print("PASS test_balans_m_post (Проведен, строк>0, идемпотентно)")
else:
    print("PASS test_balans_m_post (каркас: Проведен; регистр пуст — ОК)")
