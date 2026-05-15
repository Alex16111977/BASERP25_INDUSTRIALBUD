# -*- coding: utf-8 -*-
"""Test 2: проведення А_ФинРез_Баланс Етап 1 v1.2 — 4 ресурси балансуються, idempotent."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom
from datetime import datetime

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
MONTH = datetime(2026, 1, 1, 12, 0, 0)   # січень 2026 (закритий період, Актив=Пасив)
DDATE = datetime(2026, 1, 31, 14, 0, 0)

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(CONN)

# cleanup попередніх документів
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ссылка КАК С ИЗ Документ.А_ФинРез_Баланс"
tz = q.Выполнить().Выгрузить()
for i in range(tz.Количество()):
    tz.Получить(i).С.ПолучитьОбъект().Удалить()

q0 = conn.NewObject("Запрос")
q0.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации ГДЕ КодПоЕДРПОУ = "40645273"'
s = q0.Выполнить().Выбрать(); s.Следующий(); org = s.С

doc = conn.Документы.А_ФинРез_Баланс.СоздатьДокумент()
doc.Дата = DDATE; doc.Организация = org; doc.Месяц = MONTH; doc.ВключатьДочерние = False
mode = conn.PredefinedValue("РежимЗаписиДокумента.Проведение")
doc.Записать(mode)
assert doc.Проведен, "FAIL: документ не проведено"
ref = doc.Ссылка
print(f"[OK] проведено №{doc.Номер}")

q2 = conn.NewObject("Запрос")
q2.Текст = """
ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол,
    СУММА(СуммаНачальныйОстаток) КАК sN, СУММА(СуммаПриход) КАК sP,
    СУММА(СуммаРасход) КАК sR, СУММА(СуммаКонечныйОстаток) КАК sK
ИЗ РегистрСведений.А_ОтчетБаланс_Свод ГДЕ Регистратор = &Д
"""
q2.УстановитьПараметр("Д", ref)
r = q2.Выполнить().Выбрать(); r.Следующий()
total = int(r.Кол)
print(f"рядків={total}, sN={r.sN:,.2f}, sP={r.sP:,.2f}, sR={r.sR:,.2f}, sK={r.sK:,.2f}")
assert total > 0, "FAIL: 0 рядків у регістрі"
assert abs(float(r.sK) - (float(r.sN)+float(r.sP)-float(r.sR))) < 0.01, "FAIL: 4 ресурси не балансуються"

# idempotency — перепровести, перечитати об'єкт заново
obj2 = ref.ПолучитьОбъект(); obj2.Записать(mode)
r2 = q2.Выполнить().Выбрать(); r2.Следующий()
assert int(r2.Кол) == total, f"FAIL: не idempotent {r2.Кол} != {total}"
print(f"[OK] idempotent: {r2.Кол} рядків")
print("PASS Test 2")
