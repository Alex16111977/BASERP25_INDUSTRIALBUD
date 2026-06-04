# -*- coding: utf-8 -*-
"""PILOT Свод_ДенежныеСредства Шаг 03 — Repost А_ФинРез_Баланс за январь 2026."""
import sys, io, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG)
q.УстановитьПараметр("М1", datetime.datetime(2026, 1, 1, 0, 0, 0))
q.УстановитьПараметр("М2", datetime.datetime(2026, 1, 1, 23, 59, 59))
q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ФинРез_Баланс
ГДЕ Организация = &Орг И Месяц МЕЖДУ &М1 И &М2 И Проведен = ИСТИНА
"""
sel = q.Выполнить().Выбрать()
if not sel.Следующий(): print("[FAIL]"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Месяц: {obj.Месяц}")

print("\n  → Отменяю проведение...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        obj = DOC.ПолучитьОбъект()
        print("    [OK] Отменено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] {info[2] if info else e}"); sys.exit(2)

print("  → Провожу заново...")
try:
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print("    [OK] Проведено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] {info[2] if info else e}"); sys.exit(3)

print("\n  ✓ Запусти pilot_svod_vputi_04_verify.py")
