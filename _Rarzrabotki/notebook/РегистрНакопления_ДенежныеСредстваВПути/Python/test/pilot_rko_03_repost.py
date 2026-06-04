# -*- coding: utf-8 -*-
"""PILOT РКО Шаг 03 — Перепровести РКО N0000053020 через COM.
Запускать ТОЛЬКО ПОСЛЕ db-load-xml + db-update.
ХозОп "Конвертация валюты" не пишет в РСКПС/РСППС → COM-репост безопасен.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = "N0000053020"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("[FAIL] РКО N0000053020 не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Подразделение шапки: {S(obj.Подразделение) if erp.ЗначениеЗаполнено(obj.Подразделение) else '(пусто)'}")
print(f"  Касса: {S(obj.Касса)}")

print("\n  → Отменяю проведение...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        obj = DOC.ПолучитьОбъект()
        print("    [OK] Отменено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] Отмена: {info[2] if info else e}")
    sys.exit(2)

print("\n  → Провожу заново...")
try:
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print("    [OK] Проведено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] Проведение: {info[2] if info else e}")
    sys.exit(3)

print(f"\n  ✓ Готово. Запусти pilot_rko_04_verify.py")
