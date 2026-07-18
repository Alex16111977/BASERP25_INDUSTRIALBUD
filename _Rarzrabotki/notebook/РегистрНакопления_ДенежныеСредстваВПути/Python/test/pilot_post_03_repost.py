# -*- coding: utf-8 -*-
"""PILOT ПостБезнал Шаг 03 — Перепровести 00DL-007179.
ХозОп "Поступление ДС с другого счета" — межбанк, без РСКПС/РСППС → COM безопасен.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств ГДЕ Номер = "00DL-007179"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий(): print("[FAIL]"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Подразделение шапки: {S(obj.Подразделение) if erp.ЗначениеЗаполнено(obj.Подразделение) else '(пусто)'}")
print(f"  БС: {S(obj.БанковскийСчет)}")
print(f"  ХозОп: {S(obj.ХозяйственнаяОперация)}")

# Safety check
q.УстановитьПараметр("Док", DOC)
for reg in ("РасчетыСКлиентамиПоСрокам", "РасчетыСПоставщикамиПоСрокам"):
    q.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} ГДЕ Регистратор = &Док"
    rr = q.Выполнить().Выгрузить()
    if int(rr.Получить(0).К) > 0:
        print(f"\n[STOP] Есть {rr.Получить(0).К} движений в {reg} — COM ОПАСЕН (используй UI)")
        sys.exit(2)

print("\n  → Отменяю проведение...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        obj = DOC.ПолучитьОбъект()
        print("    [OK] Отменено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] {info[2] if info else e}"); sys.exit(3)
print("  → Провожу заново...")
try:
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print("    [OK] Проведено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] {info[2] if info else e}"); sys.exit(4)
print("\n  ✓ Запусти pilot_post_04_verify.py")
