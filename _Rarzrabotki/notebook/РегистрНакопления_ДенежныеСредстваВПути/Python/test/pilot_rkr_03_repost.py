# -*- coding: utf-8 -*-
"""PILOT РасчетКурсовых Шаг 03 — Перепровести 000Ц-000007 от 31.01.2026.

ХозОп "Переоценка денежных средств" — без РСКПС/РСППС → COM безопасен.
Особенность: документ переоценивает существующие остатки. Подразделение тянется
из ОСТАТКОВ (новое измерение), а не из шапки.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ХО = erp.Перечисления.ХозяйственныеОперации.ПереоценкаДенежныхСредств
q = erp.NewObject("Запрос")
q.УстановитьПараметр("ХО", ХО)
q.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасчетКурсовыхРазниц '
           'ГДЕ Номер = "000Ц-000007" И ХозяйственнаяОперация = &ХО И ГОД(Дата) = 2026')
sel = q.Выполнить().Выбрать()
if not sel.Следующий(): print("[FAIL] не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Дата:        {obj.Дата}")
print(f"  Организация: {S(obj.Организация)}")
print(f"  ХозОп:       {S(obj.ХозяйственнаяОперация)}")

# Safety check — РасчетКурсовых.ПереоценкаДенежныхСредств не должен трогать РСКПС/РСППС
qq = erp.NewObject("Запрос")
qq.УстановитьПараметр("Док", DOC)
for reg in ("РасчетыСКлиентамиПоСрокам", "РасчетыСПоставщикамиПоСрокам"):
    qq.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} ГДЕ Регистратор = &Док"
    rr = qq.Выполнить().Выгрузить()
    if int(rr.Получить(0).К) > 0:
        print(f"\n[STOP] Есть {rr.Получить(0).К} движений в {reg} — COM ОПАСЕН")
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

print("\n  ✓ Запусти pilot_rkr_04_verify.py")
