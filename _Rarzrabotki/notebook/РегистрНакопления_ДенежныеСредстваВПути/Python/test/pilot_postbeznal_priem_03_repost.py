# -*- coding: utf-8 -*-
"""ПостБезнал «с другого счёта» Шаг 03 — перепровести 00DL-006964 (COM, safety-check РСКПС/РСППС)."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = lambda v: erp.XMLСтрока(v)
DOC_NUM = "00DL-006964"
q = erp.NewObject("Запрос")
q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств ГДЕ Номер = "{DOC_NUM}"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий(): print(f"[FAIL] {DOC_NUM} не найден"); sys.exit(1)
DOC = sel.Ссылка
obj = DOC.ПолучитьОбъект()
print(f"Документ: {DOC_NUM}  Проведен ДО: {obj.Проведен}")
print(f"  Подразделение(шапка): {obj.Подразделение.Наименование}")
print(f"  А_ПодразделениеОтправитель: {obj.А_ПодразделениеОтправитель.Наименование if erp.ЗначениеЗаполнено(obj.А_ПодразделениеОтправитель) else '(пусто)'}")

print("\n  → Safety-check РСКПС/РСППС...")
unsafe = False
for reg in ("РасчетыСКлиентамиПоСрокам","РасчетыСПоставщикамиПоСрокам","РасчетыСКлиентами","РасчетыСПоставщиками"):
    qq = erp.NewObject("Запрос"); qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try:
        r = qq.Выполнить().Выбрать(); r.Следующий()
        if r.К and r.К>0: print(f"    [STOP] {reg}: {r.К} → только UI"); unsafe=True
    except: pass
if unsafe: print("\n  [ABORT] есть взаиморасчёты — только UI"); sys.exit(4)
print("    [OK] взаиморасчётов нет — COM безопасен")

print("\n  → Отмена + проведение...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения); obj = DOC.ПолучитьОбъект()
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print("    [OK] перепроведено")
except Exception as e:
    info = getattr(e,"excepinfo",None); print(f"    [FAIL] {info[2] if info else e}"); sys.exit(3)
print("  ✓ → pilot_postbeznal_priem_04_verify.py")
