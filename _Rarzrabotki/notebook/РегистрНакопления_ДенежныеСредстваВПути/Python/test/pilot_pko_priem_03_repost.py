# -*- coding: utf-8 -*-
"""PILOT ПКО "Поступление из другой кассы" Шаг 03 — перепровести N0000023550 через COM.
Запускать ТОЛЬКО ПОСЛЕ db-load-xml + db-update.
Safety-check: операция не пишет в РСКПС/РСППС → COM-репост безопасен.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

DOC_NUM = "N0000023550"
q = erp.NewObject("Запрос")
q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПриходныйКассовыйОрдер ГДЕ Номер = "{DOC_NUM}"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print(f"[FAIL] ПКО {DOC_NUM} не найден"); sys.exit(1)
DOC = sel.Ссылка
obj = DOC.ПолучитьОбъект()
print(f"Документ: {S(DOC)}")
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Подразделение(шапка): {S(obj.Подразделение)}")
print(f"  А_ПодразделениеОтправитель: {S(obj.А_ПодразделениеОтправитель) if erp.ЗначениеЗаполнено(obj.А_ПодразделениеОтправитель) else '(пусто)'}")
print(f"  А_ДокРасходныйКассовыйОрдерПередачиДенег: {S(obj.А_ДокРасходныйКассовыйОрдерПередачиДенег) if erp.ЗначениеЗаполнено(obj.А_ДокРасходныйКассовыйОрдерПередачиДенег) else '(пусто)'}")
print(f"  А_ОбработанКазна: {obj.А_ОбработанКазна}")

print("\n  → Safety-check РСКПС/РСППС...")
unsafe = False
for reg in ("РасчетыСКлиентамиПоСрокам", "РасчетыСПоставщикамиПоСрокам", "РасчетыСКлиентами", "РасчетыСПоставщиками"):
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try:
        r = qq.Выполнить().Выбрать(); r.Следующий()
        if r.К and r.К > 0:
            print(f"    [STOP] {reg}: {r.К} движений → только UI"); unsafe = True
    except Exception:
        pass
if unsafe:
    print("\n  [ABORT] есть взаиморасчёты — перепроводить только в UI."); sys.exit(4)
print("    [OK] взаиморасчётов нет — COM безопасен")

print("\n  → Отмена проведения...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        obj = DOC.ПолучитьОбъект()
        print("    [OK] отменено")
except Exception as e:
    info = getattr(e, "excepinfo", None); print(f"    [FAIL] {info[2] if info else e}"); sys.exit(2)

print("\n  → Проведение...")
try:
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print("    [OK] проведено")
except Exception as e:
    info = getattr(e, "excepinfo", None); print(f"    [FAIL] {info[2] if info else e}"); sys.exit(3)

obj = DOC.ПолучитьОбъект()
print(f"\n  Проведен ПОСЛЕ: {obj.Проведен}  А_ОбработанКазна: {obj.А_ОбработанКазна}")
print(f"  А_ПодразделениеОтправитель (после): {S(obj.А_ПодразделениеОтправитель) if erp.ЗначениеЗаполнено(obj.А_ПодразделениеОтправитель) else '(пусто)'}")
print("  ✓ Готово → pilot_pko_priem_04_verify.py")
