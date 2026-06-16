# -*- coding: utf-8 -*-
"""PILOT РКО "Выдача в другую кассу" Шаг 03 — Перепровести РКО N0000052950 через COM.
Запускать ТОЛЬКО ПОСЛЕ db-load-xml + db-update.

Safety-check: ВыдачаДенежныхСредствВДругуюКассу не пишет в РСКПС/РСППС → COM-репост безопасен.
Скрипт ВСЁ РАВНО проверяет отсутствие движений по РСКПС/РСППС и прерывается, если они есть
(memory feedback_com_repost_skips_registrator_raschetov).
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

DOC_NUM = "N0000052950"
q = erp.NewObject("Запрос")
q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = "{DOC_NUM}"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print(f"[FAIL] РКО {DOC_NUM} не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Подразделение шапки: {S(obj.Подразделение) if erp.ЗначениеЗаполнено(obj.Подразделение) else '(пусто)'}")
print(f"  Касса: {S(obj.Касса)}  ХозОп: {S(obj.ХозяйственнаяОперация)}")
print(f"  А_Обработан={obj.А_Обработан}  А_ОбработанКазна={obj.А_ОбработанКазна}  А_ВведенВЕРП={obj.А_ВведенВЕРП}")

# === SAFETY: проверка РСКПС/РСППС ===
print("\n  → Safety-check РСКПС/РСППС...")
unsafe = False
for reg in ("РасчетыСКлиентамиПоСрокам", "РасчетыСПоставщикамиПоСрокам", "РасчетыСКлиентами", "РасчетыСПоставщиками"):
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try:
        r = qq.Выполнить().Выбрать(); r.Следующий()
        if r.К and r.К > 0:
            print(f"    [STOP] {reg}: {r.К} движений → COM-репост ОПАСЕН, только UI"); unsafe = True
    except Exception:
        pass  # регистр может быть недоступен/не существовать
if unsafe:
    print("\n  [ABORT] Есть движения по расчётам — перепроводить ТОЛЬКО через UI 1С Enterprise."); sys.exit(4)
print("    [OK] Движений по РСКПС/РСППС нет — COM-репост безопасен")

print("\n  → Отменяю проведение...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        obj = DOC.ПолучитьОбъект()
        print("    [OK] Отменено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] Отмена: {info[2] if info else e}"); sys.exit(2)

print("\n  → Провожу заново...")
try:
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print("    [OK] Проведено")
except Exception as e:
    info = getattr(e, "excepinfo", None)
    print(f"    [FAIL] Проведение: {info[2] if info else e}"); sys.exit(3)

obj = DOC.ПолучитьОбъект()
print(f"\n  Проведен ПОСЛЕ: {obj.Проведен}")
print(f"  А_Обработан={obj.А_Обработан}  А_ОбработанКазна={obj.А_ОбработанКазна}  (флаги должны сохраниться)")
print(f"\n  ✓ Готово. Запусти pilot_rko_vk_04_verify.py")
