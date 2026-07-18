# -*- coding: utf-8 -*-
"""
ПИЛОТ Шаг 03 — Перепровести ПКО 000Ц-000001 через COM.

Запускать ТОЛЬКО ПОСЛЕ:
  1. pilot_01_baseline.py (зафиксирован snapshot)
  2. pilot_02_sql_pretest.py (SQL подтверждён рабочим)
  3. Edit BSL Documents/ПриходныйКассовыйОрдер/Ext/ManagerModule.bsl (5 блоков)
  4. /db-load-xml -Mode Partial -updateConfigDumpInfo
  5. /db-update -Dynamic+

Этот скрипт:
  1. Отменяет проведение
  2. Проводит заново
  3. Печатает результат

ХозОп ПКО 000Ц-000001 = "Конвертация валюты" — без РСКПС/РСППС, COM-репост безопасен.

Образец: _Rarzrabotki/Python/test/train_repost_single_doc.py.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)

import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПриходныйКассовыйОрдер ГДЕ Номер = "000Ц-000001"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("[FAIL] ПКО 000Ц-000001 не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")

obj = DOC.ПолучитьОбъект()
print(f"  Проведен ДО: {obj.Проведен}")
print(f"  Подразделение шапки: {S(obj.Подразделение) if erp.ЗначениеЗаполнено(obj.Подразделение) else '(пусто)'}")
print(f"  Касса: {S(obj.Касса)}")
print(f"  Касса.Подразделение (целевое для движения): "
      f"{S(obj.Касса.Подразделение) if erp.ЗначениеЗаполнено(obj.Касса.Подразделение) else '(пусто)'}")

# --- Отмена + проведение ---
print("\n  → Отменяю проведение...")
try:
    if obj.Проведен:
        obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        obj = DOC.ПолучитьОбъект()  # перечитать после отмены
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
    msg = info[2] if info else str(e)
    print(f"    [FAIL] Проведение: {msg}")
    print("\n  Возможные причины:")
    print("    - 'Поле не найдено Подразделение' → /db-update не выполнялся или регистр не обновлён")
    print("    - 'Несоответствие типов' → ошибка в SQL блока (см. pilot_02)")
    print("    - 'Не задано значение параметра ...' → MCP HTTP context, нужен Designer/Enterprise")
    sys.exit(3)

print(f"\n  ✓ Готово. Запусти pilot_04_verify.py для сравнения ДО/ПОСЛЕ.")
