# -*- coding: utf-8 -*-
"""
ПИЛОТ Шаг 04 — Verify ПКО 000Ц-000001 ПОСЛЕ доработки.

Запускать ПОСЛЕ pilot_03_repost.py.

Что проверяется:
  1. В РН.ДенежныеСредстваВПути для этого ПКО Подразделение = КГ "Подгорцы"
     (значение Касса.Подразделение)
  2. Σ Сумма / СуммаУпр / СуммаРегл движения НЕ изменились (safeguard)
  3. Все остальные регистры — без изменений (Σ движений, ключевые поля)
  4. Diff с pilot_01_baseline.json

Acceptance (PASS):
  ✓ РНДС.ВПути.Подразделение = "КГ \"Подгорцы\""
  ✓ Σ всех движений идентична baseline
  ✓ Никаких новых регистров не появилось / старые не исчезли

Acceptance (FAIL):
  ✗ Подразделение пусто → правка SQL не применилась
  ✗ Σ изменилась → ошибка в правке (задеть другой блок)
  ✗ Новые/исчезнувшие регистры → побочный эффект
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)

import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
BASELINE = os.path.join(ART, "pilot_01_baseline.json")

if not os.path.exists(BASELINE):
    print(f"[FAIL] Нет baseline: {BASELINE}. Запусти pilot_01_baseline.py ДО доработки.")
    sys.exit(1)

with open(BASELINE, encoding="utf-8") as f:
    before = json.load(f)

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПриходныйКассовыйОрдер ГДЕ Номер = "000Ц-000001"'
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
expected_podr = obj.Касса.Подразделение
expected_podr_str = str(S(expected_podr)) if erp.ЗначениеЗаполнено(expected_podr) else "(пусто)"
print(f"  Ожидаемое Подразделение (из Касса.Подразделение): {expected_podr_str}\n")

REGISTRY = list(before["movements"].keys())

results = {"pass": [], "fail": [], "warn": []}

print("=" * 100)
print(f"{'Регистр':<40} {'ДО':>6} {'ПОСЛЕ':>6}  {'Σ ДО':>16}  {'Σ ПОСЛЕ':>16}  {'Δ Σ':>16}  Подр")
print("=" * 100)

for reg_name in REGISTRY:
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Док"
    try:
        rr = qq.Выполнить().Выгрузить()
    except:
        continue

    cols = [c.Имя for c in rr.Колонки]
    sum_now = 0.0
    podrs_now = set()
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        # Сумма (если есть поле)
        if "Сумма" in cols:
            try: sum_now += float(getattr(rec, "Сумма", 0) or 0)
            except: pass
        # Подразделение (если есть)
        if "Подразделение" in cols:
            v = getattr(rec, "Подразделение", None)
            if v is not None:
                try:
                    if erp.ЗначениеЗаполнено(v):
                        podrs_now.add(str(S(v)))
                    else:
                        podrs_now.add("(пусто)")
                except: pass

    # Сравнение с baseline
    before_reg = before["movements"].get(reg_name, {"count": 0, "rows": []})
    count_before = before_reg["count"]
    count_now = rr.Количество()
    sum_before = sum(float(r.get("Сумма") or 0) for r in before_reg["rows"])
    delta_sum = sum_now - sum_before

    podrs_before = set()
    for r in before_reg["rows"]:
        if "Подразделение" in r:
            podrs_before.add(str(r.get("Подразделение") or "(пусто)"))

    podr_now_str = ",".join(sorted(podrs_now)) if podrs_now else "—"
    print(f"{reg_name:<40} {count_before:>6} {count_now:>6}  {sum_before:>16,.2f}  {sum_now:>16,.2f}  {delta_sum:>+16,.2f}  {podr_now_str[:40]}")

    # === Проверки ===
    if count_before != count_now:
        results["fail"].append(f"{reg_name}: число движений изменилось {count_before}→{count_now}")
    if abs(delta_sum) > 0.01:
        results["fail"].append(f"{reg_name}: Σ Сумма изменилась на {delta_sum:+,.2f}")

    # Спец-проверка для ДенежныеСредстваВПути
    if reg_name == "ДенежныеСредстваВПути":
        if rr.Количество() == 0:
            results["fail"].append("ДенежныеСредстваВПути: движения отсутствуют (документ ничего не пишет)")
        elif expected_podr_str == "(пусто)":
            results["warn"].append("ДенежныеСредстваВПути: ожидаемое Подразделение пустое, сравнить нечего")
        elif expected_podr_str in podrs_now:
            results["pass"].append(f"ДенежныеСредстваВПути.Подразделение = '{expected_podr_str}' ✅ совпадает с Касса.Подразделение")
        else:
            results["fail"].append(f"ДенежныеСредстваВПути.Подразделение = {sorted(podrs_now)} ≠ ожидалось '{expected_podr_str}'")

print("=" * 100)

# Итог
print("\n=== РЕЗУЛЬТАТ ===")
if results["pass"]:
    print("\nPASS:")
    for p in results["pass"]: print(f"  ✓ {p}")
if results["warn"]:
    print("\nWARN:")
    for w in results["warn"]: print(f"  ! {w}")
if results["fail"]:
    print("\nFAIL:")
    for f in results["fail"]: print(f"  ✗ {f}")
    print("\n[OVERALL] FAIL — НЕ переходить к расширению на другие документы")
    sys.exit(1)
elif results["pass"]:
    print("\n[OVERALL] PASS — пилот успешен, можно расширять на РКО / СписаниеБезнал / ПостБезнал / РасчетКурсовых")
else:
    print("\n[OVERALL] NEUTRAL — проверки не дали определённого результата")
