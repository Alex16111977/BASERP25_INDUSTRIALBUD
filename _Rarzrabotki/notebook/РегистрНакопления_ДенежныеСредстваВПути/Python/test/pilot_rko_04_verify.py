# -*- coding: utf-8 -*-
"""PILOT РКО Шаг 04 — Verify РКО N0000053020 ПОСЛЕ доработки.
Diff с pilot_rko_01_baseline.json. PASS если Подразделение в РНДС.ВПути = КГ "Подгорцы" + Σ не изменилась.
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
BASELINE = os.path.join(ART, "pilot_rko_01_baseline.json")
if not os.path.exists(BASELINE):
    print(f"[FAIL] Нет baseline: {BASELINE}. Запусти pilot_rko_01_baseline.py ДО доработки.")
    sys.exit(1)
with open(BASELINE, encoding="utf-8") as f: before = json.load(f)

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = "N0000053020"'
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
expected_podr = obj.Подразделение  # из шапки документа (как делает РНДС.Наличные)
expected_podr_str = str(S(expected_podr)) if erp.ЗначениеЗаполнено(expected_podr) else "(пусто)"
print(f"  Ожидаемое Подразделение (из шапки документа): {expected_podr_str}\n")

REGISTRY = list(before["movements"].keys())
results = {"pass": [], "fail": [], "warn": []}

print("=" * 100)
print(f"{'Регистр':<40} {'ДО':>6} {'ПОСЛЕ':>6}  {'Σ ДО':>16}  {'Σ ПОСЛЕ':>16}  {'Δ Σ':>16}  Подр")
print("=" * 100)

for reg_name in REGISTRY:
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Док"
    try: rr = qq.Выполнить().Выгрузить()
    except: continue
    cols = [c.Имя for c in rr.Колонки]
    sum_now = 0.0; podrs_now = set()
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        if "Сумма" in cols:
            try: sum_now += float(getattr(rec, "Сумма", 0) or 0)
            except: pass
        if "Подразделение" in cols:
            v = getattr(rec, "Подразделение", None)
            if v is not None:
                try:
                    if erp.ЗначениеЗаполнено(v): podrs_now.add(str(S(v)))
                    else: podrs_now.add("(пусто)")
                except: pass
    before_reg = before["movements"].get(reg_name, {"count": 0, "rows": []})
    count_before = before_reg["count"]; count_now = rr.Количество()
    sum_before = sum(float(r.get("Сумма") or 0) for r in before_reg["rows"])
    delta_sum = sum_now - sum_before
    podr_now_str = ",".join(sorted(podrs_now)) if podrs_now else "—"
    print(f"{reg_name:<40} {count_before:>6} {count_now:>6}  {sum_before:>16,.2f}  {sum_now:>16,.2f}  {delta_sum:>+16,.2f}  {podr_now_str[:40]}")

    if count_before != count_now:
        results["fail"].append(f"{reg_name}: число движений {count_before}→{count_now}")
    if abs(delta_sum) > 0.01:
        results["fail"].append(f"{reg_name}: Σ Сумма изменилась на {delta_sum:+,.2f}")
    if reg_name == "ДенежныеСредстваВПути":
        if rr.Количество() == 0:
            results["fail"].append("ДенежныеСредстваВПути: движения отсутствуют")
        elif expected_podr_str == "(пусто)":
            results["warn"].append("ДенежныеСредстваВПути: ожидаемое Подразделение пусто — нечего сравнить")
        elif expected_podr_str in podrs_now:
            results["pass"].append(f"ДенежныеСредстваВПути.Подразделение = '{expected_podr_str}' ✅ совпадает с шапкой документа")
        else:
            results["fail"].append(f"ДенежныеСредстваВПути.Подразделение = {sorted(podrs_now)} ≠ ожидалось '{expected_podr_str}'")

print("=" * 100)
print("\n=== РЕЗУЛЬТАТ ===")
if results["pass"]:
    print("\nPASS:"); [print(f"  ✓ {p}") for p in results["pass"]]
if results["warn"]:
    print("\nWARN:"); [print(f"  ! {w}") for w in results["warn"]]
if results["fail"]:
    print("\nFAIL:"); [print(f"  ✗ {f}") for f in results["fail"]]
    print("\n[OVERALL] FAIL"); sys.exit(1)
elif results["pass"]:
    print("\n[OVERALL] PASS — РКО пилот успешен")
else:
    print("\n[OVERALL] NEUTRAL")
