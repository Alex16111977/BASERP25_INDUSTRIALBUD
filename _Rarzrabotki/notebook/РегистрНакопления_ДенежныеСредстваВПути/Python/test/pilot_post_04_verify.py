# -*- coding: utf-8 -*-
"""PILOT ПостБезнал Шаг 04 — Verify 00DL-007179."""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
BASELINE = os.path.join(ART, "pilot_post_01_baseline.json")
if not os.path.exists(BASELINE):
    print(f"[FAIL] Нет baseline: {BASELINE}"); sys.exit(1)
with open(BASELINE,encoding="utf-8") as f: before = json.load(f)

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств ГДЕ Номер = "00DL-007179"'
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
expected_podr = obj.Подразделение
expected_podr_str = str(S(expected_podr)) if erp.ЗначениеЗаполнено(expected_podr) else "(пусто)"
print(f"  Ожидаемое Подразделение (из шапки): {expected_podr_str}\n")

REGISTRY = list(before["movements"].keys())
results = {"pass":[], "fail":[], "warn":[]}
print("=" * 100)
print(f"{'Регистр':<40} {'ДО':>5} {'ПОСЛЕ':>5}  {'Σ ДО':>15}  {'Σ ПОСЛЕ':>15}  {'Δ Σ':>15}  Подр")
print("=" * 100)
for reg in REGISTRY:
    qq = erp.NewObject("Запрос"); qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try: rr = qq.Выполнить().Выгрузить()
    except: continue
    cols = [c.Имя for c in rr.Колонки]
    sum_now = 0.0; podrs = set()
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        if "Сумма" in cols:
            try: sum_now += float(getattr(rec,"Сумма",0) or 0)
            except: pass
        if "Подразделение" in cols:
            v = getattr(rec,"Подразделение",None)
            if v is not None:
                try:
                    if erp.ЗначениеЗаполнено(v): podrs.add(str(S(v)))
                    else: podrs.add("(пусто)")
                except: pass
    before_reg = before["movements"].get(reg, {"count":0,"rows":[]})
    cb = before_reg["count"]; cn = rr.Количество()
    sb = sum(float(r.get("Сумма") or 0) for r in before_reg["rows"])
    ds = sum_now - sb
    pstr = ",".join(sorted(podrs)) if podrs else "—"
    print(f"{reg:<40} {cb:>5} {cn:>5}  {sb:>15,.2f}  {sum_now:>15,.2f}  {ds:>+15,.2f}  {pstr[:35]}")
    if cb != cn: results["fail"].append(f"{reg}: число движений {cb}→{cn}")
    if abs(ds) > 0.01: results["fail"].append(f"{reg}: Σ изменилась {ds:+,.2f}")
    if reg == "ДенежныеСредстваВПути":
        if rr.Количество() == 0: results["fail"].append("РНДС.ВПути: 0 движений")
        elif expected_podr_str == "(пусто)": results["warn"].append("Ожидаемое Подр пусто")
        elif expected_podr_str in podrs:
            results["pass"].append(f"РНДС.ВПути.Подразделение = '{expected_podr_str}' ✅ совпадает с шапкой")
        else:
            results["fail"].append(f"РНДС.ВПути.Подр = {sorted(podrs)} ≠ '{expected_podr_str}'")
print("=" * 100); print("\n=== РЕЗУЛЬТАТ ===")
if results["pass"]:
    print("\nPASS:"); [print(f"  ✓ {p}") for p in results["pass"]]
if results["warn"]:
    print("\nWARN:"); [print(f"  ! {w}") for w in results["warn"]]
if results["fail"]:
    print("\nFAIL:"); [print(f"  ✗ {f}") for f in results["fail"]]
    print("\n[OVERALL] FAIL"); sys.exit(1)
elif results["pass"]:
    print("\n[OVERALL] PASS — ПостБезнал пилот успешен")
else:
    print("\n[OVERALL] NEUTRAL")
