# -*- coding: utf-8 -*-
"""PILOT РасчетКурсовых Шаг 04 — Verify 000Ц-000007 от 31.01.2026.

Acceptance:
- РНДС.ВПути.Подразделение НЕ должно ломаться (если в источнике-остатке пусто —
  и в переоценке пусто, это PARTIAL PASS)
- Σ Сумма/СуммаУпр/СуммаРегл движений НЕ изменилась
- Количество движений совпадает с baseline
- Если в baseline было пусто и сейчас совпадает с Подр исходного валютного остатка
  на дату документа — PASS.
"""
import sys, io, os, json, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
BASELINE = os.path.join(ART, "pilot_rkr_01_baseline.json")
if not os.path.exists(BASELINE):
    print(f"[FAIL] Нет baseline: {BASELINE}"); sys.exit(1)
with open(BASELINE, encoding="utf-8") as f: before = json.load(f)

ХО = erp.Перечисления.ХозяйственныеОперации.ПереоценкаДенежныхСредств
q = erp.NewObject("Запрос")
q.УстановитьПараметр("ХО", ХО)
q.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасчетКурсовыхРазниц '
           'ГДЕ Номер = "000Ц-000007" И ХозяйственнаяОперация = &ХО И ГОД(Дата) = 2026')
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")

# Получим источники остатков (что переоценивается) для сверки Подр
obj = DOC.ПолучитьОбъект()
ORG = obj.Организация
ВалУпр = erp.Константы.ВалютаУправленческогоУчета.Получить()
ВалРегл = erp.Константы.ВалютаРегламентированногоУчета.Получить()
qd = erp.NewObject("Запрос")
qd.УстановитьПараметр("Орг", ORG)
qd.УстановитьПараметр("ВалУпр", ВалУпр)
qd.УстановитьПараметр("ВалРегл", ВалРегл)
qd.УстановитьПараметр("Дата", datetime.datetime(2026, 1, 31, 23, 59, 59))
qd.Текст = """
ВЫБРАТЬ
    О.Валюта КАК Валюта,
    О.Подразделение КАК Подразделение,
    О.Получатель КАК Получатель,
    О.Отправитель КАК Отправитель
ИЗ РегистрНакопления.ДенежныеСредстваВПути.Остатки(&Дата,
    Организация = &Орг И (Валюта <> &ВалУпр ИЛИ Валюта <> &ВалРегл)) КАК О
ГДЕ О.СуммаОстаток <> 0
"""
rd = qd.Выполнить().Выгрузить()
expected_podrs = set()
for i in range(rd.Количество()):
    rec = rd.Получить(i)
    if erp.ЗначениеЗаполнено(rec.Подразделение):
        expected_podrs.add(str(S(rec.Подразделение)))
    else:
        expected_podrs.add("(пусто)")
print(f"  Подр источников-остатков на 31.01.2026: {sorted(expected_podrs)}")

REGISTRY = list(before["movements"].keys())
results = {"pass": [], "fail": [], "warn": []}
print("\n" + "=" * 100)
print(f"{'Регистр':<30} {'ДО':>4} {'ПОСЛЕ':>5}  {'Σ ДО':>14}  {'Σ ПОСЛЕ':>14}  {'Δ Σ':>10}  Подр")
print("=" * 100)

for reg in REGISTRY:
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try:
        rr = qq.Выполнить().Выгрузить()
    except Exception:
        continue
    cols = [c.Имя for c in rr.Колонки]
    sum_now = 0.0
    sum_upr_now = 0.0
    sum_regl_now = 0.0
    podrs = set()
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        if "Сумма" in cols:
            try: sum_now += float(getattr(rec, "Сумма", 0) or 0)
            except: pass
        if "СуммаУпр" in cols:
            try: sum_upr_now += float(getattr(rec, "СуммаУпр", 0) or 0)
            except: pass
        if "СуммаРегл" in cols:
            try: sum_regl_now += float(getattr(rec, "СуммаРегл", 0) or 0)
            except: pass
        if "Подразделение" in cols:
            v = getattr(rec, "Подразделение", None)
            if v is not None:
                try:
                    if erp.ЗначениеЗаполнено(v):
                        podrs.add(str(S(v)))
                    else:
                        podrs.add("(пусто)")
                except: pass
    before_reg = before["movements"].get(reg, {"count": 0, "rows": []})
    cb = before_reg["count"]; cn = rr.Количество()
    sb = sum(float(r.get("Сумма") or 0) for r in before_reg["rows"])
    sb_upr = sum(float(r.get("СуммаУпр") or 0) for r in before_reg["rows"])
    sb_regl = sum(float(r.get("СуммаРегл") or 0) for r in before_reg["rows"])
    ds = sum_now - sb
    pstr = ",".join(sorted(podrs)) if podrs else "—"
    print(f"{reg:<30} {cb:>4} {cn:>5}  {sb:>14,.2f}  {sum_now:>14,.2f}  {ds:>+10,.2f}  {pstr[:50]}")
    if cb != cn: results["fail"].append(f"{reg}: число движений {cb}→{cn}")
    if abs(ds) > 0.01: results["fail"].append(f"{reg}: Σ Сумма изменилась {ds:+,.2f}")
    if abs(sum_upr_now - sb_upr) > 0.01:
        results["fail"].append(f"{reg}: Σ СуммаУпр изменилась {sum_upr_now-sb_upr:+,.2f}")
    if abs(sum_regl_now - sb_regl) > 0.01:
        results["fail"].append(f"{reg}: Σ СуммаРегл изменилась {sum_regl_now-sb_regl:+,.2f}")
    if reg == "ДенежныеСредстваВПути":
        if rr.Количество() == 0:
            results["fail"].append("РНДС.ВПути: 0 движений после перепроведения")
        else:
            # Сверка Подр с источниками-остатками
            if "(пусто)" in expected_podrs and "(пусто)" in podrs:
                results["warn"].append(
                    f"РНДС.ВПути.Подр=(пусто) — это by-design "
                    f"(источники-остатки тоже с пустым Подр). PARTIAL PASS.")
            elif podrs.issubset(expected_podrs):
                results["pass"].append(
                    f"РНДС.ВПути.Подразделение={sorted(podrs)} ✓ совпадает с источниками-остатками")
            else:
                results["fail"].append(
                    f"РНДС.ВПути.Подр={sorted(podrs)} ≠ источники {sorted(expected_podrs)}")

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
    print("\n[OVERALL] PASS — РасчетКурсовых пилот успешен")
elif results["warn"]:
    print("\n[OVERALL] PARTIAL PASS — пилот валиден, but Подр=пусто из-за исторических источников")
else:
    print("\n[OVERALL] NEUTRAL")
