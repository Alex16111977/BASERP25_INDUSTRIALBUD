# -*- coding: utf-8 -*-
"""PILOT ПКО "Поступление из другой кассы" Шаг 04 — verify N0000023550 ПОСЛЕ доработки.
PASS если:
  - ДенежныеСредстваВПути.Расход.Подразделение = Строительство (= А_ПодразделениеОтправитель), было Глобино-2
  - ПрочиеАктивыПассивы строка статьи "...в пути".Подразделение = Строительство (каскад из ВПути)
  - ПрочиеАктивыПассивы строка "...наличные" и ДенежныеСредстваНаличные = Глобино-2 (НЕ изменились)
  - Σ по каждому регистру не изменилась vs baseline
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администratор";Pwd="24043"'.replace("Администratор","Администратор"))
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
BASELINE = os.path.join(ART, "pilot_pko_priem_01_baseline.json")
before = json.load(open(BASELINE, encoding="utf-8")) if os.path.exists(BASELINE) else {"movements": {}}

DOC_NUM = "N0000023550"
q = erp.NewObject("Запрос")
q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПриходныйКассовыйОрдер ГДЕ Номер = "{DOC_NUM}"'
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка
obj = DOC.ПолучитьОбъект()
exp = str(S(obj.А_ПодразделениеОтправитель)) if erp.ЗначениеЗаполнено(obj.А_ПодразделениеОтправитель) else "(пусто)"
hdr = str(S(obj.Подразделение))
print(f"Документ: {S(DOC)}")
print(f"  Ожидаемое Подразделение ВПути/ПАП-в-пути (= А_ПодразделениеОтправитель): {exp}")
print(f"  Подразделение шапки (для Наличные/ПАП-наличные): {hdr}\n")

REGS = ["ДенежныеСредстваВПути", "ПрочиеАктивыПассивы", "ДенежныеСредстваНаличные", "ДвиженияДенежныхСредств"]
res = {"pass": [], "fail": []}

def stat_name(rec, cols):
    for c in ("Статья",):
        if c in cols:
            v = getattr(rec, c, None)
            try: return str(S(v)) if (v is not None and erp.ЗначениеЗаполнено(v)) else ""
            except: return ""
    return ""

for reg in REGS:
    qq = erp.NewObject("Запрос"); qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try: rr = qq.Выполнить().Выгрузить()
    except Exception as e:
        info = getattr(e, "excepinfo", None); print(f"[{reg}] err: {info[2] if info else e}"); continue
    cols = [c.Имя for c in rr.Колонки]
    sum_now = 0.0
    print(f"--- {reg}: {rr.Количество()} строк ---")
    rows = []
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        vd = ""
        if "ВидДвижения" in cols:
            try: vd = str(S(rec.ВидДвижения))
            except: vd = ""
        summ = float(getattr(rec, "Сумма", 0) or 0) if "Сумма" in cols else 0.0
        sum_now += summ
        podr = ""
        if "Подразделение" in cols:
            v = getattr(rec, "Подразделение", None)
            try: podr = str(S(v)) if (v is not None and erp.ЗначениеЗаполнено(v)) else "(пусто)"
            except: podr = "(пусто)"
        st = stat_name(rec, cols)
        ist = ""
        if "Источник" in cols:
            v = getattr(rec, "Источник", None)
            try: ist = str(S(v)) if (v is not None and erp.ЗначениеЗаполнено(v)) else ""
            except: ist = ""
        rows.append({"vd": vd, "sum": summ, "podr": podr, "st": st, "ist": ist})
        print(f"  {vd:<8} Σ={summ:>12,.2f}  Подр={podr:<16} Статья={st[:32]:<32} Источник={ist[:28]}")

    # Σ-инвариант
    b = before["movements"].get(reg, {"rows": []})
    sum_before = sum(float(r.get("Сумма") or 0) for r in b.get("rows", []))
    if abs(sum_now - sum_before) > 0.01:
        res["fail"].append(f"{reg}: Σ {sum_before:,.2f}→{sum_now:,.2f} (Δ{sum_now-sum_before:+,.2f})")
    else:
        res["pass"].append(f"{reg}: Σ={sum_now:,.2f} без изменений")

    # Целевые проверки
    if reg == "ДенежныеСредстваВПути":
        podrs = {r["podr"] for r in rows}
        if podrs == {exp}:
            res["pass"].append(f"ВПути.Подразделение = {exp} ✅ (= А_ПодразделениеОтправитель)")
        else:
            res["fail"].append(f"ВПути.Подразделение = {sorted(podrs)} ≠ {exp}")
    if reg == "ПрочиеАктивыПассивы":
        vputi = [r for r in rows if "пути" in r["st"].lower() or "пути" in r["ist"].lower()]
        nal = [r for r in rows if "налич" in r["st"].lower() or "налич" in r["ist"].lower()]
        if vputi and all(r["podr"] == exp for r in vputi):
            res["pass"].append(f"ПАП «в пути».Подразделение = {exp} ✅ (каскад из ВПути)")
        elif vputi:
            res["fail"].append(f"ПАП «в пути».Подразделение = {[r['podr'] for r in vputi]} ≠ {exp} → нужна правка ПАП")
        else:
            res["fail"].append("ПАП: строка статьи «в пути» не найдена")
        if nal and all(r["podr"] == hdr for r in nal):
            res["pass"].append(f"ПАП «наличные».Подразделение = {hdr} ✅ (не изменилась)")
        elif nal:
            res["fail"].append(f"ПАП «наличные».Подразделение = {[r['podr'] for r in nal]} ≠ {hdr} (должна была остаться)")
    if reg == "ДенежныеСредстваНаличные":
        podrs = {r["podr"] for r in rows}
        if podrs == {hdr}:
            res["pass"].append(f"Наличные.Подразделение = {hdr} ✅ (не изменилась — деньги пришли в кассу-получатель)")
        else:
            res["fail"].append(f"Наличные.Подразделение = {sorted(podrs)} ≠ {hdr}")
    print()

print("=== РЕЗУЛЬТАТ ===")
for p in res["pass"]: print(f"  ✓ {p}")
for f in res["fail"]: print(f"  ✗ {f}")
print("\n[OVERALL]", "FAIL" if res["fail"] else "PASS — ПКО «Поступление из другой кассы» фикс успешен")
sys.exit(1 if res["fail"] else 0)
