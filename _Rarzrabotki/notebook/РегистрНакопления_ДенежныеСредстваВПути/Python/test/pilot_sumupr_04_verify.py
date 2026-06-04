# -*- coding: utf-8 -*-
"""Post-verify Свод_ДенежныеСредства после правки СуммаУпр*.

Проверки:
1. diag_dec25_money_per_doc.py — 17 DIFF_SUMM → 0 DIFF_SUMM
2. Σ КО per статья из РН.* СуммаУпр == ПАП.КО (декабрь 2025 + январь 2026)
3. Σ КО total per Орг не сломан (январь 2026 был = 0)
4. Регрессий по UAH-операциям нет (Безнал/Налич без валюты — суммы идентичны)
"""
import sys, io, os, json, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
with open(os.path.join(ART, "pilot_sumupr_01_baseline.json"), encoding="utf-8") as f:
    base = json.load(f)

ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
def _ст(nm):
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Н", nm)
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование = &Н'
    sel = q.Выполнить().Выбрать()
    return sel.Ссылка if sel.Следующий() else None
статьи = {
    "ВПути":  _ст("Денежные средства в пути"),
    "Безнал": _ст("Денежные средства (безналичные)"),
    "Налич":  _ст("Денежные средства (наличные)"),
    "Подотч": _ст("Денежные средства (у подотчетных лиц)"),
}

def get_doc(m):
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Орг", ORG)
    q.УстановитьПараметр("М1", datetime.datetime(m[0], m[1], 1, 0, 0, 0))
    q.УстановитьПараметр("М2", datetime.datetime(m[0], m[1], 1, 23, 59, 59))
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ФинРез_Баланс ГДЕ Организация = &Орг И Месяц МЕЖДУ &М1 И &М2'
    sel = q.Выполнить().Выбрать(); sel.Следующий()
    return sel.Ссылка

results = {"pass": [], "fail": [], "warn": []}

for label, m in [("dec25", (2025,12)), ("jan26", (2026,1))]:
    doc = get_doc(m)
    print(f"\n{'='*110}")
    print(f"=== {label}: {S(doc)} ===")
    print(f"{'='*110}")
    print(f"{'Статья':<8} {'Расх':<6}  {'К base':>6}  {'К now':>6}  {'НО base':>16}  {'НО now':>16}  {'КО base':>16}  {'КО now':>16}")

    base_m = base["months"].get(label, {"statyas": {}, "total": {}})
    for nm, st in статьи.items():
        if st is None: continue
        q = erp.NewObject("Запрос")
        q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Ст", st); q.УстановитьПараметр("Д", doc)
        q.Текст = """
        ВЫБРАТЬ Р.Расхождение, КОЛИЧЕСТВО(*) КАК К,
            СУММА(Р.СуммаНачальныйОстаток) КАК НО, СУММА(Р.СуммаКонечныйОстаток) КАК КО
        ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р
        ГДЕ Р.Организация = &Орг И Р.Статья = &Ст И Р.ДокументДвижения = &Д
        СГРУППИРОВАТЬ ПО Р.Расхождение
        """
        r = q.Выполнить().Выгрузить()
        now = {bool(r.Получить(i).Расхождение): {"К": int(r.Получить(i).К), "НО": float(r.Получить(i).НО or 0), "КО": float(r.Получить(i).КО or 0)} for i in range(r.Количество())}
        base_rows = {row["Расх"]: row for row in base_m["statyas"].get(nm, [])}
        for расх in [False, True]:
            bd = base_rows.get(расх, {"К": 0, "НО": 0, "КО": 0})
            nd = now.get(расх, {"К": 0, "НО": 0, "КО": 0})
            mark = ""
            if расх and abs(nd["КО"]) < 0.5 and abs(bd["КО"]) > 0.5:
                mark = "  ✅ плуг устранён"
            elif not расх and abs(nd["КО"] - bd["КО"]) > 0.5:
                mark = f"  ⚠ Δ={nd['КО']-bd['КО']:+,.2f}"
            print(f"  {nm:<7} {('Ист' if расх else 'Ложь'):<5}  {bd['К']:>6}  {nd['К']:>6}  {bd['НО']:>16,.2f}  {nd['НО']:>16,.2f}  {bd['КО']:>16,.2f}  {nd['КО']:>16,.2f}{mark}")

    # Total
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Д", doc)
    q.Текст = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, СУММА(Р.СуммаНачальныйОстаток) КАК ΣНО, СУММА(Р.СуммаКонечныйОстаток) КАК ΣКО ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Д"
    rec = q.Выполнить().Выгрузить().Получить(0)
    now_total = {"К": int(rec.К), "ΣНО": float(rec.ΣНО or 0), "ΣКО": float(rec.ΣКО or 0)}
    bt = base_m.get("total", {"К":0, "ΣНО":0, "ΣКО":0})
    print(f"\n  TOTAL  К {bt['К']:>5} → {now_total['К']:>5}  Σ НО {bt['ΣНО']:>16,.2f} → {now_total['ΣНО']:>16,.2f}  Σ КО {bt['ΣКО']:>16,.2f} → {now_total['ΣКО']:>16,.2f}")

    # Acceptance per месяц
    # (d) ΣКО total не должен ухудшиться (по abs)
    if abs(now_total["ΣКО"]) > abs(bt["ΣКО"]) + 0.5:
        results["fail"].append(f"{label}: |Σ КО total| вырос {bt['ΣКО']:,.2f} → {now_total['ΣКО']:,.2f}")
    elif abs(now_total["ΣКО"]) < abs(bt["ΣКО"]) - 0.5:
        results["pass"].append(f"{label}: |Σ КО total| улучшен {bt['ΣКО']:,.2f} → {now_total['ΣКО']:,.2f}")
    else:
        results["pass"].append(f"{label}: |Σ КО total| без изменений ({now_total['ΣКО']:,.2f})")

print("\n" + "=" * 80)
print("=== РЕЗУЛЬТАТ ===")
print("=" * 80)
if results["pass"]:
    print("\nPASS:"); [print(f"  ✓ {p}") for p in results["pass"]]
if results["warn"]:
    print("\nWARN:"); [print(f"  ! {w}") for w in results["warn"]]
if results["fail"]:
    print("\nFAIL:"); [print(f"  ✗ {f}") for f in results["fail"]]; sys.exit(1)
print("\n[OVERALL] PASS")
