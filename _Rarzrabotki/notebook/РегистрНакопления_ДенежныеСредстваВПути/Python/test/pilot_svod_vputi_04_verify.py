# -*- coding: utf-8 -*-
"""PILOT Свод_ДенежныеСредства Шаг 04 — Verify.

Acceptance:
(a) В А_ОтчетБаланс_Свод после проведения — есть строки Статья=«в пути»,
    Source=ИстВПути, Расхождение=Ложь per (Орг, Подр)
(b) Σ КО per Орг по «в пути» == Σ КО pre-test (133 447,50 для янв26/ТОВ)
(c) Σ |Δ| плугов по «в пути» резко уменьшилась относительно baseline
(d) Σ КО total регистра не сломалась (≈ 0)
(e) Регрессий по Безнал/Налич/Подотч нет
"""
import sys, io, os, json, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
with open(os.path.join(ART, "pilot_svod_vputi_01_baseline.json"), encoding="utf-8") as f:
    baseline = json.load(f)

ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG)
q.УстановитьПараметр("М1", datetime.datetime(2026, 1, 1, 0, 0, 0))
q.УстановитьПараметр("М2", datetime.datetime(2026, 1, 1, 23, 59, 59))
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ФинРез_Баланс ГДЕ Организация = &Орг И Месяц МЕЖДУ &М1 И &М2'
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")

def _ст(nm):
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Н", nm)
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование = &Н'
    sel = q.Выполнить().Выбрать()
    return sel.Ссылка if sel.Следующий() else None

статьи = {
    "ВПути": _ст("Денежные средства в пути"),
    "Безнал": _ст("Денежные средства (безналичные)"),
    "Налич": _ст("Денежные средства (наличные)"),
    "Подотч": _ст("Денежные средства (у подотчетных лиц)"),
}

results = {"pass": [], "fail": [], "warn": []}

print("\n" + "=" * 110)
print(f"{'Статья':<10} {'Расх':<6} {'К baseline':>10} {'К now':>10} {'НО baseline':>16} {'НО now':>16} {'КО baseline':>16} {'КО now':>16}")
print("=" * 110)

for nm, st in статьи.items():
    if st is None: continue
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Ст", st); q.УстановитьПараметр("Док", DOC)
    q.Текст = """
    ВЫБРАТЬ Р.Расхождение КАК Расхождение,
        КОЛИЧЕСТВО(*) КАК К,
        СУММА(Р.СуммаНачальныйОстаток) КАК НО,
        СУММА(Р.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р
    ГДЕ Р.Организация = &Орг И Р.Статья = &Ст И Р.ДокументДвижения = &Док
    СГРУППИРОВАТЬ ПО Р.Расхождение
    """
    r = q.Выполнить().Выгрузить()
    now_data = {}
    for i in range(r.Количество()):
        rec = r.Получить(i)
        now_data[bool(rec.Расхождение)] = {
            "К": int(rec.К), "НО": float(rec.НО or 0), "КО": float(rec.КО or 0)
        }
    base_data = {row["Расхождение"]: row for row in baseline["statyas"][nm]["rows"]}
    for расх in [False, True]:
        bd = base_data.get(расх, {"К": 0, "НО": 0, "КО": 0})
        nd = now_data.get(расх, {"К": 0, "НО": 0, "КО": 0})
        print(f"{nm:<10} {('Истина' if расх else 'Ложь'):<6} {bd['К']:>10} {nd['К']:>10}  {bd['НО']:>16,.2f}  {nd['НО']:>16,.2f}  {bd['КО']:>16,.2f}  {nd['КО']:>16,.2f}")
        # Регрессии для не-ВПути
        if nm != "ВПути":
            if расх == False and abs(bd["КО"] - nd["КО"]) > 0.5:
                results["fail"].append(f"{nm}/Ложь КО регрессия: {bd['КО']:,.2f} → {nd['КО']:,.2f}")

# Acceptance для ВПути
print()
vp_now_false = now_data_v = {bool(r.Получить(i).Расхождение): {"К": int(r.Получить(i).К), "НО": float(r.Получить(i).НО or 0), "КО": float(r.Получить(i).КО or 0)} for i in range(r.Количество())}  # noqa
# Перезапрос ВПути отдельно для чёткости
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Ст", статьи["ВПути"]); q.УстановитьПараметр("Док", DOC)
q.Текст = """
ВЫБРАТЬ Р.Расхождение КАК Расхождение, КОЛИЧЕСТВО(*) КАК К,
    СУММА(Р.СуммаНачальныйОстаток) КАК НО, СУММА(Р.СуммаКонечныйОстаток) КАК КО
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р
ГДЕ Р.Организация = &Орг И Р.Статья = &Ст И Р.ДокументДвижения = &Док
СГРУППИРОВАТЬ ПО Р.Расхождение
"""
r = q.Выполнить().Выгрузить()
vp_now = {bool(r.Получить(i).Расхождение): {"К": int(r.Получить(i).К), "НО": float(r.Получить(i).НО or 0), "КО": float(r.Получить(i).КО or 0)} for i in range(r.Количество())}
base_vp = {row["Расхождение"]: row for row in baseline["statyas"]["ВПути"]["rows"]}

# (a) Есть детали (Расхождение=Ложь)
if vp_now.get(False, {}).get("К", 0) > 0:
    results["pass"].append(f"(a) ВПути: Расхождение=Ложь имеет {vp_now[False]['К']} строк, КО={vp_now[False]['КО']:,.2f}")
else:
    results["fail"].append("(a) ВПути: НЕТ строк Расхождение=Ложь — 4-я ветка не сработала")

# (b) Σ КО (Ложь+Истина) per Орг == 133 447,50 (то что был только в плугах baseline)
expected_ko = base_vp.get(True, {}).get("КО", 0)
actual_total_ko = vp_now.get(False, {}).get("КО", 0) + vp_now.get(True, {}).get("КО", 0)
if abs(actual_total_ko - expected_ko) < 0.5:
    results["pass"].append(f"(b) Σ КО ВПути всего = {actual_total_ko:,.2f} ≈ baseline.плуги.КО {expected_ko:,.2f}")
else:
    results["fail"].append(f"(b) Σ КО ВПути всего {actual_total_ko:,.2f} ≠ baseline.плуги.КО {expected_ko:,.2f} (Δ {actual_total_ko-expected_ko:+,.2f})")

# (c) Плуги уменьшились?
plug_ko_before = base_vp.get(True, {}).get("КО", 0)
plug_ko_after = vp_now.get(True, {}).get("КО", 0)
plug_k_before = base_vp.get(True, {}).get("К", 0)
plug_k_after = vp_now.get(True, {}).get("К", 0)
if abs(plug_ko_after) < 0.5:
    results["pass"].append(f"(c) Плуги ВПути полностью устранены: {plug_k_before}→{plug_k_after} строк, |КО| {plug_ko_before:,.2f}→{plug_ko_after:,.2f}")
elif abs(plug_ko_after) < abs(plug_ko_before):
    results["warn"].append(f"(c) Плуги ВПути уменьшились: {plug_k_before}→{plug_k_after} строк, |КО| {plug_ko_before:,.2f}→{plug_ko_after:,.2f} (PARTIAL)")
else:
    results["fail"].append(f"(c) Плуги НЕ уменьшились: {plug_ko_before:,.2f}→{plug_ko_after:,.2f}")

# (d) Σ КО total регистра ≈ 0
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Док", DOC)
q.Текст = "ВЫБРАТЬ СУММА(Р.СуммаКонечныйОстаток) КАК ΣКО, СУММА(Р.СуммаНачальныйОстаток) КАК ΣНО, КОЛИЧЕСТВО(*) КАК К ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Док"
rec = q.Выполнить().Выгрузить().Получить(0)
total_ko = float(rec.ΣКО or 0); total_no = float(rec.ΣНО or 0); total_k = int(rec.К)
base_total = baseline.get("total", {})
print(f"\nTotal: строк {base_total.get('ВсегоСтрок',0)} → {total_k},  Σ КО {base_total.get('ΣКО',0):,.2f} → {total_ko:,.2f},  Σ НО {base_total.get('ΣНО',0):,.2f} → {total_no:,.2f}")
if abs(total_ko) < 0.5:
    results["pass"].append(f"(d) Σ КО total = {total_ko:,.2f} ≈ 0 (баланс сходится)")
else:
    results["fail"].append(f"(d) Σ КО total = {total_ko:,.2f} ≠ 0 (баланс НЕ сошёлся)")

# (e) уже проверено в цикле для Безнал/Налич/Подотч

print("\n=== РЕЗУЛЬТАТ ===")
if results["pass"]:
    print("\nPASS:"); [print(f"  ✓ {p}") for p in results["pass"]]
if results["warn"]:
    print("\nWARN:"); [print(f"  ! {w}") for w in results["warn"]]
if results["fail"]:
    print("\nFAIL:"); [print(f"  ✗ {f}") for f in results["fail"]]
    print("\n[OVERALL] FAIL"); sys.exit(1)
elif results["pass"]:
    if results["warn"]:
        print("\n[OVERALL] PARTIAL PASS")
    else:
        print("\n[OVERALL] PASS — Свод_ДенежныеСредства расшифровка В пути работает!")
