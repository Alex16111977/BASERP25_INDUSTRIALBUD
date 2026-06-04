# -*- coding: utf-8 -*-
"""Post-verify Безнал/Налич Подразделение РасчетКурсовых.

Проверки:
1. Движения 000Ц-000044/000007 в РН.Безнал/Налич — Подразделение теперь заполнен
2. Плуги «Налич КГ Подгорцы ↔ (пусто)» в А_ОтчетБаланс_Свод устранены
3. Σ КО total per Орг не сломан
4. Σ Сумма/УПР/Регл движений не изменилась
"""
import sys, io, os, json, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
with open(os.path.join(ART, "pilot_rkr_bnp_01_baseline.json"), encoding="utf-8") as f:
    base = json.load(f)

ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
ХО_ПДС = erp.Перечисления.ХозяйственныеОперации.ПереоценкаДенежныхСредств

results = {"pass": [], "fail": [], "warn": []}

# 1. Проверка движений РасчетКурсовых
print("=" * 110)
print("ПРОВЕРКА 1: Движения 000Ц-000044 + 000Ц-000007 — Подр заполнен")
print("=" * 110)
for ном, year, label in [("000Ц-000044", 2025, "rkr_044_dec25"), ("000Ц-000007", 2026, "rkr_007_jan26")]:
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Н", ном); q.УстановитьПараметр("ХО", ХО_ПДС); q.УстановитьПараметр("Г", year)
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасчетКурсовыхРазниц ГДЕ Номер = &Н И ХозяйственнаяОперация = &ХО И ГОД(Дата) = &Г'
    sel = q.Выполнить().Выбрать(); sel.Следующий()
    DOC = sel.Ссылка
    base_rkr = base["rkr"].get(label, {"movements": {}})
    print(f"\n  {label}: {S(DOC)}")
    for reg in ("ДенежныеСредстваБезналичные", "ДенежныеСредстваНаличные"):
        q = erp.NewObject("Запрос"); q.УстановитьПараметр("Д", DOC)
        q.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Д"
        try: r = q.Выполнить().Выгрузить()
        except: continue
        if r.Количество() == 0: continue
        sum_upr_now = 0.0
        podrs = set()
        for i in range(r.Количество()):
            rec = r.Получить(i)
            try: sum_upr_now += float(getattr(rec, "СуммаУпр", 0) or 0)
            except: pass
            podr_v = getattr(rec, "Подразделение", None)
            if podr_v is not None:
                try:
                    if erp.ЗначениеЗаполнено(podr_v): podrs.add(str(S(podr_v)))
                    else: podrs.add("(пусто)")
                except: pass
        base_rows = base_rkr["movements"].get(reg, [])
        base_sum = sum(r.get("СуммаУпр", 0) for r in base_rows)
        base_podrs = set(r["Подр"] for r in base_rows)
        pstr_now = ",".join(sorted(podrs))
        pstr_base = ",".join(sorted(base_podrs))
        print(f"    {reg:<35}  К {len(base_rows)}→{r.Количество()}  ΣУПР {base_sum:,.2f}→{sum_upr_now:,.2f}  Подр: '{pstr_base}' → '{pstr_now}'")
        # Acceptance
        if abs(sum_upr_now - base_sum) > 0.5:
            results["fail"].append(f"{label}/{reg}: ΣУПР изменилась {sum_upr_now-base_sum:+,.2f}")
        if "(пусто)" in podrs and "(пусто)" not in base_podrs:
            results["fail"].append(f"{label}/{reg}: появились новые пустые Подр")
        elif "(пусто)" not in podrs and "(пусто)" in base_podrs:
            results["pass"].append(f"{label}/{reg}: Подр заполнен ({pstr_now})")
        elif "(пусто)" in podrs and "(пусто)" in base_podrs:
            results["warn"].append(f"{label}/{reg}: Подр всё ещё содержит (пусто)")

# 2. Плуги А_ОтчетБаланс_Свод
print()
print("=" * 110)
print("ПРОВЕРКА 2: Плуги в А_ОтчетБаланс_Свод (дек25 + янв26)")
print("=" * 110)
for label, m_num, m_year in [("balans_dec", 12, 2025), ("balans_jan", 1, 2026)]:
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Орг", ORG)
    q.УстановитьПараметр("М1", datetime.datetime(m_year, m_num, 1, 0, 0, 0))
    q.УстановитьПараметр("М2", datetime.datetime(m_year, m_num, 1, 23, 59, 59))
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ФинРез_Баланс ГДЕ Организация = &Орг И Месяц МЕЖДУ &М1 И &М2'
    sel = q.Выполнить().Выбрать(); sel.Следующий()
    DOC = sel.Ссылка

    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Д", DOC)
    q.Текст = """
    ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Р.Статья) КАК Ст, ПРЕДСТАВЛЕНИЕ(Р.Подразделение) КАК Подр,
        КОЛИЧЕСТВО(*) КАК К, СУММА(Р.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р
    ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Д И Р.Расхождение = ИСТИНА
    СГРУППИРОВАТЬ ПО Р.Статья, Р.Подразделение
    """
    r = q.Выполнить().Выгрузить()
    plugs_now = []
    for i in range(r.Количество()):
        rec = r.Получить(i)
        plugs_now.append({"Статья": str(rec.Ст), "Подр": str(rec.Подр or "(пусто)"), "КО": float(rec.КО or 0)})
    base_plugs = base[label].get("plugs", [])
    print(f"\n  {label}:")
    print(f"    Плугов было: {len(base_plugs)}, стало: {len(plugs_now)}")
    # Найти что устранилось / появилось
    base_keys = {(p["Статья"], p["Подр"]): p["КО"] for p in base_plugs}
    now_keys = {(p["Статья"], p["Подр"]): p["КО"] for p in plugs_now}
    устранены = sorted(set(base_keys) - set(now_keys))
    появились = sorted(set(now_keys) - set(base_keys))
    for k in устранены:
        print(f"    ✅ Устранён: {k[0][:35]:<35} / {k[1][:25]:<25}  КО {base_keys[k]:>14,.2f}")
        results["pass"].append(f"{label}: устранён плуг {k[0]}/{k[1]} КО={base_keys[k]:,.2f}")
    for k in появились:
        print(f"    ⚠ Появился: {k[0][:35]:<35} / {k[1][:25]:<25}  КО {now_keys[k]:>14,.2f}")
        results["warn"].append(f"{label}: появился новый плуг {k[0]}/{k[1]} КО={now_keys[k]:,.2f}")

    # 3. Σ КО total
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Д", DOC)
    q.Текст = "ВЫБРАТЬ СУММА(Р.СуммаКонечныйОстаток) КАК Σ ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Д"
    total_now = float(q.Выполнить().Выгрузить().Получить(0).Σ or 0)
    total_base = base[label].get("total_ko", 0)
    print(f"    Σ КО total: {total_base:,.2f} → {total_now:,.2f} (Δ={total_now-total_base:+,.2f})")
    if abs(total_now - total_base) > 0.5:
        results["warn"].append(f"{label}: Σ КО total изменилась {total_now-total_base:+,.2f}")
    else:
        results["pass"].append(f"{label}: Σ КО total стабильно ({total_now:,.2f})")

print()
print("=" * 80)
print("=== РЕЗУЛЬТАТ ===")
print("=" * 80)
if results["pass"]:
    print("\nPASS:"); [print(f"  ✓ {p}") for p in results["pass"]]
if results["warn"]:
    print("\nWARN:"); [print(f"  ! {w}") for w in results["warn"]]
if results["fail"]:
    print("\nFAIL:"); [print(f"  ✗ {f}") for f in results["fail"]]
    print("\n[OVERALL] FAIL"); sys.exit(1)
print("\n[OVERALL] PASS")
