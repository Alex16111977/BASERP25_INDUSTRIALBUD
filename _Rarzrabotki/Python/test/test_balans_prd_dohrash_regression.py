# -*- coding: utf-8 -*-
"""Regression — перепровести декабрь 2025 и январь 2026 после внедрения расшифровки.

Главный инвариант: Σ свод per (Орг,Подр,Статья) == ПАП.ОстаткиИОбороты(КонМес) до 0.01.
Доп: Σ signed КО (OD-3) = 0.00 ⇒ Актив=|Пассив|.

Примечание: абсолютный |Актив| как сумма положительных КО ИЗМЕНИТСЯ после
добавления детализации по СтатьеДоходыРасходы (был эталон 278 093 267,32 для дек2025
основан на свёрнутой структуре per Статья; теперь детали разделены — |Актив| вырос).
Это by-design, баланс per Орг сохраняется на уровне Σ signed.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client
from datetime import datetime

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)
S = erp.String
TOL = 0.01

MONTHS = {
    "2025-12": ("декабрь 2025", datetime(2025, 12, 1), datetime(2025, 12, 31, 23, 59, 59)),
    "2026-01": ("январь 2026", datetime(2026, 1, 1), datetime(2026, 1, 31, 23, 59, 59)),
}

q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Ссылка ИЗ Справочник.Организации ГДЕ КодПоЕДРПОУ = "40645273" """
own_org = q.Execute().Выгрузить().Получить(0).Ссылка

СтСС = erp.ПланыВидовХарактеристик.СтатьиАктивовПассивов.НайтиПоНаименованию("Собственные средства")
iskl_arr = erp.NewObject("Массив")
iskl_arr.Добавить(СтСС)

failures = []

for key, (name, nm, km) in MONTHS.items():
    print(f"\n=== {name} ===")
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", own_org)
    q.УстановитьПараметр("НМ", nm)
    q.УстановитьПараметр("КМ", km)
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ Документ.А_ФинРез_Баланс
    ГДЕ Организация = &Орг И НЕ ПометкаУдаления
        И Месяц >= &НМ И Месяц <= &КМ
    """
    r = q.Execute().Выгрузить()
    if r.Количество() == 0:
        print(f"  [WARN] Документ за {name} не найден — пропуск")
        continue
    doc_ref = r.Получить(0).Ссылка
    print(f"  Документ: {S(doc_ref)}")

    try:
        doc_obj = doc_ref.ПолучитьОбъект()
        doc_obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        print(f"  [OK] Перепроведение PASS")
    except Exception as e:
        failures.append((name, f"Перепроведение FAIL: {e}"))
        continue

    # Σ signed КО (OD-3) = 0 ⇒ Актив=|Пассив|
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Рег", doc_ref)
    q.УстановитьПараметр("Искл", iskl_arr)
    q.Text = """
    ВЫБРАТЬ СУММА(РС.СуммаКонечныйОстаток) КАК ΣКО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК РС
    ГДЕ РС.ДокументДвижения = &Рег
        И НЕ РС.Статья В ИЕРАРХИИ(&Искл)
    """
    sigma_ko = float(q.Execute().Выгрузить().Получить(0).ΣКО or 0)
    print(f"  Σ signed КО (OD-3) = {sigma_ko:.2f}")
    if abs(sigma_ko) > 1.0:
        failures.append((name, f"Σ КО = {sigma_ko:.2f} (Актив != |Пассив|)"))
        continue

    # Σ свод per (Орг,Подр,Статья) over Source == ПАП.signedΣ per same до 0.01
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Рег", doc_ref)
    q.УстановитьПараметр("Искл", iskl_arr)
    q.Text = """
    ВЫБРАТЬ
        РС.Подразделение, РС.Статья,
        СУММА(РС.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК РС
    ГДЕ РС.ДокументДвижения = &Рег И НЕ РС.Статья В ИЕРАРХИИ(&Искл)
    СГРУППИРОВАТЬ ПО РС.Подразделение, РС.Статья
    """
    svod = q.Execute().Выгрузить()
    svod_map = {}
    for i in range(svod.Количество()):
        row = svod.Получить(i)
        svod_map[(S(row.Подразделение), S(row.Статья))] = float(row.КО)

    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", own_org)
    q.УстановитьПараметр("КМ", km)
    q.УстановитьПараметр("Искл", iskl_arr)
    q.Text = """
    ВЫБРАТЬ
        Б.Подразделение, Б.Статья,
        СУММА(Б.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(, &КМ, Авто, ,
        Организация = &Орг И НЕ Статья В ИЕРАРХИИ(&Искл)) КАК Б
    СГРУППИРОВАТЬ ПО Б.Подразделение, Б.Статья
    """
    pap = q.Execute().Выгрузить()
    pap_map = {}
    for i in range(pap.Количество()):
        row = pap.Получить(i)
        pap_map[(S(row.Подразделение), S(row.Статья))] = float(row.КО)

    keys = set(svod_map.keys()) | set(pap_map.keys())
    mismatches = []
    for k in keys:
        sv = svod_map.get(k, 0.0)
        pp = pap_map.get(k, 0.0)
        if abs(sv - pp) > TOL:
            mismatches.append((k, sv, pp))

    if mismatches:
        print(f"  [FAIL] Σ свод != ПАП ({len(mismatches)} расхождений):")
        for k, sv, pp in mismatches[:5]:
            print(f"    {k} | свод={sv:.2f} | ПАП={pp:.2f} | Δ={sv-pp:.2f}")
        failures.append((name, f"Σ свод != ПАП per (Подр,Статья) ({len(mismatches)} расхождений)"))
    else:
        print(f"  [OK] Σ свод per (Подр,Статья) == ПАП до {TOL} ({len(keys)} ключей)")

if failures:
    print(f"\n[FAIL] {len(failures)} ошибок:")
    for name, msg in failures:
        print(f"  {name}: {msg}")
    sys.exit(1)

print(f"\n[OK] Regression PASS — Σ свод == ПАП.ОстаткиИОбороты, баланс per Орг сохранён")
