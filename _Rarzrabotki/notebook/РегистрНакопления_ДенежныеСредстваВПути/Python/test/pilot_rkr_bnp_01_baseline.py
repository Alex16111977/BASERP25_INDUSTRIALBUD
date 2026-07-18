# -*- coding: utf-8 -*-
"""Baseline ДО правки РасчетКурсовых Безнал/Налич Подразделение.

Снимаем:
- Движения РН.Налич по 000Ц-000044 (декабрь): Подр сейчас пуст
- Движения РН.Безнал по 000Ц-000007 (январь): Подр сейчас пуст
- Плуги А_ОтчетБаланс_Свод дек25 «Налич КГ Подгорцы ↔ (пусто) ±8735,60»
- Σ КО total дек+янв (контроль регрессии)
"""
import sys, io, os, json, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
print(f"Org: {S(ORG)}")

ХО_ПДС = erp.Перечисления.ХозяйственныеОперации.ПереоценкаДенежныхСредств
snapshot = {"org": str(S(ORG)), "rkr": {}, "balans_dec": {}, "balans_jan": {}}

# Найти РасчетКурсовых: 000Ц-000044 (декабрь 2025) и 000Ц-000007 (январь 2026)
for ном, year, label in [("000Ц-000044", 2025, "rkr_044_dec25"), ("000Ц-000007", 2026, "rkr_007_jan26")]:
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Н", ном); q.УстановитьПараметр("ХО", ХО_ПДС); q.УстановитьПараметр("Г", year)
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасчетКурсовыхРазниц ГДЕ Номер = &Н И ХозяйственнаяОперация = &ХО И ГОД(Дата) = &Г'
    sel = q.Выполнить().Выбрать()
    if not sel.Следующий():
        print(f"[FAIL] {ном}/{year} не найден"); continue
    DOC = sel.Ссылка
    print(f"\n=== {label}: {S(DOC)} ===")
    rkr_data = {"doc": str(S(DOC)), "movements": {}}

    for reg in ("ДенежныеСредстваБезналичные", "ДенежныеСредстваНаличные"):
        q = erp.NewObject("Запрос"); q.УстановитьПараметр("Д", DOC)
        q.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Д"
        try:
            r = q.Выполнить().Выгрузить()
        except: continue
        rows = []
        for i in range(r.Количество()):
            rec = r.Получить(i)
            podr_v = getattr(rec, "Подразделение", None)
            podr = (S(podr_v) if podr_v and erp.ЗначениеЗаполнено(podr_v) else "(пусто)") if podr_v is not None else "(нет поля)"
            obj_v = getattr(rec, "БанковскийСчет", None) or getattr(rec, "Касса", None)
            obj = S(obj_v) if obj_v and erp.ЗначениеЗаполнено(obj_v) else ""
            rows.append({
                "Период": str(getattr(rec, "Период", "")),
                "ВидДвижения": S(getattr(rec, "ВидДвижения", "")),
                "Объект": obj,
                "Подр": podr,
                "СуммаУпр": float(getattr(rec, "СуммаУпр", 0) or 0),
                "СуммаРегл": float(getattr(rec, "СуммаРегл", 0) or 0),
            })
        if rows:
            print(f"  {reg}: {len(rows)} строк")
            for row in rows[:5]:
                print(f"    {row['Период']:<25} {row['ВидДвижения']:<8} Объект={row['Объект'][:30]:<30} Подр={row['Подр']:<20}  СуммаУпр={row['СуммаУпр']:>14,.2f}")
        rkr_data["movements"][reg] = rows
    snapshot["rkr"][label] = rkr_data

# Плуги А_ОтчетБаланс_Свод декабрь и январь
def сводный(month_num, year, label):
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", ORG)
    q.УстановитьПараметр("М1", datetime.datetime(year, month_num, 1, 0, 0, 0))
    q.УстановитьПараметр("М2", datetime.datetime(year, month_num, 1, 23, 59, 59))
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ФинРез_Баланс ГДЕ Организация = &Орг И Месяц МЕЖДУ &М1 И &М2 И Проведен = ИСТИНА'
    sel = q.Выполнить().Выбрать()
    if not sel.Следующий(): return {}
    DOC = sel.Ссылка
    print(f"\n=== {label}: {S(DOC)} ===")
    # Плуги по Налич/Безнал
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Д", DOC)
    q.Текст = """
    ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Р.Статья) КАК Ст, ПРЕДСТАВЛЕНИЕ(Р.Подразделение) КАК Подр, КОЛИЧЕСТВО(*) КАК К,
        СУММА(Р.СуммаКонечныйОстаток) КАК КО
    ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р
    ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Д И Р.Расхождение = ИСТИНА
    СГРУППИРОВАТЬ ПО Р.Статья, Р.Подразделение
    """
    r = q.Выполнить().Выгрузить()
    plugs = []
    for i in range(r.Количество()):
        rec = r.Получить(i)
        plugs.append({"Статья": str(rec.Ст), "Подр": str(rec.Подр or "(пусто)"), "К": int(rec.К), "КО": float(rec.КО or 0)})
    print(f"  Плугов: {len(plugs)}")
    for p in plugs:
        print(f"    {p['Статья'][:40]:<40} Подр={p['Подр'][:25]:<25} К={p['К']}  КО={p['КО']:>14,.2f}")
    # Σ КО total
    q = erp.NewObject("Запрос"); q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Д", DOC)
    q.Текст = "ВЫБРАТЬ СУММА(Р.СуммаКонечныйОстаток) КАК Σ ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Д"
    rec = q.Выполнить().Выгрузить().Получить(0)
    total = float(rec.Σ or 0)
    print(f"  Σ КО total: {total:>16,.2f}")
    return {"doc": str(S(DOC)), "plugs": plugs, "total_ko": total}

snapshot["balans_dec"] = сводный(12, 2025, "balans_dec25")
snapshot["balans_jan"] = сводный(1, 2026, "balans_jan26")

out = os.path.join(ART, "pilot_rkr_bnp_01_baseline.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(snapshot, f, ensure_ascii=False, indent=2, default=str)
print(f"\n[OK] Snapshot: {out}")
