# -*- coding: utf-8 -*-
"""Baseline ДО правки Сумма* → СуммаУпр* в Свод_ДенежныеСредства.

Снимаем:
- Σ КО total per орг (декабрь 2025 + январь 2026) для контроля сходимости баланса
- Σ КО по 4 денежным статьям (Безнал/Налич/Подотч/ВПути), отдельно Расхождение=Ложь/Истина
- Количество строк
"""
import sys, io, os, json, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
print(f"Org: {S(ORG)}")

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

def get_doc(month):
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", ORG)
    q.УстановитьПараметр("М1", datetime.datetime(month[0], month[1], 1, 0, 0, 0))
    q.УстановитьПараметр("М2", datetime.datetime(month[0], month[1], 1, 23, 59, 59))
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ФинРез_Баланс ГДЕ Организация = &Орг И Месяц МЕЖДУ &М1 И &М2 И Проведен = ИСТИНА'
    sel = q.Выполнить().Выбрать()
    return sel.Ссылка if sel.Следующий() else None

snapshot = {"org": str(S(ORG)), "months": {}}
for label, m in [("dec25", (2025,12)), ("jan26", (2026,1))]:
    doc = get_doc(m)
    if doc is None:
        print(f"[WARN] {label}: документ не найден"); continue
    print(f"\n=== {label}: {S(doc)} ===")
    month_data = {"doc": str(S(doc)), "statyas": {}, "total": {}}

    for nm, st in статьи.items():
        if st is None: continue
        q = erp.NewObject("Запрос")
        q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Ст", st); q.УстановитьПараметр("Д", doc)
        q.Текст = """
        ВЫБРАТЬ Р.Расхождение, КОЛИЧЕСТВО(*) КАК К,
            СУММА(Р.СуммаНачальныйОстаток) КАК НО,
            СУММА(Р.СуммаКонечныйОстаток) КАК КО
        ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р
        ГДЕ Р.Организация = &Орг И Р.Статья = &Ст И Р.ДокументДвижения = &Д
        СГРУППИРОВАТЬ ПО Р.Расхождение
        """
        r = q.Выполнить().Выгрузить()
        info = []
        for i in range(r.Количество()):
            rec = r.Получить(i)
            info.append({"Расх": bool(rec.Расхождение), "К": int(rec.К), "НО": float(rec.НО or 0), "КО": float(rec.КО or 0)})
        month_data["statyas"][nm] = info
        for row in info:
            print(f"  {nm:<7} Расх={row['Расх']!s:<5} К={row['К']:>4}  НО={row['НО']:>14,.2f}  КО={row['КО']:>14,.2f}")

    # Total per Орг
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", ORG); q.УстановитьПараметр("Д", doc)
    q.Текст = "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, СУММА(Р.СуммаНачальныйОстаток) КАК ΣНО, СУММА(Р.СуммаКонечныйОстаток) КАК ΣКО ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Р ГДЕ Р.Организация = &Орг И Р.ДокументДвижения = &Д"
    rec = q.Выполнить().Выгрузить().Получить(0)
    month_data["total"] = {"К": int(rec.К), "ΣНО": float(rec.ΣНО or 0), "ΣКО": float(rec.ΣКО or 0)}
    print(f"  TOTAL К={int(rec.К)}  Σ НО={float(rec.ΣНО or 0):>16,.2f}  Σ КО={float(rec.ΣКО or 0):>16,.2f}")

    snapshot["months"][label] = month_data

out = os.path.join(ART, "pilot_sumupr_01_baseline.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(snapshot, f, ensure_ascii=False, indent=2, default=str)
print(f"\n[OK] Snapshot: {out}")
