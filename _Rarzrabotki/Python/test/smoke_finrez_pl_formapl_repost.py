# -*- coding: utf-8 -*-
"""
Smoke (bas-verify): перепровести 1 боевой А_ФинРез_PL за декабрь 2025 через COM
(новый код уже в БД после /db-update -Dynamic+; свежее COM-соединение видит правку)
и убедиться, что движения А_ОтчетPL_Свод теперь несут ФормаPL Ф1/Ф2, Σ неизменна.

Перепроведение А_ФинРез_PL через COM надёжно (это НЕ ОЗФУ и НЕ РСППС-взаиморасчёты).
"""
import sys, io
import win32com.client, pythoncom

if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

pythoncom.CoInitialize()
ERP = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = ERP.String


def snapshot(reg):
    q = ERP.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ
	ЕСТЬNULL(СУММА(Р.Сумма),0) КАК Сумма,
	ЕСТЬNULL(СУММА(Р.СуммаФ1),0) КАК СуммаФ1,
	ЕСТЬNULL(СУММА(Р.СуммаФ2),0) КАК СуммаФ2,
	ЕСТЬNULL(СУММА(Р.Сумма_Excel),0) КАК Сумма_Excel,
	ЕСТЬNULL(СУММА(Р.СуммаФ1_Excel),0) КАК СуммаФ1_Excel,
	ЕСТЬNULL(СУММА(Р.СуммаФ2_Excel),0) КАК СуммаФ2_Excel,
	КОЛИЧЕСТВО(*) КАК Строк
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р
ГДЕ Р.Регистратор = &Рег
"""
    q.УстановитьПараметр("Рег", reg)
    r = q.Выполнить().Выгрузить()[0]
    return {f: float(getattr(r, f)) for f in
            ["Сумма", "СуммаФ1", "СуммаФ2", "Сумма_Excel", "СуммаФ1_Excel", "СуммаФ2_Excel"]} | {"Строк": int(r.Строк)}


def forma_dist(reg):
    q = ERP.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ
	Р.ФормаPL КАК Форма,
	КОЛИЧЕСТВО(*) КАК Строк,
	ЕСТЬNULL(СУММА(Р.Сумма),0) КАК Сумма,
	ЕСТЬNULL(СУММА(Р.СуммаФ1),0) КАК Ф1,
	ЕСТЬNULL(СУММА(Р.СуммаФ2),0) КАК Ф2
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р
ГДЕ Р.Регистратор = &Рег
СГРУППИРОВАТЬ ПО Р.ФормаPL
"""
    q.УстановитьПараметр("Рег", reg)
    out = []
    r = q.Выполнить().Выгрузить()
    for i in range(r.Количество()):
        row = r[i]
        f = S(row.Форма)
        out.append((f if f else "<пусто>", int(row.Строк), float(row.Сумма), float(row.Ф1), float(row.Ф2)))
    return out


def main():
    # найти декабрь 2025
    q = ERP.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 ФР.Ссылка КАК Ссылка, ФР.Номер КАК Номер, ФР.Дата КАК Дата
ИЗ Документ.А_ФинРез_PL КАК ФР
ГДЕ ФР.Проведен И НЕ ФР.ПометкаУдаления И ФР.Месяц = ДАТАВРЕМЯ(2025,12,1)
"""
    r = q.Выполнить().Выгрузить()
    if r.Количество() == 0:
        print("!! не найден проведённый А_ФинРез_PL за декабрь 2025")
        sys.exit(1)
    doc = r[0].Ссылка
    print(f"Документ: №{r[0].Номер} от {r[0].Дата}  ({S(doc)})")

    before = snapshot(doc)
    print("\n=== BEFORE re-post ===")
    print(f"  строк={before['Строк']}  Σ Сумма={before['Сумма']:,.2f}  Ф1={before['СуммаФ1']:,.2f}  Ф2={before['СуммаФ2']:,.2f}")
    print("  ФормаPL распределение:")
    for f, n, s, f1, f2 in forma_dist(doc):
        print(f"    {f:<12} строк={n:<5} Σ={s:>16,.2f}")

    # перепровести
    print("\n=== RE-POST (COM) ===")
    obj = doc.ПолучитьОбъект()
    obj.Записать(ERP.РежимЗаписиДокумента.Проведение, ERP.РежимПроведенияДокумента.Неоперативный)
    print("  записан с проведением OK")

    after = snapshot(doc)
    print("\n=== AFTER re-post ===")
    print(f"  строк={after['Строк']}  Σ Сумма={after['Сумма']:,.2f}  Ф1={after['СуммаФ1']:,.2f}  Ф2={after['СуммаФ2']:,.2f}")
    print("  ФормаPL распределение:")
    dist = forma_dist(doc)
    for f, n, s, f1, f2 in dist:
        print(f"    {f:<12} строк={n:<5} Σ={s:>16,.2f}  СуммаФ1={f1:>16,.2f}  СуммаФ2={f2:>16,.2f}")

    # === проверки ===
    errs = []
    if abs(after["Сумма"] - before["Сумма"]) > 0.5:
        errs.append(f"Σ Сумма изменилась: {before['Сумма']:.2f} -> {after['Сумма']:.2f}")
    if abs(after["Сумма_Excel"] - before["Сумма_Excel"]) > 0.5:
        errs.append(f"Σ Сумма_Excel изменилась: {before['Сумма_Excel']:.2f} -> {after['Сумма_Excel']:.2f}")
    if abs(after["СуммаФ1"] + after["СуммаФ2"] - after["Сумма"]) > 0.5:
        errs.append(f"Σ(Ф1+Ф2)={after['СуммаФ1']+after['СуммаФ2']:.2f} != Σ Сумма={after['Сумма']:.2f}")
    forms_present = {f for f, *_ in dist}
    if "<пусто>" in forms_present:
        empty_rows = [d for d in dist if d[0] == "<пусто>"][0]
        errs.append(f"есть строки с пустой ФормаPL: {empty_rows[1]} строк, Σ={empty_rows[2]:,.2f}")
    if not ({"Форма1", "Форма2"} & forms_present):
        errs.append("ФормаPL не заполнена ни Форма1, ни Форма2")

    print("\n=== РЕЗУЛЬТАТ ===")
    if errs:
        for e in errs:
            print(f"  !! {e}")
        print("\n########## SMOKE FAIL ##########")
        sys.exit(1)
    print(f"  Σ Сумма неизменна: {before['Сумма']:,.2f} == {after['Сумма']:,.2f}")
    print(f"  Ф1+Ф2=Сумма: OK")
    print(f"  ФормаPL заполнена: {sorted(forms_present)}, пустых строк нет")
    print("\n########## SMOKE OK ##########")


if __name__ == "__main__":
    main()
