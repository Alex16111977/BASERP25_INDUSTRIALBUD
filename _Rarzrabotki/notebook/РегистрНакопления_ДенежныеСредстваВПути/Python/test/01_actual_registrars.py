# -*- coding: utf-8 -*-
"""
01 — Фактические регистраторы РН.ДенежныеСредстваВПути за весь период.

Цель: понять какие документы реально пишут в регистр (по типу + по количеству + по объёму),
независимо от того что заявлено в коде.

Артефакт: _artifacts/01_actual_registrars.csv
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, money, save_csv, get_type_name

erp = connect_erp()
S = erp.String

ORG = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
print(f"Орг: {S(ORG)}")

print("=" * 110)
print("01 — Фактические регистраторы РН.ДенежныеСредстваВПути")
print("=" * 110)

# === Группировка по типу регистратора ===
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG)
q.Текст = """
ВЫБРАТЬ
    Вп.Регистратор КАК Док,
    Вп.ВидДвижения КАК ВидДв,
    КОЛИЧЕСТВО(*) КАК Колво,
    СУММА(Вп.Сумма) КАК ΣСумма,
    МИНИМУМ(Вп.Период) КАК ПервыйПериод,
    МАКСИМУМ(Вп.Период) КАК ПоследнийПериод
ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Вп
ГДЕ Вп.Организация = &Орг
СГРУППИРОВАТЬ ПО Вп.Регистратор, Вп.ВидДвижения
"""
print("(запрос может занять минуту...)")
r = q.Выполнить().Выгрузить()
print(f"Уникальных пар (Регистратор × ВидДв): {r.Количество()}")

# Сводка по ТипДок
by_type = {}
for i in range(r.Количество()):
    rec = r.Получить(i)
    tp = get_type_name(erp, rec.Док)
    if not tp: continue
    by_type.setdefault(tp, {"docs": set(), "movements": 0, "Σ_приход": 0.0, "Σ_расход": 0.0,
                            "first": None, "last": None})
    by_type[tp]["docs"].add(str(S(rec.Док)))
    by_type[tp]["movements"] += int(rec.Колво or 0)
    s = float(rec.ΣСумма or 0)
    if S(rec.ВидДв) == "Приход":
        by_type[tp]["Σ_приход"] += s
    else:
        by_type[tp]["Σ_расход"] += s
    p1 = rec.ПервыйПериод; p2 = rec.ПоследнийПериод
    if by_type[tp]["first"] is None or p1 < by_type[tp]["first"]:
        by_type[tp]["first"] = p1
    if by_type[tp]["last"] is None or p2 > by_type[tp]["last"]:
        by_type[tp]["last"] = p2

print(f"\nУникальных ТипДок: {len(by_type)}\n")
print(f"{'ТипДок':<50}{'Док-в':>8}{'Движ':>8}{'Σ Приход':>20}{'Σ Расход':>20}{'Первая дата':<22}{'Последняя дата':<22}")
print("-" * 150)

rows = []
for tp, v in sorted(by_type.items(), key=lambda x: -(x[1]["Σ_приход"] + x[1]["Σ_расход"])):
    d_count = len(v["docs"])
    print(f"{tp:<50}{d_count:>8}{v['movements']:>8}{money(v['Σ_приход']):>20}{money(v['Σ_расход']):>20}"
          f"{str(v['first'])[:19]:<22}{str(v['last'])[:19]:<22}")
    rows.append({"ТипДок": tp, "ДокКолво": d_count, "ДвижКолво": v['movements'],
                 "ΣПриход": v['Σ_приход'], "ΣРасход": v['Σ_расход'],
                 "ПервыйПериод": str(v['first']), "ПоследнийПериод": str(v['last'])})

save_csv("01_actual_registrars", rows,
         ["ТипДок", "ДокКолво", "ДвижКолво", "ΣПриход", "ΣРасход", "ПервыйПериод", "ПоследнийПериод"])
print("\nАртефакт: 01_actual_registrars.csv")

# === Сводка по ВидПереводаДенежныхСредств ===
print("\n\n=== ВидПереводаДенежныхСредств × ТипДок ===")
q.Текст = """
ВЫБРАТЬ
    Вп.ВидПереводаДенежныхСредств КАК ВидПеревода,
    Вп.Регистратор КАК Док,
    КОЛИЧЕСТВО(*) КАК Колво
ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Вп
ГДЕ Вп.Организация = &Орг
СГРУППИРОВАТЬ ПО Вп.ВидПереводаДенежныхСредств, Вп.Регистратор
"""
r = q.Выполнить().Выгрузить()
matrix = {}
for i in range(r.Количество()):
    rec = r.Получить(i)
    vp = str(S(rec.ВидПеревода) or "(пусто)")
    tp = get_type_name(erp, rec.Док)
    if not tp: continue
    matrix.setdefault(vp, {}).setdefault(tp, 0)
    matrix[vp][tp] += int(rec.Колво or 0)

print(f"{'ВидПеревода':<40}{'ТипДок':<45}{'Движ':>8}")
print("-" * 93)
mtx_rows = []
for vp in sorted(matrix.keys()):
    for tp, cnt in sorted(matrix[vp].items(), key=lambda x: -x[1]):
        print(f"{vp[:38]:<40}{tp[:43]:<45}{cnt:>8}")
        mtx_rows.append({"ВидПеревода": vp, "ТипДок": tp, "Движ": cnt})

save_csv("01_vp_x_type", mtx_rows, ["ВидПеревода", "ТипДок", "Движ"])
print("\nАртефакт: 01_vp_x_type.csv")
