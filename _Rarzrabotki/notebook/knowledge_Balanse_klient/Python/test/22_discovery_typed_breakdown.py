# -*- coding: utf-8 -*-
"""
СКРИПТ 22 (Phase 0) — Pivot Σ Δ по (ТипДокумента, ХозОперация)

ЧТО ДАЁТ:
    _artifacts/22_typed_breakdown.csv
    print: матрицу типов документов и операций.
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, money, save_csv, ARTIFACTS_DIR

erp = connect_erp()

rows = []
with open(os.path.join(ARTIFACTS_DIR, "20_full_discovery.csv"), encoding="utf-8-sig", newline="") as f:
    rows = list(csv.DictReader(f, delimiter=";"))

print(f"Загружено {len(rows)} строк из 20_full_discovery.csv")

# Сборка по уникальным документам — суммарная Δ
по_документу = {}
for r in rows:
    док_имя = r["Документ"]
    тип = r["ТипДок"]
    дельта = float(str(r["Дельта"]).replace(",", ".").replace(" ", ""))
    ключ = (тип, док_имя)
    по_документу.setdefault(ключ, 0.0); по_документу[ключ] += дельта

print(f"Уникальных документов с расхождением: {len(по_документу)}")

# Группируем документы по типу
по_типам = {}
for (тип, док_имя), delta in по_документу.items():
    по_типам.setdefault(тип, []).append((док_имя, delta))

# Для каждого типа пакетно получаем ХозОперацию (если у документа есть)
matrix = {}  # {(тип, хозоп): {"count":n, "sum_delta":x}}
for тип, документы in по_типам.items():
    print(f"  {тип}: {len(документы)} док")
    if not тип:
        ключ = (тип, "<нет ХозОп>")
        matrix.setdefault(ключ, {"count": 0, "sum_delta": 0.0})
        for _, d in документы:
            matrix[ключ]["count"] += 1; matrix[ключ]["sum_delta"] += d
        continue
    # Пакетный запрос ХозОперации
    try:
        q = erp.NewObject("Запрос")
        q.Текст = f'''
        ВЫБРАТЬ Д.Ссылка, ПРЕДСТАВЛЕНИЕ(Д.Ссылка) КАК Имя, ПРЕДСТАВЛЕНИЕ(Д.ХозяйственнаяОперация) КАК ХозОп
        ИЗ Документ.{тип} КАК Д
        ГДЕ ПРЕДСТАВЛЕНИЕ(Д.Ссылка) В (&Имена)
        '''
        имена = erp.NewObject("Массив")
        for имя, _ in документы:
            имена.Добавить(имя)
        q.УстановитьПараметр("Имена", имена)
        res = q.Выполнить().Выгрузить()
        d_to_op = {}
        for i in range(res.Количество()):
            rec = res.Получить(i)
            d_to_op[str(rec.Имя)] = str(rec.ХозОп)
    except Exception:
        d_to_op = {}

    for имя, delta in документы:
        op = d_to_op.get(имя, "<без ХозОп>")
        ключ = (тип, op)
        matrix.setdefault(ключ, {"count": 0, "sum_delta": 0.0})
        matrix[ключ]["count"] += 1; matrix[ключ]["sum_delta"] += delta

out_rows = [
    {"ТипДок": k[0], "ХозОперация": k[1], "КолДокументов": v["count"], "Σ Δ": v["sum_delta"]}
    for k, v in sorted(matrix.items(), key=lambda x: -abs(x[1]["sum_delta"]))
]
path = save_csv("22_typed_breakdown", out_rows, ["ТипДок", "ХозОперация", "КолДокументов", "Σ Δ"])

print("\nМатрица (ТипДок × ХозОперация), топ-30 по |Σ Δ|:")
print(f"{'ТипДок':<40} {'ХозОп':<35} {'Кол':>6} {'Σ Δ':>16}")
print("-" * 105)
for r in out_rows[:30]:
    print(f"{r['ТипДок'][:40]:<40} {r['ХозОперация'][:35]:<35} {r['КолДокументов']:>6} {money(r['Σ Δ']):>16}")

print(f"\nАртефакт: {path}")
