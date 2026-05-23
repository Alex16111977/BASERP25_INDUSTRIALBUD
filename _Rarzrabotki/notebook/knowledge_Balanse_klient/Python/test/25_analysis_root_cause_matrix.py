# -*- coding: utf-8 -*-
"""
СКРИПТ 25 (Phase 1) — Классификация документов по корневой причине

КАТЕГОРИИ:
    C1: ВозвратОплатыКлиенту / ВозвратДенежныхСредствВДругуюОрганизацию
    C2: ЗачётАванса (Реализация/Поступление с переносом аванса)
    C3: Прочее (требует расследования)

Подход: для каждого типа документа из 20-арт делаем ОДИН запрос на ВСЕ
документы этого типа за 2025 (получаем Имя→ХозОп карту). Затем матчим
каждый документ из 20-арт по Имени.
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, money, save_csv, ARTIFACTS_DIR

erp = connect_erp()


def f(x):
    try: return float(str(x).replace(",", ".").replace(" ", ""))
    except: return 0.0


# Загружаем 20-арт
rows20 = []
with open(os.path.join(ARTIFACTS_DIR, "20_full_discovery.csv"), encoding="utf-8-sig", newline="") as fin:
    rows20 = list(csv.DictReader(fin, delimiter=";"))

print(f"Загружено {len(rows20)} строк из 20_full_discovery.csv")

# Группируем по ТипДок
по_типам = {}
for r in rows20:
    тип = r["ТипДок"]
    по_типам.setdefault(тип, []).append(r)

# Для каждого типа — один запрос на ХозОп
имя_to_op = {}  # (тип, имя) → хозоп
q = erp.NewObject("Запрос")
for тип, items in по_типам.items():
    if not тип:
        continue
    try:
        q.Текст = f"""
        ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Д.Ссылка) КАК Имя,
                ПРЕДСТАВЛЕНИЕ(Д.ХозяйственнаяОперация) КАК ХозОп
        ИЗ Документ.{тип} КАК Д
        ГДЕ Д.Дата МЕЖДУ ДАТАВРЕМЯ(2025,1,1) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
        """
        res = q.Выполнить().Выгрузить()
        for i in range(res.Количество()):
            rec = res.Получить(i)
            имя_to_op[(тип, str(rec.Имя))] = str(rec.ХозОп or "")
        print(f"  {тип}: {res.Количество()} док (тип в 20-арт: {len(items)})")
    except Exception as e:
        print(f"  {тип}: ОШИБКА — {e}")


# Категории
C1_OPS = ("Возврат оплаты клиенту", "Возврат денежных средств в другую организацию",
          "ВозвратОплатыКлиенту", "ВозвратДенежныхСредствВДругуюОрганизацию")
C2_OPS = ("Реализация клиенту", "Покупка у поставщика", "РеализацияКлиенту",
          "ПокупкаУПоставщика", "ПостоплатаПоставщику")


def classify(тип, хозоп):
    хозоп = хозоп or ""
    if any(c in хозоп for c in C1_OPS):
        return "C1"
    if "Перенос" in хозоп or "перенос" in хозоп:
        return "C2"
    if тип in ("РеализацияТоваровУслуг", "ПриобретениеТоваровУслуг",
               "АктВыполненныхРабот", "ПриобретениеУслугПрочихАктивов",
               "ПоступлениеБезналичныхДенежныхСредств") and any(c in хозоп for c in C2_OPS):
        return "C2"
    return "C3"


# Группируем по уникальным документам с Σ Δ
по_документу = {}
for r in rows20:
    ключ = (r["ТипДок"], r["Документ"])
    по_документу.setdefault(ключ, 0.0); по_документу[ключ] += f(r["Дельта"])

# Классифицируем
cat_sums = {"C1": {"count": 0, "sum": 0.0, "examples": []},
            "C2": {"count": 0, "sum": 0.0, "examples": []},
            "C3": {"count": 0, "sum": 0.0, "examples": []}}

out_rows = []
for (тип, имя), delta in по_документу.items():
    op = имя_to_op.get((тип, имя), "<неизв>")
    cat = classify(тип, op)
    cat_sums[cat]["count"] += 1; cat_sums[cat]["sum"] += delta
    if len(cat_sums[cat]["examples"]) < 5:
        cat_sums[cat]["examples"].append(f"{тип}/{op[:30]}")
    out_rows.append({"ТипДок": тип, "Документ": имя, "ХозОперация": op, "Δ": delta, "Категория": cat})

out_rows.sort(key=lambda r: -abs(r["Δ"]))

print("\n" + "=" * 100)
print("КЛАССИФИКАЦИЯ ПО КАТЕГОРИЯМ")
print("=" * 100)
print(f"{'Кат':<5} {'Кол док':>10} {'Σ Δ':>20} Примеры")
print("-" * 100)
for cat in ("C1", "C2", "C3"):
    v = cat_sums[cat]
    examples = "; ".join(v["examples"][:3])
    print(f"{cat:<5} {v['count']:>10} {money(v['sum']):>20} {examples[:60]}")

total = sum(v["sum"] for v in cat_sums.values())
print(f"\nΣ всех категорий: {money(total)}")

# Pivot по (Тип, ХозОп, Кат)
pivot = {}
for r in out_rows:
    key = (r["ТипДок"], r["ХозОперация"], r["Категория"])
    pivot.setdefault(key, {"count": 0, "sum": 0.0})
    pivot[key]["count"] += 1; pivot[key]["sum"] += r["Δ"]

pivot_rows = [
    {"ТипДок": k[0], "ХозОперация": k[1], "Категория": k[2],
     "КолДок": v["count"], "Σ Δ": v["sum"]}
    for k, v in sorted(pivot.items(), key=lambda x: -abs(x[1]["sum"]))
]

print("\nТоп-25 пар (ТипДок × ХозОп × Кат):")
print(f"{'Тип':<35} {'ХозОп':<28} {'Кат':<4} {'Кол':>6} {'Σ Δ':>20}")
print("-" * 105)
for r in pivot_rows[:25]:
    print(f"{r['ТипДок'][:35]:<35} {r['ХозОперация'][:28]:<28} {r['Категория']:<4} {r['КолДок']:>6} {money(r['Σ Δ']):>20}")

save_csv("25_root_cause_matrix", out_rows, ["ТипДок", "Документ", "ХозОперация", "Δ", "Категория"])
save_csv("25_root_cause_pivot", pivot_rows, ["ТипДок", "ХозОперация", "Категория", "КолДок", "Σ Δ"])
print(f"\nАртефакты: 25_root_cause_matrix.csv, 25_root_cause_pivot.csv")
