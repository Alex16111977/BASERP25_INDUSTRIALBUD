# -*- coding: utf-8 -*-
"""
СКРИПТ 25v2 — Классификация остаточных расхождений C1/C2/C3
                (на основе 20v2_full_discovery.csv)

КАТЕГОРИИ (из ANALYSIS_REPORT.md):
    C1: ВозвратОплатыКлиенту / ВозвратДенежныхСредствВДругуюОрганизацию (наш fix — должно быть ≈0)
    C2: ЗачётАванса при Реализации/Поступлении с ПереносАванса (наш fix — должно быть ≈0)
    C3: Прочее — известные классы из ANALYSIS_REPORT.md (ВводОстатков, ПриобретениеТоваровУслуг, и др.)
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import connect_erp, money, save_csv, ARTIFACTS_DIR

erp = connect_erp()


def f(x):
    try: return float(str(x).replace(",", ".").replace(" ", ""))
    except: return 0.0


# Загружаем 20v2
rows20 = []
with open(os.path.join(ARTIFACTS_DIR, "20v2_full_discovery.csv"), encoding="utf-8-sig", newline="") as fin:
    rows20 = list(csv.DictReader(fin, delimiter=";"))

print(f"Загружено {len(rows20)} строк из 20v2_full_discovery.csv")

# Группируем по уникальному документу
по_документу = {}
for r in rows20:
    ключ = (r["ТипДок"], r["Документ"])
    по_документу.setdefault(ключ, 0.0); по_документу[ключ] += f(r["Дельта"])
print(f"Уникальных документов: {len(по_документу)}")

# По типам
по_типам = {}
for (тип, имя), delta in по_документу.items():
    по_типам.setdefault(тип, []).append((имя, delta))

# Для каждого типа — один запрос: парсим Номер из имени и матчим
имя_to_op = {}
for тип, items in по_типам.items():
    if not тип:
        continue
    try:
        q = erp.NewObject("Запрос")
        q.Текст = f"""
        ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Д.Ссылка) КАК Имя,
                ПРЕДСТАВЛЕНИЕ(Д.ХозяйственнаяОперация) КАК ХозОп
        ИЗ Документ.{тип} КАК Д
        ГДЕ Д.Дата МЕЖДУ ДАТАВРЕМЯ(2025,12,1) И ДАТАВРЕМЯ(2026,12,31,23,59,59)
        """
        res = q.Выполнить().Выгрузить()
        for i in range(res.Количество()):
            rec = res.Получить(i)
            имя_to_op[(тип, str(rec.Имя))] = str(rec.ХозОп or "")
        print(f"  {тип}: загружено {res.Количество()} док (в 20v2: {len(items)})")
    except Exception as e:
        # Не у всех документов есть ХозОперация (например ВводОстатков, ВзаимозачётЗадолженности)
        print(f"  {тип}: нет ХозОп — {str(e)[:80]}")
        for имя, _ in items:
            имя_to_op[(тип, имя)] = ""

# Категоризация
C1_OPS = ("Возврат оплаты клиенту", "Возврат денежных средств в другую организацию",
          "ВозвратОплатыКлиенту", "ВозвратДенежныхСредствВДругуюОрганизацию")


def classify(тип, хозоп):
    хозоп = хозоп or ""
    if any(c in хозоп for c in C1_OPS):
        return "C1_ВозвратОплаты"
    if "Перенос" in хозоп or "перенос" in хозоп:
        return "C2_ПереносАванса"
    # C3 — расширенная классификация по типам
    if тип == "ВводОстатков":
        return "C3a_ВводОстатков"
    if тип == "ПриобретениеТоваровУслуг":
        return "C3b_ПриобретениеТоваровУслуг"
    if тип == "СписаниеБезналичныхДенежныхСредств":
        return "C3c_СписаниеБезнал"
    if тип == "ПоступлениеБезналичныхДенежныхСредств":
        return "C3d_ПоступлениеБезнал"
    if тип == "ПриходныйКассовыйОрдер":
        return "C3e_ПКО"
    if тип == "РасходныйКассовыйОрдер":
        return "C3f_РКО"
    if тип == "ВзаимозачетЗадолженности":
        return "C3g_Взаимозачёт"
    if тип == "РеализацияТоваровУслуг":
        return "C3h_РеализацияТовУсл"
    if тип == "АктВыполненныхРабот":
        return "C3i_АктВыпРабот"
    if тип == "АвансовыйОтчет":
        return "C3j_АвансовыйОтчёт"
    return f"C3z_Прочее"


cat_sums = {}
out_rows = []
for (тип, имя), delta in по_документу.items():
    op = имя_to_op.get((тип, имя), "<неизв>")
    cat = classify(тип, op)
    if cat not in cat_sums:
        cat_sums[cat] = {"count": 0, "sum": 0.0}
    cat_sums[cat]["count"] += 1; cat_sums[cat]["sum"] += delta
    out_rows.append({"ТипДок": тип, "Документ": имя, "ХозОперация": op,
                     "Δ": delta, "Категория": cat})

print("\n" + "=" * 90)
print("КЛАССИФИКАЦИЯ ОСТАТОЧНЫХ РАСХОЖДЕНИЙ ПО КАТЕГОРИЯМ")
print("=" * 90)
print(f"{'Кат':<35} {'Кол док':>10} {'Σ Δ':>22}")
print("-" * 90)
total = 0.0
for cat, v in sorted(cat_sums.items(), key=lambda x: -abs(x[1]["sum"])):
    print(f"{cat:<35} {v['count']:>10} {money(v['sum']):>22}")
    total += v["sum"]
print("-" * 90)
print(f"{'ИТОГО':<35} {sum(v['count'] for v in cat_sums.values()):>10} {money(total):>22}")

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

print("\nТоп-30 пар (ТипДок × ХозОп × Кат):")
print(f"{'Тип':<35} {'ХозОп':<30} {'Кат':<25} {'Кол':>5} {'Σ Δ':>18}")
print("-" * 115)
for r in pivot_rows[:30]:
    print(f"{r['ТипДок'][:35]:<35} {r['ХозОперация'][:30]:<30} {r['Категория']:<25} {r['КолДок']:>5} {money(r['Σ Δ']):>18}")

save_csv("25v2_root_cause_matrix", out_rows, ["ТипДок", "Документ", "ХозОперация", "Δ", "Категория"])
save_csv("25v2_root_cause_pivot", pivot_rows, ["ТипДок", "ХозОперация", "Категория", "КолДок", "Σ Δ"])
print(f"\nАртефакты: 25v2_root_cause_matrix.csv, 25v2_root_cause_pivot.csv")
