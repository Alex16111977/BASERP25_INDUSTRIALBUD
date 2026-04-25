# -*- coding: utf-8 -*-
"""
Unit-тест жадного алгоритма разбиения дисбаланса на пары.

Алгоритм:
  - Дебиторы = [{Подр, Контроль>0}]  → отсортировать по убыванию Контроль
  - Кредиторы = [{Подр, |Контроль<0|}] → отсортировать по убыванию |Контроль|
  - Итерация: head(Дебиторы) + head(Кредиторы), сумма пары = min(...), уменьшить; кто обнулился — удалить.

Инварианты:
  - Сумма всех пар == Сумма всех Контроль>0 == Сумма всех |Контроль<0|
  - Число пар <= |Дебиторы| + |Кредиторы| - 1
  - Каждая пара имеет Сумма > 0
"""
import sys
from typing import List, Dict, Tuple

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def greedy_split(disbalance: Dict[str, float], eps: float = 0.005) -> List[Dict]:
    """
    disbalance: {"Подр1": +100.0, "Подр2": -60.0, "Подр3": -40.0}
    Возвращает [{"Дт": "Подр1", "Кт": "Подр2", "Сумма": 60.0}, ...]
    """
    debitors = [[k, round(v, 2)] for k, v in disbalance.items() if round(v, 2) > 0]
    creditors = [[k, round(-v, 2)] for k, v in disbalance.items() if round(v, 2) < 0]
    debitors.sort(key=lambda x: -x[1])
    creditors.sort(key=lambda x: -x[1])

    pairs = []
    i = j = 0
    while i < len(debitors) and j < len(creditors):
        d = debitors[i]
        c = creditors[j]
        summa = round(min(d[1], c[1]), 2)
        if summa < eps:
            # защита от бесконечного цикла на копейках
            if d[1] < eps:
                i += 1
                continue
            if c[1] < eps:
                j += 1
                continue
        pairs.append({"Дт": d[0], "Кт": c[0], "Сумма": summa})
        d[1] = round(d[1] - summa, 2)
        c[1] = round(c[1] - summa, 2)
        if d[1] < eps:
            i += 1
        if c[1] < eps:
            j += 1
    return pairs


# === Тест-кейсы ===
def test_case(name, disbalance, expected_pairs):
    print(f"\n--- {name} ---")
    print(f"  Вход: {disbalance}")
    pairs = greedy_split(disbalance)
    print(f"  Получено {len(pairs)} пар(ы):")
    for p in pairs:
        print(f"    Дт={p['Дт']} <- Кт={p['Кт']}: {p['Сумма']}")

    # Инвариант 1: сумма пар = сумма положительных контролей
    sum_pos = round(sum(v for v in disbalance.values() if v > 0), 2)
    sum_pairs = round(sum(p["Сумма"] for p in pairs), 2)
    ok1 = abs(sum_pairs - sum_pos) < 0.01
    print(f"  Сумма пар = {sum_pairs} vs Сумма Дт = {sum_pos}: {'OK' if ok1 else 'FAIL'}")

    # Инвариант 2: количество пар
    n = len([v for v in disbalance.values() if round(v, 2) != 0])
    ok2 = len(pairs) <= n
    print(f"  Пар ({len(pairs)}) <= N-1 ({n-1}) или N ({n}): {'OK' if ok2 else 'FAIL'}")

    # Инвариант 3: совпадение с ожидаемым числом пар
    ok3 = len(pairs) == expected_pairs
    print(f"  Ожидалось пар = {expected_pairs}: {'OK' if ok3 else 'FAIL'}")

    return ok1 and ok2 and ok3


results = []

# 1. Минимальный кейс: 1+1 (старое поведение) = 1 пара
results.append(test_case(
    "Min: 1 Дт + 1 Кт",
    {"A": 100.0, "B": -100.0},
    expected_pairs=1,
))

# 2. Реальный кейс ІБ00-008058: 3 Дт + 1 Кт = 3 пары
results.append(test_case(
    "ІБ00-008058: 3 Дт + 1 Кт (ЦО)",
    {
        "Строительство": 13107.50,
        "Спецтехника": 1123.00,
        "Логистика": 375.00,
        "ЦО": -14605.50,
    },
    expected_pairs=3,
))

# 3. Равные 2+2 = ровно 2 пары
results.append(test_case(
    "2 Дт + 2 Кт равные",
    {"A": 100.0, "B": 100.0, "X": -100.0, "Y": -100.0},
    expected_pairs=2,
))

# 4. 2+2 неравные = возможно 3 пары
results.append(test_case(
    "2 Дт + 2 Кт неравные",
    {"A": 150.0, "B": 50.0, "X": -120.0, "Y": -80.0},
    expected_pairs=3,
))

# 5. Копеечное расхождение (обрезание по 0.01)
results.append(test_case(
    "Копеечные остатки",
    {"A": 100.01, "B": 0.005, "X": -100.015},
    expected_pairs=1,  # сумма = 100.01, остаток 0.005 на B и -0.005 на X — ниже порога
))

# 6. Пустой вход
results.append(test_case(
    "Пустой",
    {},
    expected_pairs=0,
))

# 7. Только 1 подразделение (невалидный случай, должен дать 0 пар)
results.append(test_case(
    "Только Дт без Кт",
    {"A": 100.0},
    expected_pairs=0,
))

print("\n" + "=" * 60)
print(f"ИТОГ: {sum(results)}/{len(results)} тестов прошли")
if all(results):
    print("  ВСЕ ТЕСТЫ OK")
else:
    print("  ЕСТЬ ПРОВАЛЫ")
