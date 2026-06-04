# -*- coding: utf-8 -*-
"""
СКРИПТ 03 — Reconcile остатков ПАП vs РСКПС на 4 контрольных точках

ЧТО ИЩЕТ:
    Δ = ПАП − РСКПС на (НМ_ноя, КМ_ноя=НМ_дек, КМ_дек).
    Цель: подтвердить что эталон ноября сходится (Δ=0) и зафиксировать
    размер декабрьского расхождения (Δ_КМ_дек ≠ 0).

КАК СЧИТАЕТ:
    Читает артефакты 01_pap_balances.json и 02_rsk_balances.json,
    выводит таблицу сравнения и сохраняет:
        _artifacts/03_balance_reconciliation.csv

ЧТО ДАЁТ:
    print: 3 строки сверки + размер расхождения декабря в UAH.
    csv: ContPoint, ПАП, РСКПС, Δ, %расхождения.
"""
from _common import load_json, save_csv, money


def main():
    print("=" * 80)
    print("СКРИПТ 03 — Reconcile ПАП vs РСКПС")
    print("=" * 80)

    pap = load_json("01_pap_balances")
    rsk = load_json("02_rsk_balances")

    points = [
        ("01.11.2025 НМ ноя", pap["ноябрь"]["НМ"], rsk["ноябрь"]["ДолгКлиентов_НМ"]),
        ("30.11.2025 КМ ноя", pap["контр_30_11_2025"], rsk["контр_30_11_2025_ДолгКлиентов"]),
        ("01.12.2025 НМ дек", pap["контр_01_12_2025"], rsk["декабрь"]["ДолгКлиентов_НМ"]),
        ("31.12.2025 КМ дек", pap["контр_31_12_2025"], rsk["контр_31_12_2025_ДолгКлиентов"]),
    ]

    rows = []
    print(f"\n{'Точка':<22} {'ПАП':>18} {'РСКПС':>18} {'Δ=ПАП-РСКПС':>18} {'%':>10}")
    print("-" * 90)
    for name, p, r in points:
        delta = p - r
        pct = (delta / r * 100) if r else 0
        status = "✓ OK" if abs(delta) < 0.01 else "⚠️ РАСХОЖДЕНИЕ"
        print(
            f"{name:<22} {money(p):>18} {money(r):>18} {money(delta):>18} "
            f"{pct:>9.4f}% {status}"
        )
        rows.append({
            "Точка": name,
            "ПАП": p,
            "РСКПС": r,
            "Delta": delta,
            "Pct": pct,
            "Status": status,
        })

    # Главный вывод
    delta_кон_дек = points[3][1] - points[3][2]
    print()
    print(f"{'═' * 80}")
    if abs(delta_кон_дек) > 0.01:
        print(f"🎯 РАСХОЖДЕНИЕ декабря: ПАП={money(points[3][1])} vs РСКПС={money(points[3][2])}")
        print(f"   Δ = {money(delta_кон_дек)} UAH ({'ПАП БОЛЬШЕ' if delta_кон_дек > 0 else 'РСКПС БОЛЬШЕ'})")
        # Какой месяц вносит расхождение
        delta_дек_только = (points[3][1] - points[3][2]) - (points[2][1] - points[2][2])
        print(f"   из них появилось в декабре: {money(delta_дек_только)} UAH")
    else:
        print(f"✓ КМ декабря: ПАП = РСКПС (расхождения нет)")
    print(f"{'═' * 80}")

    path = save_csv(
        "03_balance_reconciliation",
        rows,
        ["Точка", "ПАП", "РСКПС", "Delta", "Pct", "Status"],
    )
    print(f"\nАртефакт сохранён: {path}")


if __name__ == "__main__":
    main()
