# -*- coding: utf-8 -*-
"""
СКРИПТ 24 (Phase 0 финал) — Сборка DISCOVERY_REPORT.md

ЧТО ДЕЛАЕТ:
    Читает 20-23 артефакты, сводит в MD-отчёт.
    Gate: должен показать что Σ Δ ≠ 0 и затронуто >1 подразделения.
"""
import sys, io, csv, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import ARTIFACTS_DIR, money

DOCS_DIR = os.path.abspath(os.path.join(os.path.dirname(ARTIFACTS_DIR), "..", "..", "docs"))


def load(name):
    p = os.path.join(ARTIFACTS_DIR, f"{name}.csv")
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def f(x):
    try: return float(str(x).replace(",", ".").replace(" ", ""))
    except: return 0.0


# 20
r20 = load("20_full_discovery")
by_podr = {}
for r in r20:
    by_podr.setdefault(r["Подразделение"], 0.0)
    by_podr[r["Подразделение"]] += f(r["Дельта"])
sum_total_20 = sum(by_podr.values())

# 21
r21 = load("21_perenosavansa_rows")
sum_pa_rsk = sum(f(r["Σ_Δ_отрезано"]) for r in r21 if r["Регистр"] == "РСК")
sum_pa_rsp = sum(f(r["Σ_Δ_отрезано"]) for r in r21 if r["Регистр"] == "РСП")

# 22
r22 = load("22_typed_breakdown")

# 23
r23 = load("23_etalon_uprbalance")
sum_otch = sum(f(r["КО_отчет"]) for r in r23)
sum_nash = sum(f(r["КО_наш"]) for r in r23)

lines = []
lines.append("# DISCOVERY_REPORT — Phase 0\n")
lines.append("> Сгенерировано скриптом 24 на основе 20-23 артефактов.\n")
lines.append("## Сводка\n")
lines.append(f"- **Σ Δ ПАП vs РСКПС+РСППС за 2025**: {money(sum_total_20)} UAH")
lines.append(f"- **Подразделений с расхождениями**: {len(by_podr)}")
lines.append(f"- **Документов-первичек с расхождениями**: {len(r20)} пар (Подразделение × Документ)")
lines.append(f"- **РСК: Σ ПереносАванса отрезано фильтром**: {money(sum_pa_rsk)}")
lines.append(f"- **РСП: Σ ПереносАванса отрезано фильтром**: {money(sum_pa_rsp)}")
lines.append(f"- **Сверка со штатным Отчёт.УпрБаланс**: КО_наш={money(sum_nash)}, КО_отчёт={money(sum_otch)}, Δ={money(sum_nash - sum_otch)}\n")

lines.append("## Топ-15 подразделений по |Σ Δ|\n")
lines.append("| Подразделение | Σ Δ |")
lines.append("|---|---:|")
for подр, d in sorted(by_podr.items(), key=lambda x: -abs(x[1]))[:15]:
    lines.append(f"| {подр or '<пусто>'} | {money(d)} |")

lines.append("\n## Топ-15 пар (ТипДок × ХозОперация)\n")
lines.append("| ТипДок | ХозОперация | Кол | Σ Δ |")
lines.append("|---|---|---:|---:|")
r22.sort(key=lambda r: -abs(f(r["Σ Δ"])))
for r in r22[:15]:
    lines.append(f"| {r['ТипДок']} | {r['ХозОперация']} | {r['КолДокументов']} | {money(f(r['Σ Δ']))} |")

lines.append("\n## Gate-критерий Phase 1\n")
if abs(sum_total_20) > 0.01 and len(by_podr) > 1:
    lines.append("✓ **PASS** — Σ Δ ≠ 0 и затронуто >1 подразделения. Переходим к Phase 1 (Analysis).")
else:
    lines.append("⚠️ **FAIL** — расхождения не зафиксированы или одно подразделение. Пересмотреть spec.")

lines.append("\n## Артефакты\n")
lines.append("- `_artifacts/20_full_discovery.csv` — полная карта расхождений")
lines.append("- `_artifacts/21_perenosavansa_rows.csv` — строки ПереносАванса")
lines.append("- `_artifacts/22_typed_breakdown.csv` — pivot ТипДок × ХозОп")
lines.append("- `_artifacts/23_etalon_uprbalance.csv` — сверка со штатным отчётом")

out = os.path.join(DOCS_DIR, "DISCOVERY_REPORT.md")
os.makedirs(DOCS_DIR, exist_ok=True)
with open(out, "w", encoding="utf-8") as fout:
    fout.write("\n".join(lines))

print(f"DISCOVERY_REPORT записан: {out}")
print(f"Σ Δ всего: {money(sum_total_20)}, подразделений: {len(by_podr)}")
