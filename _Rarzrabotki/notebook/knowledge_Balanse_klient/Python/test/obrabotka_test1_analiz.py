# -*- coding: utf-8 -*-
"""
TEST 1 — Анализ через Python COM вызов внешнего отчёта.

Вызывает АнализРасхождений() и сравнивает результат ТЧ с baseline
(obrabotka_baseline.json, зафиксированным в Task 1).

Acceptance:
  - ТЧ содержит ровно столько же строк сколько baseline
  - Σ Δ по подразделениям совпадает с baseline до копейки
"""
import sys, io, json, os, datetime as dt
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, ARTIFACTS_DIR

ERF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ОбработкаДисбалансаПоСтатьямБаланса.epf"
BASELINE_PATH = os.path.join(ARTIFACTS_DIR, "obrabotka_baseline.json")

erp = connect_erp()
# Теперь это обработка (ExternalDataProcessor), а не отчёт
report = erp.ВнешниеОбработки.Создать(ERF_PATH, False)

report.НачалоПериода = dt.datetime(2026, 4, 1)
report.ОкончаниеПериода = dt.datetime(2026, 4, 30, 23, 59, 59)

print(f"Запуск АнализРасхождений() для апреля 2026...")
report.АнализРасхождений()

tch = report.ДокументыРасхождения
print(f"ТЧ ДокументыРасхождения: {tch.Количество()} строк")

# Загрузить baseline
with open(BASELINE_PATH, "r", encoding="utf-8") as f:
    baseline = json.load(f)

baseline_rows = baseline["total_rows"]
baseline_sums = baseline["sums_by_podr"]

# Собрать Σ Δ из ТЧ обработки — фильтруем только наши 9 подразделений + статья ЗПП
TARGET_CODES = set(baseline_sums.keys())
actual_sums = {}
actual_count_in_scope = 0
total_rows = tch.Количество()
for i in range(total_rows):
    row = tch.Получить(i)
    podr_code = str(row.Подразделение.Код) if row.Подразделение else ""
    statya_name = str(erp.String(row.Статья)) if row.Статья else ""
    if podr_code in TARGET_CODES and statya_name == "Задолженность перед поставщиками":
        actual_sums[podr_code] = actual_sums.get(podr_code, 0) + float(row.Дельта)
        actual_count_in_scope += 1

# Сравнить
print(f"\n=== Acceptance (фильтр: 9 подр ЗПП) ===")
print(f"Baseline: {baseline_rows} строк / актуально в scope: {actual_count_in_scope}")
print(f"Всего строк ТЧ (вкл. другие статьи/подр): {total_rows}")
errors = 0
if actual_count_in_scope != baseline_rows:
    print(f"FAIL: rows in scope {actual_count_in_scope} != baseline {baseline_rows}")
    errors += 1
else:
    print(f"OK: rows in scope = {actual_count_in_scope}")

for code, expected in sorted(baseline_sums.items()):
    actual = actual_sums.get(code, 0)
    diff = abs(actual - expected)
    if diff < 0.01:
        print(f"OK   {code}: Σ Δ = {actual:+,.2f} (== baseline)")
    else:
        print(f"FAIL {code}: ожидали {expected:+,.2f}, факт {actual:+,.2f}, diff {diff:.2f}")
        errors += 1

if errors == 0:
    print("\n*** TEST 1 PASS ***")
    sys.exit(0)
else:
    print(f"\n*** TEST 1 FAIL ({errors} errors) ***")
    sys.exit(1)
