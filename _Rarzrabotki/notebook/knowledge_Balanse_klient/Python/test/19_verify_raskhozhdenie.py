# -*- coding: utf-8 -*-
"""
СКРИПТ 19 — Verify: проверка что Расхождение=Истина уменьшилось/исчезло

Сравнивает текущее состояние с baseline (snapshot до перепроведения).
Если baseline нет — сохраняет текущее как baseline.

Параметр: BASELINE | VERIFY
  python 19_verify_raskhozhdenie.py BASELINE   — сохранить текущее
  python 19_verify_raskhozhdenie.py VERIFY     — сравнить с baseline

Артефакты:
  _artifacts/19_baseline.json  — snapshot до правки
  _artifacts/19_current.json   — snapshot сейчас
  _artifacts/19_diff.md        — отчёт diff
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, get_refs, money, ARTIFACTS_DIR

erp = connect_erp()
refs = get_refs(erp)
ORG = refs["Орг"]
S = erp.String

mode = sys.argv[1].upper() if len(sys.argv) > 1 else "VERIFY"
print(f"СКРИПТ 19 — Verify Расхождение=Истина (mode={mode})\n")

def snapshot():
    """Текущая сводка по (Месяц × Source × ПодрКод × СтатьяКод)."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", ORG)
    q.Текст = """
ВЫБРАТЬ
    Т.Регистратор.Месяц КАК Месяц,
    Т.Source КАК SourceRef,
    Т.Подразделение.Код КАК ПодрКод,
    Т.Статья.Код КАК СтатьяКод,
    КОЛИЧЕСТВО(*) КАК Колво,
    СУММА(Т.СуммаКонечныйОстаток) КАК СуммаKM
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Т
ГДЕ Т.Расхождение = ИСТИНА
    И Т.Организация = &Орг
    И Т.Регистратор.Месяц МЕЖДУ ДАТАВРЕМЯ(2025,12,1) И ДАТАВРЕМЯ(2026,12,31,23,59,59)
СГРУППИРОВАТЬ ПО
    Т.Регистратор.Месяц, Т.Source,
    Т.Подразделение.Код, Т.Статья.Код
"""
    r = q.Выполнить().Выгрузить()
    rows = []
    for i in range(r.Количество()):
        rec = r.Получить(i)
        мес_str = rec.Месяц.strftime("%Y-%m") if hasattr(rec.Месяц, "strftime") else str(rec.Месяц)[:7]
        src_str = str(S(rec.SourceRef)) if erp.ЗначениеЗаполнено(rec.SourceRef) else ""
        rows.append({
            "Месяц": мес_str,
            "Source": src_str,
            "ПодрКод": str(rec.ПодрКод or ""),
            "СтатьяКод": str(rec.СтатьяКод or ""),
            "Колво": int(rec.Колво or 0),
            "СуммаKM": float(rec.СуммаKM or 0),
        })
    return rows

current = snapshot()
baseline_path = os.path.join(ARTIFACTS_DIR, "19_baseline.json")
current_path = os.path.join(ARTIFACTS_DIR, "19_current.json")

with open(current_path, "w", encoding="utf-8") as f:
    json.dump(current, f, ensure_ascii=False, indent=2)
print(f"Сохранено текущее состояние: {current_path}")
print(f"Записей с Расхождение=Истина: {len(current)}")

# Сводка по месяцу × source
by_ms = {}
for r in current:
    k = (r["Месяц"], r["Source"])
    if k not in by_ms:
        by_ms[k] = {"Колво": 0, "ΣABS": 0.0, "ΣSign": 0.0}
    by_ms[k]["Колво"] += r["Колво"]
    by_ms[k]["ΣABS"] += abs(r["СуммаKM"])
    by_ms[k]["ΣSign"] += r["СуммаKM"]

print(f"\nСводка (Месяц × Source):")
print(f"{'Месяц':<10}{'Source':<35}{'Колво':>8}{'ΣABS':>20}{'ΣSign':>20}")
print("-" * 95)
total_abs = total_sign = 0.0
total_kol = 0
for (мес, src), v in sorted(by_ms.items()):
    print(f"{мес:<10}{src[:33]:<35}{v['Колво']:>8}{money(v['ΣABS']):>20}{money(v['ΣSign']):>20}")
    total_abs += v["ΣABS"]; total_sign += v["ΣSign"]; total_kol += v["Колво"]
print("-" * 95)
print(f"{'ИТОГО':<10}{'':<35}{total_kol:>8}{money(total_abs):>20}{money(total_sign):>20}")

# Если mode=BASELINE — копируем current в baseline
if mode == "BASELINE":
    with open(baseline_path, "w", encoding="utf-8") as f:
        json.dump(current, f, ensure_ascii=False, indent=2)
    print(f"\n→ Сохранено как baseline: {baseline_path}")
    sys.exit(0)

# VERIFY: сравнить с baseline
if not os.path.exists(baseline_path):
    print(f"\n[!] Baseline не найден ({baseline_path}). Запусти сначала с BASELINE.")
    sys.exit(2)

with open(baseline_path, encoding="utf-8") as f:
    baseline = json.load(f)

print(f"\nBaseline: {len(baseline)} записей, Current: {len(current)}")
print(f"Δ записей: {len(current) - len(baseline)}")

# По месяцам
bb = {}
for r in baseline:
    k = (r["Месяц"], r["Source"])
    bb[k] = bb.get(k, 0) + abs(r["СуммаKM"])
cc = {}
for r in current:
    k = (r["Месяц"], r["Source"])
    cc[k] = cc.get(k, 0) + abs(r["СуммаKM"])
all_keys = sorted(set(bb) | set(cc))
print(f"\nИзменения по (Месяц × Source) — ΣABS:")
print(f"{'Месяц':<10}{'Source':<35}{'Baseline':>20}{'Current':>20}{'Δ':>20}")
print("-" * 110)
for k in all_keys:
    b = bb.get(k, 0); c = cc.get(k, 0)
    marker = "✓ улучшение" if c < b else ("✗ ухудшение" if c > b else "= без изм")
    print(f"{k[0]:<10}{k[1][:33]:<35}{money(b):>20}{money(c):>20}{money(c-b):>20}  {marker}")
