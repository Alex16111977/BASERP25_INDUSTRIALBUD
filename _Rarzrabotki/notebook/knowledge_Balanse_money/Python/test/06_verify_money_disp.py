# -*- coding: utf-8 -*-
"""
СКРИПТ 06 — Verify: BASELINE/VERIFY snapshot Расхождение=Истина по деньгам

Аналог 19_verify_raskhozhdenie.py из knowledge_Balanse_klient.

Использование:
  python 06_verify_money_disp.py BASELINE   — сохранить текущее в _artifacts/06_baseline.json
  python 06_verify_money_disp.py VERIFY     — сравнить с baseline + diff

Артефакты:
  _artifacts/06_baseline.json
  _artifacts/06_current.json
  _artifacts/06_diff.md
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, get_refs, money, ARTIFACTS_DIR

erp = connect_erp()
refs = get_refs(erp)
S = erp.String

mode = sys.argv[1].upper() if len(sys.argv) > 1 else "VERIFY"
print(f"СКРИПТ 06 — Verify деньги (mode={mode})\n")

def snapshot():
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", refs["Орг"])
    q.УстановитьПараметр("ИБ", refs["Ист_Безнал"])
    q.УстановитьПараметр("ИН", refs["Ист_Налич"])
    q.УстановитьПараметр("ИП", refs["Ист_Подотч"])
    q.УстановитьПараметр("ИВ", refs["Ист_ВПути"])
    q.Текст = """
ВЫБРАТЬ
    Т.Регистратор.Месяц КАК Месяц,
    Т.Source КАК SourceRef,
    Т.Подразделение.Код КАК ПодрКод,
    Т.Статья.Код КАК СтКод,
    КОЛИЧЕСТВО(*) КАК Колво,
    СУММА(Т.СуммаКонечныйОстаток) КАК СуммаКМ
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Т
ГДЕ Т.Расхождение = ИСТИНА
    И Т.Организация = &Орг
    И Т.Source В (&ИБ, &ИН, &ИП, &ИВ)
    И Т.Регистратор.Месяц МЕЖДУ ДАТАВРЕМЯ(2025,12,1) И ДАТАВРЕМЯ(2026,12,31,23,59,59)
СГРУППИРОВАТЬ ПО Т.Регистратор.Месяц, Т.Source, Т.Подразделение.Код, Т.Статья.Код
"""
    r = q.Выполнить().Выгрузить()
    rows = []
    for i in range(r.Количество()):
        rec = r.Получить(i)
        mes = rec.Месяц.strftime("%Y-%m") if hasattr(rec.Месяц, "strftime") else str(rec.Месяц)[:7]
        src = str(S(rec.SourceRef)) if erp.ЗначениеЗаполнено(rec.SourceRef) else ""
        rows.append({
            "Месяц": mes, "Source": src,
            "ПодрКод": str(rec.ПодрКод or ""), "СтКод": str(rec.СтКод or ""),
            "Колво": int(rec.Колво or 0), "СуммаКМ": float(rec.СуммаКМ or 0),
        })
    return rows

current = snapshot()
baseline_path = os.path.join(ARTIFACTS_DIR, "06_baseline.json")
current_path = os.path.join(ARTIFACTS_DIR, "06_current.json")

with open(current_path, "w", encoding="utf-8") as f:
    json.dump(current, f, ensure_ascii=False, indent=2)
print(f"Сохранено текущее: {current_path}")
print(f"Записей Расхождение=Истина: {len(current)}")

# Сводка
by_ms = {}
for r in current:
    k = (r["Месяц"], r["Source"])
    by_ms.setdefault(k, {"Колво": 0, "ΣABS": 0.0, "ΣSign": 0.0})
    by_ms[k]["Колво"] += r["Колво"]
    by_ms[k]["ΣABS"] += abs(r["СуммаКМ"])
    by_ms[k]["ΣSign"] += r["СуммаКМ"]

print(f"\nСводка (Месяц × Source):")
print(f"{'Месяц':<10}{'Source':<40}{'Колво':>6}{'ΣABS':>18}{'ΣSign':>18}")
print("-" * 92)
ta = ts = 0; tk = 0
for k, v in sorted(by_ms.items()):
    print(f"{k[0]:<10}{k[1][:38]:<40}{v['Колво']:>6}{money(v['ΣABS']):>18}{money(v['ΣSign']):>18}")
    ta += v["ΣABS"]; ts += v["ΣSign"]; tk += v["Колво"]
print("-" * 92)
print(f"{'ИТОГО':<10}{'':<40}{tk:>6}{money(ta):>18}{money(ts):>18}")

if mode == "BASELINE":
    with open(baseline_path, "w", encoding="utf-8") as f:
        json.dump(current, f, ensure_ascii=False, indent=2)
    print(f"\n→ Сохранено как baseline: {baseline_path}")
    sys.exit(0)

if not os.path.exists(baseline_path):
    print(f"\n[!] Baseline не найден ({baseline_path}). Запусти сначала с BASELINE.")
    sys.exit(2)

with open(baseline_path, encoding="utf-8") as f:
    baseline = json.load(f)

bb = {}; cc = {}
for r in baseline:
    k = (r["Месяц"], r["Source"])
    bb[k] = bb.get(k, 0) + abs(r["СуммаКМ"])
for r in current:
    k = (r["Месяц"], r["Source"])
    cc[k] = cc.get(k, 0) + abs(r["СуммаКМ"])

print(f"\nBaseline: {len(baseline)} записей, Current: {len(current)}, Δзаписей: {len(current)-len(baseline)}")
print(f"\nИзменения (Месяц × Source) — ΣABS:")
print(f"{'Месяц':<10}{'Source':<40}{'Baseline':>16}{'Current':>16}{'Δ':>16}")
print("-" * 100)
for k in sorted(set(bb) | set(cc)):
    b = bb.get(k, 0); c = cc.get(k, 0)
    marker = "✓ улучш" if c < b else ("✗ ухудш" if c > b else "= б/изм")
    print(f"{k[0]:<10}{k[1][:38]:<40}{money(b):>16}{money(c):>16}{money(c-b):>16}  {marker}")
