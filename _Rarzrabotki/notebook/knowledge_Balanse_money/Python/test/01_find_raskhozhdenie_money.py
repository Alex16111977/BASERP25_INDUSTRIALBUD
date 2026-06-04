# -*- coding: utf-8 -*-
"""
СКРИПТ 01 — Entry-point: Расхождение=Истина по деньгам в А_ОтчётБаланс_Свод

Источники: Безналичные / Наличные / У подотчётных / В пути.
Период: декабрь 2025 — апрель 2026.

Артефакты:
  _artifacts/01_summary.csv   — (Месяц × Source × Подр × Статья) суммы
  _artifacts/01_per_source.csv — pivot по источнику
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, get_refs, money, save_csv, ARTIFACTS_DIR

erp = connect_erp()
refs = get_refs(erp)
ORG = refs["Орг"]
S = erp.String

print("=" * 110)
print("СКРИПТ 01 — Расхождение=Истина по ДЕНЬГАМ (дек 2025 — апр 2026)")
print("=" * 110)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG)
q.УстановитьПараметр("ИБ", refs["Ист_Безнал"])
q.УстановитьПараметр("ИН", refs["Ист_Налич"])
q.УстановитьПараметр("ИП", refs["Ист_Подотч"])
q.УстановитьПараметр("ИВ", refs["Ист_ВПути"])
q.Текст = """
ВЫБРАТЬ
    Т.Регистратор.Месяц КАК Месяц,
    Т.Source КАК Source,
    Т.Подразделение.Код КАК ПодрКод,
    ПРЕДСТАВЛЕНИЕ(Т.Подразделение) КАК Подр,
    Т.Статья.Код КАК СтКод,
    ПРЕДСТАВЛЕНИЕ(Т.Статья) КАК Статья,
    КОЛИЧЕСТВО(*) КАК Колво,
    СУММА(Т.СуммаНачальныйОстаток) КАК НО,
    СУММА(Т.СуммаКонечныйОстаток) КАК КМ
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Т
ГДЕ Т.Организация = &Орг
    И Т.Расхождение = ИСТИНА
    И Т.Регистратор.Месяц МЕЖДУ ДАТАВРЕМЯ(2025,12,1) И ДАТАВРЕМЯ(2026,12,31,23,59,59)
    И Т.Source В (&ИБ, &ИН, &ИП, &ИВ)
СГРУППИРОВАТЬ ПО
    Т.Регистратор.Месяц, Т.Source,
    Т.Подразделение.Код, Т.Подразделение,
    Т.Статья.Код, Т.Статья
УПОРЯДОЧИТЬ ПО Месяц, Source, ПодрКод
"""
r = q.Выполнить().Выгрузить()
print(f"\nВсего групп Расхождение=Истина: {r.Количество()}")

# Сводка по (Месяц × Source)
by_ms = {}
rows = []
for i in range(r.Количество()):
    rec = r.Получить(i)
    mes = str(rec.Месяц)[:10] if rec.Месяц else ""
    src = str(S(rec.Source)) if erp.ЗначениеЗаполнено(rec.Source) else "(пусто)"
    podr = str(rec.Подр or "(пусто)")
    rows.append({
        "Месяц": mes, "Source": src,
        "ПодрКод": str(rec.ПодрКод or ""), "Подр": podr,
        "СтКод": str(rec.СтКод or ""), "Статья": str(rec.Статья or ""),
        "Колво": int(rec.Колво or 0),
        "НО": float(rec.НО or 0),
        "КМ": float(rec.КМ or 0),
    })
    k = (mes, src)
    by_ms.setdefault(k, {"кол": 0, "ΣABS_КМ": 0.0, "ΣКМ": 0.0})
    by_ms[k]["кол"] += int(rec.Колво or 0)
    by_ms[k]["ΣABS_КМ"] += abs(float(rec.КМ or 0))
    by_ms[k]["ΣКМ"] += float(rec.КМ or 0)

print(f"\nСводка (Месяц × Source):")
print(f"{'Месяц':<12}{'Source':<40}{'Колво':>6}{'ΣABS_КМ':>18}{'ΣКМ':>18}")
print("-" * 94)
total = 0.0
for k, v in sorted(by_ms.items()):
    print(f"{k[0]:<12}{k[1][:38]:<40}{v['кол']:>6}{money(v['ΣABS_КМ']):>18}{money(v['ΣКМ']):>18}")
    total += v["ΣABS_КМ"]
print(f"\nИТОГО ΣABS_КМ: {money(total)}")

save_csv("01_summary", rows, ["Месяц", "Source", "ПодрКод", "Подр", "СтКод", "Статья", "Колво", "НО", "КМ"])

# Pivot по Source
print(f"\n--- Pivot по источнику (агрегат за весь период) ---")
by_src = {}
for r in rows:
    by_src.setdefault(r["Source"], {"кол": 0, "ΣABS": 0.0, "ΣКМ": 0.0})
    by_src[r["Source"]]["кол"] += r["Колво"]
    by_src[r["Source"]]["ΣABS"] += abs(r["КМ"])
    by_src[r["Source"]]["ΣКМ"] += r["КМ"]
print(f"{'Source':<45}{'Колво':>6}{'ΣABS':>18}{'ΣКМ':>18}")
print("-" * 87)
for src, v in sorted(by_src.items(), key=lambda x: -x[1]["ΣABS"]):
    print(f"{src[:43]:<45}{v['кол']:>6}{money(v['ΣABS']):>18}{money(v['ΣКМ']):>18}")

src_rows = [{"Source": k, "Колво": v["кол"], "ΣABS": v["ΣABS"], "ΣКМ": v["ΣКМ"]} for k, v in by_src.items()]
save_csv("01_per_source", src_rows, ["Source", "Колво", "ΣABS", "ΣКМ"])

print(f"\nАртефакты: 01_summary.csv, 01_per_source.csv")
