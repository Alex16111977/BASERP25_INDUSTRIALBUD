# -*- coding: utf-8 -*-
"""Диагностика per-документ декабрь 2025 (Безнал + Налич).

Цель: найти КАКИЕ ДОКУМЕНТЫ ломают баланс — где ПАП.Подр ≠ Регистр.Подр
или |ПАП.Σ - Регистр.Σ| > 0.01

Алгоритм:
1. ПАП.Безнал+Налич за дек25 per (Регистратор, Подр, Источник) → ΣПАП
2. РН.Безнал per (Регистратор, Подр, БС) → ΣРег
3. РН.Налич per (Регистратор, Подр, Касса) → ΣРег
4. JOIN по Регистратору → классификация нарушений (ASYM_PODR / DIFF_SUMM / MISSING_*)
5. CSV-выгрузка топ-нарушителей
"""
from _common import (
    connect_erp, get_refs, money, save_csv, get_uuid, get_type_name, ARTIFACTS_DIR
)
import os
from collections import defaultdict

erp = connect_erp()
S = erp.String
refs = get_refs(erp)
print(f"Org: {S(refs['Орг'])}")
print(f"Ист_Безнал: {S(refs['Ист_Безнал'])}")
print(f"Ист_Налич:  {S(refs['Ист_Налич'])}")

# === ШАГ 1 — ПАП per (Регистратор, Подр, Источник) ===
print("\n[1/4] Запрос ПАП.{Безнал,Налич} за декабрь 2025...")
qp = erp.NewObject("Запрос")
qp.УстановитьПараметр("Орг", refs["Орг"])
qp.УстановитьПараметр("ИБ", refs["Ист_Безнал"])
qp.УстановитьПараметр("ИН", refs["Ист_Налич"])
qp.Текст = """
ВЫБРАТЬ
    Т.Регистратор КАК Регистратор,
    ЕСТЬNULL(Т.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)) КАК Подр,
    Т.Источник КАК Источник,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК ΣПАП
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Организация = &Орг
    И Т.Период >= ДАТАВРЕМЯ(2025,12,1,0,0,0)
    И Т.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
    И Т.Источник В (&ИБ, &ИН)
СГРУППИРОВАТЬ ПО
    Т.Регистратор,
    ЕСТЬNULL(Т.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)),
    Т.Источник
"""
rp = qp.Выполнить().Выгрузить()
print(f"  ПАП строк: {rp.Количество()}")

# pap_by_reg[(reg_uuid, ист_имя)] = {"reg_ref": ..., "ист_имя": ..., "rows": [(podr_uuid, podr_name, summ)...]}
pap_by_reg = defaultdict(lambda: {"reg_ref": None, "ист": None, "rows": []})
ист_безнал_str = S(refs["Ист_Безнал"])
ист_налич_str = S(refs["Ист_Налич"])
for i in range(rp.Количество()):
    rec = rp.Получить(i)
    reg_uuid = get_uuid(erp, rec.Регистратор)
    ист_str = S(rec.Источник)
    ист_name = "Безнал" if ист_str == ист_безнал_str else "Налич"
    podr_uuid = get_uuid(erp, rec.Подр)
    podr_name = S(rec.Подр) if erp.ЗначениеЗаполнено(rec.Подр) else "(пусто)"
    key = (reg_uuid, ист_name)
    pap_by_reg[key]["reg_ref"] = rec.Регистратор
    pap_by_reg[key]["ист"] = ист_name
    pap_by_reg[key]["rows"].append((podr_uuid, podr_name, float(rec.ΣПАП or 0)))

# === ШАГ 2 — РН.Безнал per (Регистратор, Подр, БС) ===
print("\n[2/4] Запрос РегистрНакопления.ДенежныеСредстваБезналичные за декабрь 2025...")
qb = erp.NewObject("Запрос")
qb.УстановитьПараметр("Орг", refs["Орг"])
qb.Текст = """
ВЫБРАТЬ
    Б.Регистратор КАК Регистратор,
    ЕСТЬNULL(Б.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)) КАК Подр,
    Б.БанковскийСчет КАК Объект,
    СУММА(ВЫБОР КОГДА Б.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА Б.Сумма ИНАЧЕ -Б.Сумма КОНЕЦ) КАК ΣРег
ИЗ РегистрНакопления.ДенежныеСредстваБезналичные КАК Б
ГДЕ Б.Организация = &Орг
    И Б.Период >= ДАТАВРЕМЯ(2025,12,1,0,0,0)
    И Б.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
СГРУППИРОВАТЬ ПО Б.Регистратор,
    ЕСТЬNULL(Б.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)),
    Б.БанковскийСчет
"""
rb = qb.Выполнить().Выгрузить()
print(f"  Безнал строк: {rb.Количество()}")

reg_by_reg = defaultdict(lambda: {"reg_ref": None, "ист": None, "rows": []})
for i in range(rb.Количество()):
    rec = rb.Получить(i)
    reg_uuid = get_uuid(erp, rec.Регистратор)
    podr_uuid = get_uuid(erp, rec.Подр)
    podr_name = S(rec.Подр) if erp.ЗначениеЗаполнено(rec.Подр) else "(пусто)"
    obj_name = S(rec.Объект) if erp.ЗначениеЗаполнено(rec.Объект) else "(пусто)"
    key = (reg_uuid, "Безнал")
    reg_by_reg[key]["reg_ref"] = rec.Регистратор
    reg_by_reg[key]["ист"] = "Безнал"
    reg_by_reg[key]["rows"].append((podr_uuid, podr_name, obj_name, float(rec.ΣРег or 0)))

# === ШАГ 3 — РН.Налич per (Регистратор, Подр, Касса) ===
print("\n[3/4] Запрос РегистрНакопления.ДенежныеСредстваНаличные за декабрь 2025...")
qn = erp.NewObject("Запрос")
qn.УстановитьПараметр("Орг", refs["Орг"])
qn.Текст = """
ВЫБРАТЬ
    Н.Регистратор КАК Регистратор,
    ЕСТЬNULL(Н.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)) КАК Подр,
    Н.Касса КАК Объект,
    СУММА(ВЫБОР КОГДА Н.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА Н.Сумма ИНАЧЕ -Н.Сумма КОНЕЦ) КАК ΣРег
ИЗ РегистрНакопления.ДенежныеСредстваНаличные КАК Н
ГДЕ Н.Организация = &Орг
    И Н.Период >= ДАТАВРЕМЯ(2025,12,1,0,0,0)
    И Н.Период <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
СГРУППИРОВАТЬ ПО Н.Регистратор,
    ЕСТЬNULL(Н.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)),
    Н.Касса
"""
rn = qn.Выполнить().Выгрузить()
print(f"  Налич строк: {rn.Количество()}")

for i in range(rn.Количество()):
    rec = rn.Получить(i)
    reg_uuid = get_uuid(erp, rec.Регистратор)
    podr_uuid = get_uuid(erp, rec.Подр)
    podr_name = S(rec.Подр) if erp.ЗначениеЗаполнено(rec.Подр) else "(пусто)"
    obj_name = S(rec.Объект) if erp.ЗначениеЗаполнено(rec.Объект) else "(пусто)"
    key = (reg_uuid, "Налич")
    reg_by_reg[key]["reg_ref"] = rec.Регистратор
    reg_by_reg[key]["ист"] = "Налич"
    reg_by_reg[key]["rows"].append((podr_uuid, podr_name, obj_name, float(rec.ΣРег or 0)))

# === ШАГ 4 — JOIN + классификация ===
print("\n[4/4] Классификация...")
all_keys = set(pap_by_reg.keys()) | set(reg_by_reg.keys())
findings = []
for key in all_keys:
    reg_uuid, ист = key
    pap = pap_by_reg.get(key)
    reg = reg_by_reg.get(key)

    if pap and not reg:
        тип = "MISSING_REG"
        σпап = sum(r[2] for r in pap["rows"])
        σрег = 0.0
        pap_podrs = sorted(set(r[1] for r in pap["rows"]))
        reg_podrs = []
        objects = []
        ref = pap["reg_ref"]
    elif reg and not pap:
        тип = "MISSING_PAP"
        σпап = 0.0
        σрег = sum(r[3] for r in reg["rows"])
        pap_podrs = []
        reg_podrs = sorted(set(r[1] for r in reg["rows"]))
        objects = sorted(set(r[2] for r in reg["rows"]))
        ref = reg["reg_ref"]
    else:
        σпап = sum(r[2] for r in pap["rows"])
        σрег = sum(r[3] for r in reg["rows"])
        pap_podrs = sorted(set(r[1] for r in pap["rows"]))
        reg_podrs = sorted(set(r[1] for r in reg["rows"]))
        objects = sorted(set(r[2] for r in reg["rows"]))
        ref = pap["reg_ref"]
        if abs(σпап - σрег) > 0.01:
            тип = "DIFF_SUMM"
        elif set(pap_podrs) != set(reg_podrs):
            тип = "ASYM_PODR"
        else:
            тип = "OK"

    if тип == "OK":
        continue

    дата = ""
    номер = ""
    тип_док = ""
    try:
        дата = str(ref.Дата)[:19] if ref else ""
        номер = str(ref.Номер) if ref else ""
        тип_док = get_type_name(erp, ref)
    except Exception:
        pass

    findings.append({
        "Регистратор": S(ref) if ref else "",
        "ТипДок": тип_док,
        "Дата": дата,
        "Номер": номер,
        "Источник": ист,
        "ПАП_Подр": " | ".join(pap_podrs),
        "Регистр_Подр": " | ".join(reg_podrs),
        "БС_Касса": " | ".join(objects),
        "ПАП_Σ": σпап,
        "Регистр_Σ": σрег,
        "ΔΣ": σпап - σрег,
        "ТипНарушения": тип,
    })

# Сводка
print(f"\n=== Сводка ===")
print(f"Всего регистраторов нарушителей: {len(findings)}")
by_type = defaultdict(list)
for f in findings:
    by_type[f["ТипНарушения"]].append(f)
for t, items in by_type.items():
    σ = sum(x["ΔΣ"] if t == "DIFF_SUMM" else x["ПАП_Σ"] for x in items)
    label = "Σ ΔΣ" if t == "DIFF_SUMM" else "Σ ΣПАП"
    print(f"  {t:<12}: {len(items):>4}  ({label} = {money(σ)})")

# Топ-10 по |ΔΣ| или по |ΣПАП| для ASYM_PODR
findings.sort(key=lambda x: -abs(x["ΔΣ"]) if x["ТипНарушения"] == "DIFF_SUMM" else -abs(x["ПАП_Σ"]))
print(f"\n=== Топ-10 по |ΔΣ| / |ΣПАП| ===")
for f in findings[:10]:
    print(f"  [{f['ТипНарушения']:<11}] {f['Дата']:<19}  {f['ТипДок']:<35}  Ном={f['Номер']:<14}  "
          f"ΣПАП={money(f['ПАП_Σ']):>16}  ΣРег={money(f['Регистр_Σ']):>16}  ΔΣ={money(f['ΔΣ']):>14}")
    print(f"    ПАП.Подр: {f['ПАП_Подр']}")
    print(f"    Рег.Подр: {f['Регистр_Подр']}")
    print(f"    БС/Касса: {f['БС_Касса']}")

# CSV (форматируем числа)
csv_rows = []
for f in findings:
    csv_rows.append({
        "Регистратор": f["Регистратор"],
        "ТипДок": f["ТипДок"],
        "Дата": f["Дата"],
        "Номер": f["Номер"],
        "Источник": f["Источник"],
        "ПАП_Подр": f["ПАП_Подр"],
        "Регистр_Подр": f["Регистр_Подр"],
        "БС_Касса": f["БС_Касса"],
        "ПАП_Σ": money(f["ПАП_Σ"]),
        "Регистр_Σ": money(f["Регистр_Σ"]),
        "ΔΣ": money(f["ΔΣ"]),
        "ТипНарушения": f["ТипНарушения"],
    })

headers = ["Регистратор", "ТипДок", "Дата", "Номер", "Источник",
           "ПАП_Подр", "Регистр_Подр", "БС_Касса",
           "ПАП_Σ", "Регистр_Σ", "ΔΣ", "ТипНарушения"]
out_path = save_csv("diag_dec25_money_per_doc", csv_rows, headers)
print(f"\n[OK] CSV: {out_path}")
print(f"[OK] {len(csv_rows)} строк сохранено")
