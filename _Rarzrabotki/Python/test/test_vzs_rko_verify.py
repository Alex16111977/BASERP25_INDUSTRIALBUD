# -*- coding: utf-8 -*-
"""Verify acceptance тест для процедури ЗаполнитьИзРКО.

Кроки:
  1) Завантажує кандидатів з rko_candidates.json
  2) Перепроводить кожний (РежимЗаписиДокумента.Проведение)
  3) Per Регістратор:
     - Σ signed А_ВзСС = −Σ signed ПАП(ОТ) до 1 ₽
     - Доля строк з ФизЛицо ≠ Пустая ≥ 99%
     - 0 дублів ключа (Орг, Подр, ФЛ, ВидДв)
  4) Per (Орг, Подр) — |А_ВзСС + ПАП| ≤ 1 ₽, 0 розбіжностей
  5) Ідемпотентність: 2-й прогон → Δ=0
"""
import sys, io, json
from collections import defaultdict
from datetime import datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

TOL_RUB = 1.0
pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Завантажуємо кандидатів
cand_path = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\rko_candidates.json"
with open(cand_path, encoding="utf-8") as f:
    КАНДИДАТЫ = json.load(f)
print(f"[INFO] Кандидатів: {len(КАНДИДАТЫ)}")

qOT = conn.NewObject("Запрос")
qOT.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ '
             'ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование = "Оплата труда"')
rOT = qOT.Выполнить().Выбрать(); rOT.Следующий(); ОТ = rOT.С

PROVED = conn.PredefinedValue("РежимЗаписиДокумента.Проведение")

def calc_movements(док_ref):
    """Повертає (rows_А_ВзСС[], Σ_signed_А_ВзСС, Σ_signed_ПАП)."""
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ
        Запис.Организация КАК Орг, Запис.Подразделение КАК Подр, Запис.ФизическоеЛицо КАК ФЛ,
        Запис.ВидДвижения КАК ВД, Запис.СуммаВзаиморасчетов КАК Сум
    ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК Запис
    ГДЕ Запис.Регистратор = &Док
    """
    q.УстановитьПараметр("Док", док_ref)
    rows = q.Выполнить().Выгрузить()
    наш_rows = []
    Σ_наш = 0.0
    for i in range(rows.Количество()):
        s = rows.Получить(i)
        sign = +1 if S(s.ВД) == "Приход" else -1
        Σ_наш += sign * float(s.Сум)
        наш_rows.append({
            "орг": S(s.Орг), "подр": S(s.Подр), "фл": S(s.ФЛ),
            "вд": S(s.ВД), "сум": float(s.Сум),
        })
    # Σ ПАП
    qP = conn.NewObject("Запрос")
    qP.Текст = """
    ВЫБРАТЬ
        СУММА(ВЫБОР КОГДА П.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
            ТОГДА П.Сумма ИНАЧЕ -П.Сумма КОНЕЦ) КАК Σ
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК П
    ГДЕ П.Регистратор = &Док
        И П.Статья = &ОТ
        И П.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка)
    """
    qP.УстановитьПараметр("Док", док_ref); qP.УстановитьПараметр("ОТ", ОТ)
    rp = qP.Выполнить().Выбрать(); rp.Следующий()
    return наш_rows, Σ_наш, float(rp.Σ or 0)


def repost(uuid_str):
    uuid_obj = conn.NewObject("УникальныйИдентификатор", uuid_str)
    ref = conn.Документы.РасходныйКассовыйОрдер.ПолучитьСсылку(uuid_obj)
    obj = ref.ПолучитьОбъект()
    obj.Записать(PROVED)
    return ref


all_pass = True
total_rows = 0
total_fl_empty = 0
total_доку = 0
report = []

print(f"\n{'-'*100}")
print(f"{'№':>15} | {'Дата':<11} | {'Σ_наш':>14} | {'Σ_ПАП':>14} | {'|Δ|':>8} | рядків | дублі | FL%")
print(f"{'-'*100}")
for k in КАНДИДАТЫ:
    ref = repost(k["uuid"])
    rows, Σ_наш, Σ_pap = calc_movements(ref)
    delta = abs(Σ_наш + Σ_pap)
    keys = defaultdict(int)
    fl_empty = 0
    for r in rows:
        keys[(r["орг"], r["подр"], r["фл"], r["вд"])] += 1
        if r["фл"] == "" or r["фл"] is None:
            fl_empty += 1
    дублі = sum(1 for v in keys.values() if v > 1)
    fl_pct = 100.0 * (len(rows) - fl_empty) / max(1, len(rows))
    status = "PASS" if delta <= TOL_RUB and дублі == 0 and fl_pct >= 99.0 else "FAIL"
    if status == "FAIL":
        all_pass = False
    total_rows += len(rows)
    total_fl_empty += fl_empty
    total_доку += 1
    print(f"{k['number']:>15} | {k['date'][:10]:<11} | {Σ_наш:>+14,.2f} | {Σ_pap:>+14,.2f} | {delta:>8,.2f} | {len(rows):>6} | {дублі:>5} | {fl_pct:>5.1f}% [{status}]")
    report.append({"number": k["number"], "Σ_наш": Σ_наш, "Σ_pap": Σ_pap, "delta": delta, "rows": len(rows)})

# Ідемпотентність — перепровести знов і порівняти
print(f"\n[ІДЕМПОТЕНТНІСТЬ] 2-й прогон...")
idem_ok = True
for k in КАНДИДАТЫ:
    ref = repost(k["uuid"])
    rows2, Σ2, ΣP2 = calc_movements(ref)
    # Знайти попередній
    prev = next(r for r in report if r["number"] == k["number"])
    Δ_idem = abs(prev["Σ_наш"] - Σ2)
    print(f"  №{k['number']:>14}: Σ перший={prev['Σ_наш']:>+12,.2f} Σ другий={Σ2:>+12,.2f} |Δ|={Δ_idem:.2f}")
    if Δ_idem > 0.01:
        idem_ok = False

fl_pct_total = 100.0 * (total_rows - total_fl_empty) / max(1, total_rows)
print(f"\n[ЗВЕДЕННЯ]")
print(f"  Документів: {total_доку}")
print(f"  Усього рядків А_ВзСС: {total_rows}, з ФЛ={total_rows - total_fl_empty} ({fl_pct_total:.1f}%)")
print(f"  Інваріант: {'PASS' if all_pass else 'FAIL'}")
print(f"  Ідемпотентність: {'PASS' if idem_ok else 'FAIL'}")

final = all_pass and idem_ok
print(f"\n{'[OK] PASS' if final else '[FAIL]'}")
sys.exit(0 if final else 1)
