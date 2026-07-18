# -*- coding: utf-8 -*-
"""
Auit and fix Месяц field on А_ФинРез_DDS documents.

Rule: Месяц = НачалоМесяца(Дата).

If current Месяц mismatches → unpost → update Месяц → re-post.
Duplicate-protection collision (>1 doc same month/org) — report and skip.
"""
import sys
from datetime import datetime
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")


def banner(t):
    print()
    print("=" * 100)
    print(f"  {t}")
    print("=" * 100)


def first_of_month(dt):
    return datetime(dt.year, dt.month, 1, 0, 0, 0)


v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


# === STEP 1: List all docs ===
banner("STEP 1: Усі документи Финансовый результат ДДС / Cashflow (для OLAP)")

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Д.Ссылка                 КАК Ссылка,
    Д.Номер                  КАК Номер,
    Д.Дата                   КАК Дата,
    Д.Месяц                  КАК Месяц,
    Д.Организация            КАК Организация,
    Д.Проведен               КАК Проведен,
    Д.ПометкаУдаления        КАК Удалено
ИЗ Документ.А_ФинРез_DDS КАК Д
УПОРЯДОЧИТЬ ПО Д.Дата
"""
res = q.Execute().Выгрузить()
docs = []
for i in range(res.Количество()):
    r = res.Получить(i)
    docs.append({
        "ref": r.Ссылка, "num": r.Номер, "date": r.Дата,
        "month": r.Месяц, "org": r.Организация,
        "posted": r.Проведен, "deleted": r.Удалено,
    })

print(f"Знайдено документів: {len(docs)}")
print(f"\n{'#':>3} | {'Номер':<12} | {'Дата':<11} | {'Месяц current':<13} | {'Месяц target':<13} | {'Проведен':<8} | {'Удалено':<7} | Дія")
print("-" * 120)
to_fix = []
for idx, d in enumerate(docs):
    target = first_of_month(d["date"])
    current = d["month"]
    mismatch = (current != target) if current else True
    action = "—"
    if d["deleted"]:
        action = "skip (видалений)"
    elif mismatch:
        action = "FIX"
        to_fix.append({**d, "target_month": target})
    print(f"{idx+1:>3} | {d['num'][:12]:<12} | {d['date'].strftime('%d.%m.%Y'):<11} | {current.strftime('%d.%m.%Y') if current else '(пусто)':<13} | {target.strftime('%d.%m.%Y'):<13} | {'Так' if d['posted'] else 'Ні':<8} | {'Так' if d['deleted'] else 'Ні':<7} | {action}")


# === STEP 2: Check duplicates after applying targets ===
banner("STEP 2: Перевірка дублікатів (Організація+Месяц) серед target_month")

dup_map = {}
for d in to_fix:
    # Plus existing posted docs that are NOT in to_fix list
    pass
all_after = []
for d in docs:
    if d["deleted"]:
        continue
    target = first_of_month(d["date"]) if d["month"] != first_of_month(d["date"]) else d["month"]
    all_after.append({**d, "after_month": target})

dup_check = {}
for d in all_after:
    key = (S(d["org"]), d["after_month"])
    dup_check.setdefault(key, []).append(d["num"])

dups = {k: v for k, v in dup_check.items() if len(v) > 1}
if dups:
    print("⚠ ЗНАЙДЕНО конфлікти (декілька документів в одному місяці):")
    for (org, m), nums in dups.items():
        print(f"  Організація={org}, Месяц={m.strftime('%d.%m.%Y')}: документи {', '.join(nums)}")
    print("\n→ Кожен конфлікт = тільки один документ зможе провестись (інші впадуть з 'не можна провести').")
    print("  Інші — або помічити на видалення, або змінити Дата на інший місяць перед фіксом.")
else:
    print("✓ Дублікатів не виявлено — кожен (Організація, Месяц) унікальний.")


# === STEP 3: Apply fixes ===
banner(f"STEP 3: Застосування виправлень ({len(to_fix)} документів)")

results = []
for d in to_fix:
    ref = d["ref"]
    target = d["target_month"]
    rec = {"num": d["num"], "old_month": d["month"], "new_month": target, "status": "—"}
    try:
        obj = ref.ПолучитьОбъект()
        if obj is None:
            rec["status"] = "FAIL: ПолучитьОбъект=None"
            results.append(rec)
            continue

        # Step A: unpost if posted
        was_posted = d["posted"]
        if was_posted:
            try:
                obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
            except Exception as e:
                msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
                rec["status"] = f"FAIL unpost: {msg[:80]}"
                results.append(rec)
                continue
            obj = ref.ПолучитьОбъект()  # re-acquire

        # Step B: update Месяц
        obj.Месяц = target
        try:
            obj.Записать()
        except Exception as e:
            msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
            rec["status"] = f"FAIL write: {msg[:80]}"
            results.append(rec)
            continue

        # Step C: re-post if was posted before
        if was_posted:
            try:
                obj = ref.ПолучитьОбъект()
                obj.Записать(erp.РежимЗаписиДокумента.Проведение)
                rec["status"] = "✓ оновлено + проведено"
            except Exception as e:
                msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
                rec["status"] = f"⚠ оновлено, але провести не вдалось: {msg[:80]}"
        else:
            rec["status"] = "✓ оновлено (не проведено, був вимкнений)"
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        rec["status"] = f"FAIL: {msg[:100]}"
    results.append(rec)


# === STEP 4: Report ===
banner("STEP 4: Підсумок")
print(f"{'Номер':<12} | {'Old Месяц':<11} | {'New Месяц':<11} | Status")
print("-" * 110)
ok = 0
fail = 0
for r in results:
    om = r["old_month"].strftime('%d.%m.%Y') if r["old_month"] else "(пусто)"
    nm = r["new_month"].strftime('%d.%m.%Y')
    status = r["status"]
    if status.startswith("✓"):
        ok += 1
    else:
        fail += 1
    print(f"{r['num'][:12]:<12} | {om:<11} | {nm:<11} | {status}")

print(f"\nΣ Успіх: {ok}, Провал: {fail}")


# === STEP 5: Verification — show final state ===
banner("STEP 5: Фінальний стан документів")

q5 = erp.NewObject("Запрос")
q5.Text = """
ВЫБРАТЬ Д.Номер, Д.Дата, Д.Месяц, Д.Проведен, Д.ПометкаУдаления
ИЗ Документ.А_ФинРез_DDS КАК Д
УПОРЯДОЧИТЬ ПО Д.Дата
"""
res5 = q5.Execute().Выгрузить()
print(f"{'#':>3} | {'Номер':<12} | {'Дата':<11} | {'Месяц':<11} | {'Проведен':<8} | {'Видалений':<10}")
print("-" * 80)
for i in range(res5.Количество()):
    r = res5.Получить(i)
    target = first_of_month(r.Дата)
    ok_mark = "✓" if r.Месяц == target else "✗"
    print(f"{i+1:>3} | {r.Номер[:12]:<12} | {r.Дата.strftime('%d.%m.%Y'):<11} | {r.Месяц.strftime('%d.%m.%Y') if r.Месяц else '(пусто)':<11} | {'Так' if r.Проведен else 'Ні':<8} | {'Так' if r.ПометкаУдаления else 'Ні':<10} | {ok_mark}")

print("\nDONE.")
