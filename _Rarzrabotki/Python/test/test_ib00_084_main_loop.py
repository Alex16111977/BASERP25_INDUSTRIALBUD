# -*- coding: utf-8 -*-
"""
Test: simulate main loop logic for IB00-000084
to find why ПодразделениеПриход/Расход are empty
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    try:
        if obj is None: return "<None>"
        if not conn.ValueIsFilled(obj): return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

# Find IB00-000084
q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.СуммаРегл = 244812.62 "
    "   И Док.А_ДокументОснованиеДляБаланса = НЕОПРЕДЕЛЕНО"
)
res = q.Execute().Select()
res.Next()
ref_084 = res.Ссылка
print(f"Doc: {s(ref_084)}")

# Simulate the MAIN query (with filter that includes vzaimozachety)
q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ "
    "   Об.Регистратор КАК Документ, "
    "   Об.Подразделение КАК Подразделение, "
    "   Об.Организация КАК Организация, "
    "   Об.СуммаПриход КАК Дебет, "
    "   Об.СуммаРасход КАК Кредит, "
    "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты("
    "   &Нач, &Кон, Регистратор, ) КАК Об "
    "ГДЕ НЕ Об.Статья В ("
    "   ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),"
    "   ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств)) "
    "   И Об.Регистратор = &Рег"
)
q2.SetParameter("Нач", datetime.datetime(2025, 12, 1))
q2.SetParameter("Кон", datetime.datetime(2025, 12, 31, 23, 59, 59))
q2.SetParameter("Рег", ref_084)

r2 = q2.Execute().Select()

print(f"\nMovements in query:")
rows = []
while r2.Next():
    dept = s(r2.Подразделение)
    ctrl = round(float(r2.Контроль), 2)
    org = r2.Организация
    print(f"  {dept}: Dt={r2.Дебет:.2f} Kt={r2.Кредит:.2f} Ctrl={ctrl}")
    rows.append({"dept": dept, "ctrl": ctrl, "dept_ref": r2.Подразделение, "org": org})

# Simulate grouping (like the processor does)
print(f"\n--- Processor logic ---")
total_ctrl = sum(r["ctrl"] for r in rows)
print(f"1. Total Ctrl (org balance): {round(total_ctrl, 2)}")
org_ok = round(total_ctrl, 2) == 0
print(f"   OK: {org_ok}")

if not org_ok:
    print(f"   => SKIPPED (org balance != 0)")
    exit()

# Count depts with disbalance
disb = [r for r in rows if round(r["ctrl"], 2) != 0]
print(f"2. Departments with disbalance: {len(disb)}")
if len(disb) < 2:
    print(f"   => SKIPPED (< 2 departments)")
    exit()

# Find debtor/creditor
max_k = 0
min_k = 0
debtor = None
creditor = None
for r in rows:
    if r["ctrl"] > max_k:
        debtor = r["dept"]
        max_k = r["ctrl"]
    if r["ctrl"] < min_k:
        creditor = r["dept"]
        min_k = r["ctrl"]

print(f"3. Debtor: {debtor} ({max_k})")
print(f"   Creditor: {creditor} ({min_k})")

if debtor is None or creditor is None:
    print(f"   => SKIPPED (no debtor or creditor)")
    exit()

if debtor == creditor:
    print(f"   => SKIPPED (debtor == creditor)")
    exit()

# Department filter (КМД-2)
print(f"4. Department filter: КМД-2")
print(f"   Debtor matches: {'КМД-2' in debtor}")
print(f"   Creditor matches: {'КМД-2' in creditor}")
if 'КМД-2' not in debtor and 'КМД-2' not in creditor:
    print(f"   => SKIPPED by department filter!")
else:
    print(f"   => PASSES filter")

print(f"\n5. Result:")
print(f"   ПодразделениеПриход = {debtor}")
print(f"   ПодразделениеРасход = {creditor}")
print(f"   СуммаДокумента = {max_k}")

# Check: does the query even include this doc?
# If the filter is И НЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности
# then IB00-000084 is EXCLUDED and never reaches the grouping step
print(f"\n6. KEY CHECK: Is doc excluded by vzaimozachet filter?")
print(f"   Doc is ВзаимозачетЗадолженности: YES")
print(f"   If filter 'И НЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности':")
print(f"   => Doc IS EXCLUDED from main query!")
print(f"   => Departments are NEVER determined!")
print(f"   => Row appears from STALE block (without departments)")
