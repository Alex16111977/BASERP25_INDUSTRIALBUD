# -*- coding: utf-8 -*-
"""
Test: why ВзаимозачетЗадолженности ІБ00-000084 doesn't fill
ПодразделениеПриход and ПодразделениеРасход in the analysis table.

ІБ00-000084 is a manual vzaimozachet (no А_ДокументОснованиеДляБаланса).
It creates movements in ПрочиеАктивыПассивы between departments.
The mass processor should detect the disbalance and show departments.
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    try:
        if obj is None: return "<None>"
        if not conn.ValueIsFilled(obj): return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

# 1. Find IB00-000084 by sum (unique 244812.62)
print("=" * 70)
print("1. Find IB00-000084")
print("=" * 70)

q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка, Док.Номер КАК Н, Док.СуммаРегл, Док.Проведен "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.СуммаРегл = 244812.62 "
    "   И Док.А_ДокументОснованиеДляБаланса = НЕОПРЕДЕЛЕНО"
)
res = q.Execute().Select()
ref_084 = None
while res.Next():
    ref_084 = res.Ссылка
    print(f"  {s(res.Н).strip()}: sum={res.СуммаРегл} posted={res.Проведен}")

if not ref_084:
    print("  NOT FOUND!"); exit()

# 2. Check movements in ПрочиеАктивыПассивы
print("\n" + "=" * 70)
print("2. Movements in register (Обороты)")
print("=" * 70)

q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ "
    "   Об.Подразделение КАК П, "
    "   Об.Организация КАК О, "
    "   Об.Статья КАК Ст, "
    "   Об.СуммаПриход КАК Д, "
    "   Об.СуммаРасход КАК К, "
    "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
    "ГДЕ Об.Регистратор = &Р "
    "   И НЕ Об.Статья В ("
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),"
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"
)
q2.SetParameter("Р", ref_084)
r2 = q2.Execute().Select()

depts = {}
total_ctrl = 0
row_count = 0
while r2.Next():
    row_count += 1
    dept = s(r2.П)
    ctrl = round(float(r2.Контроль), 2)
    total_ctrl += ctrl
    print(f"  {dept}: Dt={r2.Д:.2f} Kt={r2.К:.2f} Ctrl={ctrl} ({s(r2.Ст)})")
    if dept not in depts:
        depts[dept] = 0
    depts[dept] += ctrl

print(f"\n  Rows: {row_count}")
print(f"  Total Ctrl: {round(total_ctrl, 2)}")
print(f"  Departments: {depts}")

# 3. Apply same logic as the processor
print("\n" + "=" * 70)
print("3. Processor logic simulation")
print("=" * 70)

# Check: org balance = 0?
org_ok = abs(round(total_ctrl, 2)) == 0
print(f"  Org balance OK (sum Ctrl = 0): {org_ok}")

# Check: min 2 departments with Ctrl <> 0
disb_count = sum(1 for v in depts.values() if round(v, 2) != 0)
print(f"  Departments with disbalance: {disb_count}")

# Find debtor (max Ctrl > 0) and creditor (min Ctrl < 0)
debtor = None
creditor = None
max_k = 0
min_k = 0
for dept, ctrl in depts.items():
    if ctrl > max_k:
        debtor = dept
        max_k = ctrl
    if ctrl < min_k:
        creditor = dept
        min_k = ctrl

print(f"  Debtor (Ctrl>0): {debtor} ({max_k})")
print(f"  Creditor (Ctrl<0): {creditor} ({min_k})")

if debtor == creditor:
    print(f"  *** PROBLEM: debtor == creditor! Same department! ***")

if debtor is None:
    print(f"  *** PROBLEM: no debtor (no Ctrl > 0)! ***")
if creditor is None:
    print(f"  *** PROBLEM: no creditor (no Ctrl < 0)! ***")

# 4. Check the query filter
print("\n" + "=" * 70)
print("4. Query filter check")
print("=" * 70)

# The mass processor uses:
# И НЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности
# This EXCLUDES all vzaimozachety! Including IB00-000084.
print(f"  IB00-000084 is ВзаимозачетЗадолженности: YES")
print(f"  Query filter: И НЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности")
print(f"  => IB00-000084 is EXCLUDED from analysis!")
print(f"  => That's why ПодразделениеПриход/Расход are empty!")
print(f"")
print(f"  The document APPEARS in the table only via the 'stale vzaimozachety'")
print(f"  block (section 8), which does NOT fill ПодразделениеПриход/Расход!")

# 5. Verify: check the stale block
print("\n" + "=" * 70)
print("5. How IB00-000084 appears in table (stale block)")
print("=" * 70)

# The stale block queries vzaimozachety with А_ДокументОснованиеДляБаланса filled
# IB00-000084 has А_ДокументОснованиеДляБаланса = НЕОПРЕДЕЛЕНО
# So it should NOT appear via stale block either!

# Actually, IB00-000084 appears because it's NOT excluded by the main query
# (we reverted to simple filter). Wait - we DID revert to exclude ALL vzaimozachety.
# So IB00-000084 should NOT appear at all!

# Unless it appears from somewhere else. Let me check the current filter.
print(f"  А_ДокументОснованиеДляБаланса: {s(ref_084.А_ДокументОснованиеДляБаланса)}")
print(f"  => If EMPTY, stale block also skips it (requires <> НЕОПРЕДЕЛЕНО)")
print(f"")
print(f"  CONCLUSION: IB00-000084 should NOT appear in the table at all!")
print(f"  If it does appear - it's from the main query, which means the filter")
print(f"  was changed. The stale block fills rows WITHOUT departments.")
print(f"")
print(f"  FIX: The stale block should also fill ПодразделениеПриход/Расход")
print(f"  by querying movements of the source document.")
