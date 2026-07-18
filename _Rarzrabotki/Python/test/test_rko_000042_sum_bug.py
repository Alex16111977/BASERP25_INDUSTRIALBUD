# -*- coding: utf-8 -*-
"""
Test: why vzaimozachet for РКО 000Ц-000042 is 0.01 instead of 0.11

Movements:
  Строительство: +0.01 (Задолженность) + +0.10 (Выданные авансы) = +0.11
  Спецтехника: -0.11 (ДС наличные)

The Обороты virtual table returns separate rows per Статья.
The processor takes individual row values instead of SUM per department.
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    try:
        if not conn.ValueIsFilled(obj): return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

# Find the document
q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Д.Ссылка ИЗ Документ.РасходныйКассовыйОрдер КАК Д "
    "ГДЕ Д.Номер ПОДОБНО \"%000042%\" "
    "   И Д.Дата МЕЖДУ &Н И &К"
)
q.SetParameter("Н", datetime.datetime(2026, 1, 28))
q.SetParameter("К", datetime.datetime(2026, 1, 28, 23, 59, 59))
res = q.Execute().Select()
if not res.Next():
    print("Doc not found!"); exit()
doc_ref = res.Ссылка
print(f"Doc: {s(doc_ref)}")

# Query Обороты (same as processor)
print("\n=== Raw Обороты rows (as processor sees them) ===")
q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ "
    "   Об.Подразделение КАК П, Об.Статья КАК Ст, "
    "   Об.СуммаПриход КАК Д, Об.СуммаРасход КАК К, "
    "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
    "ГДЕ Об.Регистратор = &Р "
    "   И НЕ Об.Статья В ("
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),"
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"
)
q2.SetParameter("Р", doc_ref)
r2 = q2.Execute().Select()

rows = []
while r2.Next():
    dept = s(r2.П)
    ctrl = round(float(r2.Контроль), 2)
    article = s(r2.Ст)
    print(f"  {dept}: Ctrl={ctrl:+.2f} ({article})")
    rows.append({"dept": dept, "ctrl": ctrl})

# Simulate CURRENT processor logic (bug)
print("\n=== CURRENT logic (А_ДвижениеПоВзаимозачетПоБалансуУпр) ===")
print("Takes LAST row per side, NOT sum per dept:")
debtor = None
creditor = None
sum_vz = 0
for r in rows:
    if r["ctrl"] > 0:
        debtor = r["dept"]
        sum_vz = r["ctrl"]  # Overwrites each time!
        print(f"  Ctrl>0: {r['dept']} = {r['ctrl']} -> СуммаВзаимозачета = {r['ctrl']}")
    elif r["ctrl"] < 0:
        creditor = r["dept"]

print(f"\n  RESULT: debtor={debtor}, creditor={creditor}")
print(f"  СуммаВзаимозачета = {sum_vz}")
print(f"  BUG: should be sum of ALL positive rows for debtor dept!")

# Correct logic
print("\n=== CORRECT logic (sum per department) ===")
dept_sums = {}
for r in rows:
    dept_sums[r["dept"]] = dept_sums.get(r["dept"], 0) + r["ctrl"]

print(f"  Department sums: {dept_sums}")

correct_debtor = max(dept_sums, key=dept_sums.get)
correct_creditor = min(dept_sums, key=dept_sums.get)
correct_sum = dept_sums[correct_debtor]

print(f"  Debtor: {correct_debtor} ({dept_sums[correct_debtor]:+.2f})")
print(f"  Creditor: {correct_creditor} ({dept_sums[correct_creditor]:+.2f})")
print(f"  Correct СуммаВзаимозачета = {correct_sum}")

# Also check mass processor logic
print("\n=== MASS PROCESSOR logic (А_ОбработкаДисбалансаПоПодразделениям) ===")
print("Takes MAX single row, NOT sum per dept:")
max_k = 0
min_k = 0
for r in rows:
    if r["ctrl"] > max_k:
        max_k = r["ctrl"]
    if r["ctrl"] < min_k:
        min_k = r["ctrl"]
print(f"  МаксК (max single row) = {max_k}")
print(f"  Correct sum = {correct_sum}")
print(f"  BUG in mass processor too: {max_k} != {correct_sum}")
