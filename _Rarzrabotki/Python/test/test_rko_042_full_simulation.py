# -*- coding: utf-8 -*-
"""
Full simulation: how А_ДвижениеПоВзаимозачетПоБалансуУпр processes
РКО 000Ц-000042 от 28.01.2026

Step by step reproducing the EXACT 1C code logic to find the bug.
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    try:
        if not conn.ValueIsFilled(obj): return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

# Find document
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
print(f"Doc: {s(doc_ref)}\n")

# === Step 1: Query (exact same as processor) ===
print("=== Step 1: Query Обороты ===")
q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ "
    "   Об.Подразделение КАК Подразделение, "
    "   Об.Организация КАК Организация, "
    "   Об.СуммаПриход КАК Дебет, "
    "   Об.СуммаРасход КАК Кредит, "
    "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
    "ГДЕ Об.Регистратор = &Р "
    "   И НЕ Об.Статья В ("
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),"
    "       ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"
)
q2.SetParameter("Р", doc_ref)
table = q2.Execute().Выгрузить()

print(f"Rows in table: {table.Количество()}")
for i in range(table.Количество()):
    row = table.Получить(i)
    dept = s(row.Подразделение)
    ctrl = round(float(row.Контроль), 2)
    print(f"  [{i}] {dept}: Dt={row.Дебет:.2f} Kt={row.Кредит:.2f} Ctrl={ctrl}")

# === Step 1.5: Org balance check ===
print("\n=== Step 1.5: Org balance ===")
total = 0
for i in range(table.Количество()):
    total += table.Получить(i).Контроль
print(f"Total Ctrl: {round(total, 2)} (should be 0)")

# === Step 2: Group by department (NEW code with Соответствие) ===
print("\n=== Step 2: Group by department (Соответствие) ===")
dept_map = {}  # XMLСтрока(Подразделение) -> {Подразделение, Организация, Контроль}
for i in range(table.Количество()):
    row = table.Получить(i)
    key = conn.XMLСтрока(row.Подразделение)
    if key not in dept_map:
        dept_map[key] = {
            "dept": row.Подразделение,
            "org": row.Организация,
            "ctrl": float(row.Контроль)
        }
    else:
        dept_map[key]["ctrl"] += float(row.Контроль)

print(f"Departments after grouping: {len(dept_map)}")
for key, data in dept_map.items():
    print(f"  {s(data['dept'])}: Ctrl={round(data['ctrl'], 2)}")

# === Step 3: Determine debtor/creditor ===
print("\n=== Step 3: Determine debtor/creditor ===")
debtor = None
creditor = None
sum_vz = 0

for key, data in dept_map.items():
    if data["ctrl"] > sum_vz:
        debtor = data
        sum_vz = data["ctrl"]
    if data["ctrl"] < 0:
        if creditor is None or data["ctrl"] < creditor["ctrl"]:
            creditor = data

if debtor:
    print(f"Debtor: {s(debtor['dept'])} Ctrl={round(debtor['ctrl'], 2)}")
if creditor:
    print(f"Creditor: {s(creditor['dept'])} Ctrl={round(creditor['ctrl'], 2)}")
print(f"СуммаВзаимозачета = {round(sum_vz, 2)}")

# === Check existing vzaimozachet ===
print("\n=== Existing vzaimozachet ===")
q3 = conn.NewObject("Query")
q3.Text = (
    "ВЫБРАТЬ Док.Ссылка, Док.Номер КАК Н, Док.СуммаРегл "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.А_ДокументОснованиеДляБаланса = &Осн "
    "   И НЕ Док.ПометкаУдаления"
)
q3.SetParameter("Осн", doc_ref)
r3 = q3.Execute().Select()
if r3.Next():
    print(f"Found: {s(r3.Н).strip()} sum={r3.СуммаРегл}")
    print(f"MISMATCH: existing={r3.СуммаРегл} vs correct={round(sum_vz, 2)}")
else:
    print("Not found")

print("\n=== VERDICT ===")
print(f"The vzaimozachet should be created for {round(sum_vz, 2)}")
print(f"NOT for 0.01 (last raw row) or 0.10 (max single row)")
