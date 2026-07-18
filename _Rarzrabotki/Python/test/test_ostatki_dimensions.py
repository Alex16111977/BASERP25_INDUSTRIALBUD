# -*- coding: utf-8 -*-
"""
Тест: перевірити що ОстаткиИОбороты обох регістрів повертають
НаправлениеДеятельности та ДокументПередачи як виміри.
"""
import sys
import pythoncom
import win32com.client

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
D1 = win32com.client.pywintypes.TimeType(2026, 1, 1, 0, 0, 0)
D2 = win32com.client.pywintypes.TimeType(2026, 3, 31, 23, 59, 59)

# ===== ERP =====
print("=" * 60)
print("ERP: A_ДенежныеСредстваФинАгенты.ОстаткиИОбороты")
print("=" * 60)

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
erp = v8.Connect(CONN_ERP)

# Test 1: Check all columns from virtual table
q = erp.NewObject("Query")
q.Text = (
    "SELECT TOP 1 * "
    "FROM AccumulationRegister.А_ДенежныеСредстваФинАгенты.BalanceAndTurnovers(&D1, &D2, , , ) AS T"
)
q.SetParameter("D1", D1)
q.SetParameter("D2", D2)
res = q.Execute()
cols = res.Columns
print(f"All columns ({cols.Count()}):")
for i in range(cols.Count()):
    col = cols.Get(i)
    print(f"  [{i}] {col.Name}")

# Test 2: Explicit SELECT with НаправлениеДеятельности and ДокументПередачи
print("\nExplicit SELECT with dimensions:")
q2 = erp.NewObject("Query")
q2.Text = (
    "SELECT TOP 5 "
    "T.ФинАгент AS ФинАгент, "
    "T.НаправлениеДеятельности AS НаправлениеДеятельности, "
    "T.ДокументПередачи AS ДокументПередачи, "
    "T.СуммаКонечныйОстаток AS КонОст "
    "FROM AccumulationRegister.А_ДенежныеСредстваФинАгенты.BalanceAndTurnovers(&D1, &D2, , , ) AS T "
    "WHERE T.СуммаКонечныйОстаток <> 0"
)
q2.SetParameter("D1", D1)
q2.SetParameter("D2", D2)
try:
    res2 = q2.Execute()
    sel = res2.Choose()
    count = 0
    while sel.Next():
        count += 1
        print(f"  {count}. ФинАгент={sel.ФинАгент}")
        print(f"     НаправлениеДеятельности={sel.НаправлениеДеятельности}")
        print(f"     ДокументПередачи={sel.ДокументПередачи}")
        print(f"     КонОст={sel.КонОст}")
    print(f"ERP OK: {count} rows with dimensions")
except Exception as e:
    print(f"ERP ERROR: {e}")

# ===== КАЗНА =====
print("\n" + "=" * 60)
print("КАЗНА: ДенежныеСредстваФинАгенты.ОстаткиИОбороты")
print("=" * 60)

CONN_KAZN = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'
kazn = v8.Connect(CONN_KAZN)

# Test 3: All columns
q3 = kazn.NewObject("Query")
q3.Text = (
    "SELECT TOP 1 * "
    "FROM AccumulationRegister.ДенежныеСредстваФинАгенты.BalanceAndTurnovers(&D1, &D2, , , ) AS T"
)
q3.SetParameter("D1", D1)
q3.SetParameter("D2", D2)
res3 = q3.Execute()
cols3 = res3.Columns
print(f"All columns ({cols3.Count()}):")
for i in range(cols3.Count()):
    col = cols3.Get(i)
    print(f"  [{i}] {col.Name}")

# Test 4: Explicit SELECT with Направление and ДокументПередачи
print("\nExplicit SELECT with dimensions:")
q4 = kazn.NewObject("Query")
q4.Text = (
    "SELECT TOP 5 "
    "T.ФинАгент AS ФинАгент, "
    "T.Направление AS Направление, "
    "T.ДокументПередачи AS ДокументПередачи, "
    "T.СуммаКонечныйОстаток AS КонОст "
    "FROM AccumulationRegister.ДенежныеСредстваФинАгенты.BalanceAndTurnovers(&D1, &D2, , , ) AS T "
    "WHERE T.СуммаКонечныйОстаток <> 0"
)
q4.SetParameter("D1", D1)
q4.SetParameter("D2", D2)
try:
    res4 = q4.Execute()
    sel = res4.Choose()
    count = 0
    while sel.Next():
        count += 1
        fa = str(sel.ФинАгент) if sel.ФинАгент else "<empty>"
        nd_name = ""
        nd_code = ""
        try:
            nd_name = str(kazn.String(sel.Направление))
            nd_code = str(kazn.String(sel.Направление.Код))
        except:
            pass
        dp = ""
        try:
            dp = str(kazn.String(sel.ДокументПередачи))
        except:
            pass
        print(f"  {count}. ФинАгент={fa}")
        print(f"     Направление={nd_name} (Code={nd_code})")
        print(f"     ДокументПередачи={dp}")
        print(f"     КонОст={sel.КонОст}")
    print(f"KAZNA OK: {count} rows with dimensions")
except Exception as e:
    print(f"KAZNA ERROR: {e}")

# ===== МАППІНГ Направление -> НаправленияДеятельности =====
print("\n" + "=" * 60)
print("MAPPING: Kazna Направление -> ERP НаправленияДеятельности")
print("=" * 60)

q5 = kazn.NewObject("Query")
q5.Text = "SELECT Code, Description FROM Catalog.Направления WHERE NOT DeletionMark"
try:
    res5 = q5.Execute()
    sel = res5.Choose()
    matched = 0
    total = 0
    while sel.Next():
        total += 1
        code = str(kazn.String(sel.Code))
        name = str(kazn.String(sel.Description))
        erp_ref = erp.Catalogs.НаправленияДеятельности.FindByCode(code.strip())
        erp_filled = erp.ValueIsFilled(erp_ref)
        status = "OK" if erp_filled else "NOT FOUND"
        if erp_filled:
            matched += 1
        print(f"  Kazna: Code={code}, Name={name} -> ERP: {status}")
    print(f"\nMapping: {matched}/{total} matched")
except Exception as e:
    print(f"MAPPING ERROR: {e}")

print("\n=== TEST COMPLETED ===")
