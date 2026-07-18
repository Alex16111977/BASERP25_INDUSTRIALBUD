# -*- coding: utf-8 -*-
"""Check 000Ц-000422 and IB00-000084"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

# 000Ц-000422 by UUID
uuid_422 = "40907ccd-2ea5-11f1-a2e7-eb06fbf9b98b"
mgr = conn.Документы.ВзаимозачетЗадолженности
ref_422 = mgr.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", uuid_422))

print("=== 000C-000422 ===")
print(f"  Posted: {ref_422.Проведен}")
print(f"  Sum: {ref_422.СуммаРегл}")
print(f"  А_ДокОсн: filled={conn.ValueIsFilled(ref_422.А_ДокументОснованиеДляБаланса)}")
print(f"  А_ВведенВЕРП: {ref_422.А_ВведенВЕРП}")

# Find ALL vzaimozachety on 22.12.2025
print("\n=== ALL vzaimozachety 22.12.2025 ===")
q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка, Док.Номер, Док.Дата, Док.СуммаРегл, "
    "   Док.А_ДокументОснованиеДляБаланса КАК Осн "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.Дата МЕЖДУ &Н И &К"
)
q.SetParameter("Н", datetime.datetime(2025, 12, 22))
q.SetParameter("К", datetime.datetime(2025, 12, 22, 23, 59, 59))

res = q.Execute().Select()
ref_084 = None
while res.Next():
    num = str(res.Номер).strip()
    has_osn = conn.ValueIsFilled(res.Осн)
    print(f"  {num} dt={res.Дата} sum={res.СуммаРегл} basis_filled={has_osn}")
    if "084" in num and not "000Ц" in num:
        ref_084 = res.Ссылка
        print(f"    ^^^ THIS IS IB00-000084")

if ref_084:
    print(f"\n=== Movements IB00-000084 ===")
    q2 = conn.NewObject("Query")
    q2.Text = (
        "ВЫБРАТЬ Об.Подразделение КАК П, Об.Статья КАК С, "
        "   Об.СуммаПриход КАК Д, Об.СуммаРасход КАК К, "
        "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
        "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
        "ГДЕ Об.Регистратор = &Р"
    )
    q2.SetParameter("Р", ref_084)
    res2 = q2.Execute().Select()
    depts = {}
    while res2.Next():
        dept = str(res2.П)
        ctrl = round(float(res2.Контроль), 2)
        print(f"  {dept}: Dt={res2.Д:.2f} Kt={res2.К:.2f} Ctrl={ctrl} ({res2.С})")
        depts[dept] = depts.get(dept, 0) + ctrl

    if depts:
        print(f"\n  Depts summary: {depts}")
        debtor = max(depts, key=depts.get)
        creditor = min(depts, key=depts.get)
        print(f"  Debtor: {debtor} ({depts[debtor]:.2f}), Creditor: {creditor} ({depts[creditor]:.2f})")

    # Search vzaimozachet with basis = ref_084
    print(f"\n=== НайтиСуществующийВзаимозачет(IB00-000084) ===")
    q3 = conn.NewObject("Query")
    q3.Text = (
        "ВЫБРАТЬ Док.Ссылка, Док.Номер, Док.Проведен, Док.СуммаРегл "
        "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
        "ГДЕ Док.А_ДокументОснованиеДляБаланса = &Осн И НЕ Док.ПометкаУдаления"
    )
    q3.SetParameter("Осн", ref_084)
    res3 = q3.Execute().Select()
    if res3.Next():
        print(f"  FOUND: {str(res3.Номер).strip()} posted={res3.Проведен} sum={res3.СуммаРегл}")
    else:
        print("  NOT FOUND!")
        print("  => 000Ц-000422.А_ДокОсн is EMPTY, so it can't be found!")
        print()

        # Fix: set the field
        print("=== FIX: Setting А_ДокументОснованиеДляБаланса on 000Ц-000422 ===")
        try:
            obj = ref_422.ПолучитьОбъект()
            obj.А_ДокументОснованиеДляБаланса = ref_084
            obj.Записать(1)  # РежимЗаписиДокумента.Проведение = 1
            print("  SUCCESS! Field set and document posted.")

            # Verify
            ref_422_new = mgr.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", uuid_422))
            filled = conn.ValueIsFilled(ref_422_new.А_ДокументОснованиеДляБаланса)
            print(f"  Verify: basis_filled={filled}")
            if filled:
                print(f"  basis = {ref_422_new.А_ДокументОснованиеДляБаланса}")

            # Re-check НайтиСуществующийВзаимозачет
            res3b = q3.Execute().Select()
            if res3b.Next():
                print(f"  НайтиСуществующийВзаимозачет NOW finds: {str(res3b.Номер).strip()}")
            else:
                print("  STILL NOT FOUND after fix!")
        except Exception as e:
            print(f"  ERROR: {e}")
else:
    print("  IB00-000084 not found on 22.12.2025!")
