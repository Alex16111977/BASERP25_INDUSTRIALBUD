# -*- coding: utf-8 -*-
"""
Fix: set А_ДокументОснованиеДляБаланса on 000Ц-000422 = IB00-000084
And test that НайтиСуществующийВзаимозачет works after
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

mgr = conn.Документы.ВзаимозачетЗадолженности

# Get refs by UUID
uuid_422 = "40907ccd-2ea5-11f1-a2e7-eb06fbf9b98b"
ref_422 = mgr.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", uuid_422))

# Find IB00-000084 by sum (unique)
q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.СуммаРегл = 244812.62 И Док.Ссылка <> &Ref422"
)
q.SetParameter("Ref422", ref_422)
res = q.Execute().Select()
if not res.Next():
    print("ERROR: IB00-000084 not found!")
    exit()

ref_084 = res.Ссылка
print(f"IB00-000084: {ref_084}")
print(f"000C-000422: {ref_422}")

# Check current state
print(f"\nBEFORE fix:")
print(f"  000C-000422.А_ДокументОснованиеДляБаланса = {ref_422.А_ДокументОснованиеДляБаланса}")
print(f"  Filled: {conn.ValueIsFilled(ref_422.А_ДокументОснованиеДляБаланса)}")

# Fix: set the field
print(f"\nApplying fix...")
try:
    obj = ref_422.ПолучитьОбъект()
    obj.ОбменДанными.Загрузка = True  # Skip ПередЗаписью handlers
    obj.А_ДокументОснованиеДляБаланса = ref_084
    obj.Записать()  # Write without posting mode to preserve current state
    print("  Written OK (with ОбменДанными.Загрузка=True)")
except Exception as e:
    print(f"  ERROR write: {e}")
    print("  Trying without Загрузка flag...")
    try:
        obj = ref_422.ПолучитьОбъект()
        obj.А_ДокументОснованиеДляБаланса = ref_084
        obj.Записать()
        print("  Written OK")
    except Exception as e2:
        print(f"  ERROR: {e2}")
        exit()

# Verify
print(f"\nAFTER fix:")
# Re-read from DB
ref_422_v2 = mgr.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", uuid_422))
basis = ref_422_v2.А_ДокументОснованиеДляБаланса
print(f"  000C-000422.А_ДокументОснованиеДляБаланса = {basis}")
print(f"  Filled: {conn.ValueIsFilled(basis)}")
print(f"  Posted: {ref_422_v2.Проведен}")

# Test НайтиСуществующийВзаимозачет
print(f"\nTest НайтиСуществующийВзаимозачет(IB00-000084):")
q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка, Док.Номер КАК Н, Док.СуммаРегл КАК С, Док.Проведен КАК П "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.А_ДокументОснованиеДляБаланса = &Осн И НЕ Док.ПометкаУдаления"
)
q2.SetParameter("Осн", ref_084)
res2 = q2.Execute().Select()
if res2.Next():
    print(f"  FOUND: {str(res2.Н).strip()} posted={res2.П} sum={res2.С}")
    print(f"  => IB00-000084 will show Обработан=True!")
else:
    print(f"  NOT FOUND! Fix didn't work.")
    print(f"  Possible: ОбменДанными.Загрузка didn't help, ПередЗаписью cleared the field")
