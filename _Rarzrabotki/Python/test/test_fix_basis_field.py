# -*- coding: utf-8 -*-
"""
Test: set А_ДокументОснованиеДляБаланса on 000Ц-000422 = IB00-000084
Now the field type is cfg:DocumentRef (accepts any document)
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

def s(obj):
    try:
        if not conn.ValueIsFilled(obj): return "<EMPTY>"
        return conn.String(obj)
    except:
        return str(obj)

# Find both docs
q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ Док.Ссылка, Док.Номер КАК Н, "
    "   Док.А_ДокументОснованиеДляБаланса КАК Осн "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.СуммаРегл = 244812.62"
)
res = q.Execute().Select()
ref_084 = None
ref_422 = None
while res.Next():
    num = s(res.Н).strip()
    has_osn = conn.ValueIsFilled(res.Осн)
    print(f"{num}: А_ДокОсн filled={has_osn}")
    if "084" in num and "000Ц" not in num:
        ref_084 = res.Ссылка
    elif "422" in num:
        ref_422 = res.Ссылка

if not ref_084 or not ref_422:
    print("Docs not found!"); exit()

print(f"\nIB00-000084: {s(ref_084)}")
print(f"000Ц-000422: {s(ref_422)}")

# Try to set the field
print(f"\nSetting А_ДокОсн on 000Ц-000422 = IB00-000084...")
try:
    obj = ref_422.ПолучитьОбъект()
    obj.ОбменДанными.Загрузка = True
    obj.А_ДокументОснованиеДляБаланса = ref_084
    obj.Записать()
    print("  Written OK!")
except Exception as e:
    print(f"  ERROR: {e}")

# Verify
print(f"\nVerify:")
q2 = conn.NewObject("Query")
q2.Text = (
    "ВЫБРАТЬ Док.А_ДокументОснованиеДляБаланса = НЕОПРЕДЕЛЕНО КАК Пусто "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.СуммаРегл = 244812.62 И Док.Ссылка = &Р"
)
q2.SetParameter("Р", ref_422)
r2 = q2.Execute().Select()
if r2.Next():
    print(f"  000Ц-000422.А_ДокОсн is empty: {r2.Пусто}")
    if not r2.Пусто:
        print(f"  SUCCESS! Field is now filled!")
    else:
        print(f"  FAILED - still empty")

# Check НайтиСуществующийВзаимозачет
print(f"\nНайтиСуществующийВзаимозачет(IB00-000084):")
q3 = conn.NewObject("Query")
q3.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Номер КАК Н "
    "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ Док.А_ДокументОснованиеДляБаланса = &Осн "
    "   И НЕ Док.ПометкаУдаления"
)
q3.SetParameter("Осн", ref_084)
r3 = q3.Execute().Select()
if r3.Next():
    print(f"  FOUND: {s(r3.Н).strip()}")
else:
    print(f"  NOT FOUND")
