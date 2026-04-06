# -*- coding: utf-8 -*-
"""
Тест: ручний ВзаимозачетЗадолженности IB00-000084 від 22.12.2025
+ авто 000Ц-000422 — перевірка фільтрації та зв'язку
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)

d1 = datetime.datetime(2025, 12, 22)
d2 = datetime.datetime(2025, 12, 23)
d_start = datetime.datetime(2025, 12, 1)
d_end = datetime.datetime(2025, 12, 31, 23, 59, 59)

print("=" * 70)
print("1. REKVIZITY OBOH DOKUMENTIV")
print("=" * 70)

q = conn.NewObject("Query")
q.Text = (
    "ВЫБРАТЬ "
    "   Док.Ссылка КАК Ссылка, "
    "   Док.Номер КАК Номер, "
    "   Док.Проведен КАК Проведен, "
    "   Док.СуммаРегл КАК СуммаРегл, "
    "   Док.А_ДокументОснованиеДляБаланса КАК ДокОснование "
    "ИЗ "
    "   Документ.ВзаимозачетЗадолженности КАК Док "
    "ГДЕ "
    "   Док.Дата МЕЖДУ &Нач И &Кон "
    "   И (Док.Номер ПОДОБНО \"%084\" ИЛИ Док.Номер ПОДОБНО \"%422\")"
)
q.SetParameter("Нач", d1)
q.SetParameter("Кон", d2)

res = q.Execute().Select()
docs = {}
while res.Next():
    num = str(res.Номер).strip()
    has_basis = conn.ValueIsFilled(res.ДокОснование)
    print(f"  {num}: posted={res.Проведен} sum={res.СуммаРегл} basis_filled={has_basis}")
    if has_basis:
        print(f"    basis = {res.ДокОснование}")
    docs[num] = {"ref": res.Ссылка, "has_basis": has_basis, "posted": res.Проведен, "sum": res.СуммаРегл}

# === 2. Рухи в ПрочиеАктивыПассивы ===
print("\n" + "=" * 70)
print("2. MOVEMENTS in ПрочиеАктивыПассивы")
print("=" * 70)

for num, info in docs.items():
    q2 = conn.NewObject("Query")
    q2.Text = (
        "ВЫБРАТЬ "
        "   Об.Подразделение КАК Подразделение, "
        "   Об.Статья КАК Статья, "
        "   Об.СуммаПриход КАК Дебет, "
        "   Об.СуммаРасход КАК Кредит, "
        "   Об.СуммаПриход - Об.СуммаРасход КАК Контроль "
        "ИЗ "
        "   РегистрНакопления.ПрочиеАктивыПассивы.Обороты(,,Регистратор,) КАК Об "
        "ГДЕ Об.Регистратор = &Рег"
    )
    q2.SetParameter("Рег", info["ref"])
    res2 = q2.Execute().Select()

    print(f"\n  {num} (posted={info['posted']}):")
    depts = {}
    while res2.Next():
        dept = str(res2.Подразделение)
        ctrl = round(float(res2.Контроль), 2)
        print(f"    {dept}: Dt={res2.Дебет:.2f} Kt={res2.Кредит:.2f} Ctrl={ctrl} ({res2.Статья})")
        depts[dept] = depts.get(dept, 0) + ctrl

    if depts:
        debtor = max(depts, key=depts.get)
        creditor = min(depts, key=depts.get)
        print(f"    => Debtor: {debtor} ({depts[debtor]:.2f}), Creditor: {creditor} ({depts[creditor]:.2f})")

# === 3. НайтиСуществующийВзаимозачет ===
print("\n" + "=" * 70)
print("3. FIND EXISTING VZAIMOZACHET for IB00-000084")
print("=" * 70)

ref_084 = None
for num, info in docs.items():
    if "084" in num:
        ref_084 = info["ref"]
        break

if ref_084:
    q3 = conn.NewObject("Query")
    q3.Text = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 "
        "   Док.Ссылка, Док.Номер, Док.СуммаРегл, Док.Проведен "
        "ИЗ Документ.ВзаимозачетЗадолженности КАК Док "
        "ГДЕ Док.А_ДокументОснованиеДляБаланса = &ДокОснование "
        "   И НЕ Док.ПометкаУдаления"
    )
    q3.SetParameter("ДокОснование", ref_084)
    res3 = q3.Execute().Select()
    if res3.Next():
        print(f"  FOUND: {str(res3.Номер).strip()} posted={res3.Проведен} sum={res3.СуммаРегл}")
        print(f"  => IB00-000084 Обработан = sums_match AND posted")
        sum_084 = docs.get("ІБ00-000084", docs.get(next(k for k in docs if "084" in k), {})).get("sum", 0)
        # disbalance sum for 084 (from movements)
    else:
        print("  NOT FOUND! 000Ц-000422 does not have А_ДокументОснованиеДляБаланса = IB00-000084")

# === 4. Тест ВЫРАЗИТЬ ===
print("\n" + "=" * 70)
print("4. TEST VYRAZIT FILTER in query")
print("=" * 70)

q5 = conn.NewObject("Query")
q5.Text = (
    "ВЫБРАТЬ РАЗЛИЧНЫЕ "
    "   ВЫРАЗИТЬ(Об.Регистратор КАК Документ.ВзаимозачетЗадолженности).Номер КАК Номер, "
    "   ВЫРАЗИТЬ(Об.Регистратор КАК Документ.ВзаимозачетЗадолженности).А_ДокументОснованиеДляБаланса КАК ДокОснование "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(&Нач, &Кон, Регистратор, ) КАК Об "
    "ГДЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности "
    "   И (ВЫРАЗИТЬ(Об.Регистратор КАК Документ.ВзаимозачетЗадолженности).Номер ПОДОБНО \"%084\" "
    "      ИЛИ ВЫРАЗИТЬ(Об.Регистратор КАК Документ.ВзаимозачетЗадолженности).Номер ПОДОБНО \"%422\")"
)
q5.SetParameter("Нач", d_start)
q5.SetParameter("Кон", d_end)

try:
    res5 = q5.Execute().Select()
    while res5.Next():
        num = str(res5.Номер).strip()
        has_basis = conn.ValueIsFilled(res5.ДокОснование)
        action = "EXCLUDE (auto)" if has_basis else "INCLUDE (manual)"
        print(f"  {num}: basis_filled={has_basis} => {action}")
    print("  VYRAZIT works OK in query!")
except Exception as e:
    print(f"  VYRAZIT ERROR: {e}")
    print("  => Need code-level filtering!")

# === 5. Код-рівень ===
print("\n" + "=" * 70)
print("5. CODE-LEVEL: all vzaimozachety in period")
print("=" * 70)

q6 = conn.NewObject("Query")
q6.Text = (
    "ВЫБРАТЬ РАЗЛИЧНЫЕ Об.Регистратор КАК Документ "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(&Нач, &Кон, Регистратор, ) КАК Об "
    "ГДЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности"
)
q6.SetParameter("Нач", d_start)
q6.SetParameter("Кон", d_end)

res6 = q6.Execute().Select()
manual_count = 0
auto_count = 0
while res6.Next():
    doc = res6.Документ
    try:
        has_basis = conn.ValueIsFilled(doc.А_ДокументОснованиеДляБаланса)
    except:
        has_basis = False
    if has_basis:
        auto_count += 1
    else:
        manual_count += 1
        print(f"  MANUAL: {doc}")

print(f"\n  Total: auto={auto_count} (exclude), manual={manual_count} (include)")

print("\n" + "=" * 70)
print("DONE")
