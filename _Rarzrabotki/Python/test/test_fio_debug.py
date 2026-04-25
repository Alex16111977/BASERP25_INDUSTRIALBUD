# -*- coding: utf-8 -*-
import win32com.client, sys, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Get doc
q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q.SetParameter("Н", "000Ц-000001")
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

# Get ФИО from document А_РаспределениеКазна (empty Сотрудник rows)
doc_fios = set()
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    if not conn.ЗначениеЗаполнено(row.Сотрудник):
        doc_fios.add(S(row.ФИО))

# Get ФИО from Kazna
conn_kazn = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
Sk = conn_kazn.String

period = obj.ПериодРегистрации
ps = S(period)
parts = ps.split(".")
month, year = int(parts[1]), int(parts[2].split()[0])
beg = datetime.datetime(year, month, 1)
end = datetime.datetime(year + (1 if month == 12 else 0), (month % 12) + 1, 1) - datetime.timedelta(seconds=1)

qk = conn_kazn.NewObject("Запрос")
qk.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    БДДС.Сотрудник.ИНН КАК ИНН,
    БДДС.Сотрудник.Наименование КАК ФИО,
    БДДС.Сотрудник.Код КАК КодС
ИЗ РегистрНакопления.БДДС КАК БДДС
ГДЕ БДДС.Период МЕЖДУ &Н И &К
    И БДДС.Регистратор.Организация.А_ПереносЕРП
    И (БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
        ИЛИ БДДС.Регистратор ССЫЛКА Документ.РаспределениеФ2)"""
qk.SetParameter("Н", beg)
qk.SetParameter("К", end)
selk = qk.Execute().Choose()

kazna_fios = {}
while selk.Next():
    fio = Sk(selk.ФИО)  # WITHOUT strip - raw value
    inn = Sk(selk.ИНН)
    kod = Sk(selk.КодС)
    kazna_fios[fio] = (inn, kod)

# Compare
print("=== FIO comparison: Document vs Kazna ===")
print(f"Document empty-Сотрудник unique ФИО: {len(doc_fios)}")
print(f"Kazna unique ФИО: {len(kazna_fios)}")

# Pick one example from each
sample_doc = sorted(doc_fios)[:3]
sample_kazna = sorted(kazna_fios.keys())[:3]
print(f"\nDoc samples:")
for f in sample_doc:
    print(f"  [{f}] len={len(f)} bytes={f.encode('utf-8').hex()[:40]}")
print(f"\nKazna samples:")
for f in sample_kazna:
    print(f"  [{f}] len={len(f)} bytes={f.encode('utf-8').hex()[:40]}")

# Try matching
matched = 0
not_matched = []
for doc_fio in sorted(doc_fios):
    # Exact
    if doc_fio in kazna_fios:
        matched += 1
        continue
    # Strip
    doc_stripped = doc_fio.strip()
    found = False
    for kf in kazna_fios:
        if kf.strip() == doc_stripped:
            matched += 1
            found = True
            break
    if not found:
        not_matched.append(doc_fio)

print(f"\nMatched (exact or strip): {matched}")
print(f"Not matched: {len(not_matched)}")

if not_matched:
    print("\nNot matched ФИО (first 5):")
    for f in not_matched[:5]:
        print(f"  Doc: [{f}] len={len(f)}")
        # Find closest in kazna
        for kf in kazna_fios:
            if f.strip()[:10] == kf.strip()[:10]:
                print(f"  Kazna: [{kf}] len={len(kf)}")
                # Show byte diff
                for pos, (c1, c2) in enumerate(zip(f, kf)):
                    if c1 != c2:
                        print(f"    DIFF at pos {pos}: doc='{c1}'(U+{ord(c1):04X}) kazna='{c2}'(U+{ord(c2):04X})")
                        break
                break

# Now test: what if we use V83.String() differently?
print("\n=== Test V83.String encoding ===")
qk2 = conn_kazn.NewObject("Запрос")
qk2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 БДДС.Сотрудник.Наименование КАК ФИО
ИЗ РегистрНакопления.БДДС КАК БДДС
ГДЕ БДДС.Период МЕЖДУ &Н И &К
    И БДДС.Регистратор.Организация.А_ПереносЕРП
    И БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
    И БДДС.Сотрудник.Наименование ПОДОБНО "%Безсмертн%"
"""
qk2.SetParameter("Н", beg)
qk2.SetParameter("К", end)
selk2 = qk2.Execute().Choose()
if selk2.Next():
    kazna_fio_raw = Sk(selk2.ФИО)
    print(f"Kazna Безсмертній: [{kazna_fio_raw}] len={len(kazna_fio_raw)}")
    # Compare with doc
    for df in doc_fios:
        if "Безсмертн" in df:
            print(f"Doc   Безсмертній: [{df}] len={len(df)}")
            # Char-by-char comparison
            maxl = max(len(kazna_fio_raw), len(df))
            for pos in range(maxl):
                c1 = kazna_fio_raw[pos] if pos < len(kazna_fio_raw) else '∅'
                c2 = df[pos] if pos < len(df) else '∅'
                if c1 != c2:
                    print(f"  pos {pos}: kazna='{c1}'(U+{ord(c1):04X}) doc='{c2}'(U+{ord(c2):04X})")
            break
