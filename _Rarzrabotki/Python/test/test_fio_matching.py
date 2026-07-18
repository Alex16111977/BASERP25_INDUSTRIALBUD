# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка "
    "ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док "
    "ГДЕ НЕ Док.ПометкаУдаления УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
)
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

# Get empty employee rows with ФИО
print("=== Rows in A_Kazna with empty Сотрудник ===")
empty_fios = set()
for i in range(obj.А_РаспределениеКазна.Count()):
    row = obj.А_РаспределениеКазна.Get(i)
    if not conn.ЗначениеЗаполнено(row.Сотрудник):
        fio = S(row.ФИО).strip()
        if fio and fio not in empty_fios:
            empty_fios.add(fio)
            print(f"  ФИО=[{fio}] len={len(fio)}")

# Get employee names from A_Nalogi
print("\n=== Employee names from A_Nalogi (Строка(Сотрудник)) ===")
nalogi_names = {}
for i in range(obj.А_РаспределениеНалоги.Count()):
    row = obj.А_РаспределениеНалоги.Get(i)
    if conn.ЗначениеЗаполнено(row.Сотрудник):
        name = S(row.Сотрудник).strip()
        if name not in nalogi_names:
            nalogi_names[name] = row.Сотрудник

# Match
print(f"\n=== Matching ===")
print(f"Empty ФИО: {len(empty_fios)}")
print(f"Nalogi names: {len(nalogi_names)}")
matched = 0
not_matched = []
for fio in sorted(empty_fios):
    if fio in nalogi_names:
        matched += 1
    else:
        not_matched.append(fio)
        # Try without trailing spaces
        fio_stripped = fio.rstrip()
        if fio_stripped in nalogi_names:
            print(f"  Would match with strip: [{fio}] -> [{fio_stripped}]")

print(f"\nMatched: {matched}")
print(f"Not matched: {len(not_matched)}")
for nm in not_matched:
    print(f"  [{nm}]")
