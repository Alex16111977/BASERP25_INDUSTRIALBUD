"""Сравнение группировок: Рш.Подразделение (ПодрОрг) vs Рш.Подразделение.А_Подразделение (СтрПредпр)
для ВКассу 000Ц-000009 от 10.12.2025 (РКО N0000052986)."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Рш.Подразделение КАК ПодрОрг,
    Рш.Подразделение.А_Подразделение КАК АПодр,
    Рш.КВыплате КАК Сумма
ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу.А_РасшифровкаВыплатыЗарплатаПоФизлицам КАК Рш
ГДЕ Рш.Ссылка.Номер = &Ном
"""
q.SetParameter("Ном", "000Ц-000009")
r = q.Execute().Выгрузить()
print(f"Строк в Ведомость.А_Расшифровке: {r.Количество()}")

# Группировка по ПодрОрг
g_podrorg = {}
g_apodr = {}
total = 0.0
for row in r:
    k1 = S(row.ПодрОрг)
    k2 = S(row.АПодр)
    s = float(row.Сумма)
    total += s
    g_podrorg[k1] = g_podrorg.get(k1, 0.0) + s
    g_apodr[k2] = g_apodr.get(k2, 0.0) + s

print(f"\nΣ = {total}")
print(f"\nГруппировка по Рш.Подразделение (ПодрОрг) — {len(g_podrorg)} групп:")
for k, v in sorted(g_podrorg.items()):
    print(f"  {k:40} {v}")

print(f"\nГруппировка по Рш.Подразделение.А_Подразделение (СтрПредпр) — {len(g_apodr)} групп:")
for k, v in sorted(g_apodr.items()):
    print(f"  {k:40} {v}")

# Свод: какие ПодрОрг ссылаются на какие А_Подр
print(f"\nСвод ПодрОрг → А_Подр:")
mapping = {}
for row in r:
    k1 = S(row.ПодрОрг)
    k2 = S(row.АПодр)
    mapping[k1] = k2
for k, v in sorted(mapping.items()):
    mark = "" if k == v else "  ⚠ ПЕРЕВОД"
    print(f"  {k:40} → {v}{mark}")
