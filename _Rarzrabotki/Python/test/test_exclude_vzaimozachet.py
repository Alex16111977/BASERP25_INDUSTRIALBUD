# -*- coding: utf-8 -*-
"""
Test: isklyuchit VzaimozachetZadolzhennosti iz zaprosa disbalansa.
Vzaimozachet - eto instrument ispravleniya, a ne istochnik problemy.

Proverka:
1. SpisanieNedostach IB00-000756 dolzhen popadat (eto istochnik disbalansa)
2. Vzaimozachet 000C-000006 NE dolzhen popadat (eto instrument)
"""
import win32com.client, pythoncom
from datetime import datetime
from collections import defaultdict

CONN_STR = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK\n")

NACH = datetime(2025, 12, 1)
KON = datetime(2025, 12, 31, 23, 59, 59)
FILTR_DEPT = "Астарта. Тищенки"

# Zapros BEZ VzaimozachetZadolzhennosti
q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Об.Регистратор КАК Документ,
    Об.Подразделение КАК Подразделение,
    Об.Организация КАК Организация,
    Об.СуммаПриход - Об.СуммаРасход КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(
    &НачалоПериода, &ОкончаниеПериода, Регистратор, ) КАК Об
ГДЕ НЕ Об.Статья В (
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
    И НЕ Об.Регистратор ССЫЛКА Документ.ВзаимозачетЗадолженности"""
q.SetParameter("НачалоПериода", NACH)
q.SetParameter("ОкончаниеПериода", KON)
tab = q.Execute().Выгрузить()
print(f"Strok BEZ vzaimozachetov: {tab.Количество()}")

# Gruppirovka
docs = defaultdict(list)
for i in range(tab.Количество()):
    row = tab.Получить(i)
    dok_xml = conn.XMLСтрока(row.Документ)
    dept_name = row.Подразделение.Наименование if conn.ЗначениеЗаполнено(row.Подразделение) else "?"
    docs[dok_xml].append({
        'dok_ref': row.Документ,
        'dept_name': dept_name,
        'kontrol': float(row.Контроль),
    })

# Filtracia
results = []
for dok_xml, rows in docs.items():
    total = sum(r['kontrol'] for r in rows)
    if abs(total) > 0.01:
        continue
    depts_disb = [r for r in rows if abs(r['kontrol']) > 0.01]
    if len(depts_disb) < 2:
        continue
    debitor = max(depts_disb, key=lambda x: x['kontrol'])
    kreditor = min(depts_disb, key=lambda x: x['kontrol'])
    if debitor['kontrol'] <= 0 or kreditor['kontrol'] >= 0:
        continue
    if debitor['dept_name'] == kreditor['dept_name']:
        continue
    if FILTR_DEPT:
        if debitor['dept_name'] != FILTR_DEPT and kreditor['dept_name'] != FILTR_DEPT:
            continue
    results.append({
        'dok': str(rows[0]['dok_ref']),
        'deb': debitor['dept_name'],
        'kred': kreditor['dept_name'],
        'summa': abs(debitor['kontrol']),
    })

print(f"S filtrom '{FILTR_DEPT}': {len(results)} dokumentov\n")

# Proverki
print("=== Rezultaty ===")
found_756 = False
found_vzaimozachet = False
for r in results:
    if "000756" in r['dok']:
        found_756 = True
    if "Взаимозачет" in r['dok'] or "000Ц-000" in r['dok']:
        found_vzaimozachet = True
        print(f"  !!! VZAIMOZACHET v spiske: {r['dok'][:60]}")
    print(f"  {r['dok'][:60]:60s} Deb={r['deb']:25s} Kred={r['kred']:25s} Sum={r['summa']:12.2f}")

print()
if found_756:
    print("OK: IB00-000756 NAJDEN v spiske")
else:
    print("PROBLEM: IB00-000756 NE najden!")

if not found_vzaimozachet:
    print("OK: Vzaimozachety NE popadayut v spisok")
else:
    print("BUG: Vzaimozachety popadayut v spisok!")
