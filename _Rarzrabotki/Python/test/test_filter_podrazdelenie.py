# -*- coding: utf-8 -*-
"""
Test: filtr po podrazdeleniyu v obrabotke A_ObrabotkaDisbalansa
Proverka: pri filtre "Astarta. Tishchenki" dokument PeremTovarov IB00-001591
NE dolzhen popadat (tam Globino-2 i KMD-2, net Astarta)
"""
import win32com.client, pythoncom
from datetime import datetime
from collections import defaultdict

CONN_STR = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK\n")

NACH = datetime(2025, 12, 1)
KON = datetime(2025, 12, 31, 23, 59, 59)
FILTR_DEPT = "Астарта. Тищенки"

# 1. Zapros vsekh oborotov (BEZ filtra v zaprose)
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
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"""
q.SetParameter("НачалоПериода", NACH)
q.SetParameter("ОкончаниеПериода", KON)

tab = q.Execute().Выгрузить()
print(f"Vsego strok: {tab.Количество()}")

# 2. Gruppirovka po registratoru
docs = defaultdict(list)
for i in range(tab.Количество()):
    row = tab.Получить(i)
    dok_xml = conn.XMLСтрока(row.Документ)
    dept_name = row.Подразделение.Наименование if conn.ЗначениеЗаполнено(row.Подразделение) else "?"
    docs[dok_xml].append({
        'dok_ref': row.Документ,
        'dept_name': dept_name,
        'dept_ref': row.Подразделение,
        'kontrol': float(row.Контроль),
    })

print(f"Unikalnykh dokumentov: {len(docs)}")

# 3. Filtracia + opredelenie storon
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

    # FILTR: dokument popadaet tolko esli filtr-podrazdelenie = debitor ILI kreditor
    if FILTR_DEPT:
        if debitor['dept_name'] != FILTR_DEPT and kreditor['dept_name'] != FILTR_DEPT:
            continue

    results.append({
        'dok': str(rows[0]['dok_ref']),
        'deb': debitor['dept_name'],
        'kred': kreditor['dept_name'],
        'summa': abs(debitor['kontrol']),
    })

print(f"\nS filtrom '{FILTR_DEPT}': {len(results)} dokumentov\n")

# 4. Proverka
ib001591_found = False
for r in results:
    if "001591" in r['dok']:
        ib001591_found = True
        print(f"  !!! IB00-001591 NAJDEN: Deb={r['deb']}, Kred={r['kred']}")

    # Pokazat pervye 10
    if results.index(r) < 10:
        print(f"  {r['dok'][:60]:60s} Deb={r['deb']:25s} Kred={r['kred']:25s} Sum={r['summa']:12.2f}")

if not ib001591_found:
    print(f"\n  OK: IB00-001591 NE najden pri filtre '{FILTR_DEPT}'")
else:
    print(f"\n  BUG: IB00-001591 NAJDEN pri filtre '{FILTR_DEPT}' - ne dolzhen!")

# 5. Proverka chto IB00-001591 imeet Globino-2 i KMD-2
print("\n=== Proverka IB00-001591 ===")
for dok_xml, rows in docs.items():
    if any("001591" in str(r['dok_ref']) for r in rows):
        for r in rows:
            print(f"  {r['dept_name']:25s} K={r['kontrol']:+.2f}")
        break
