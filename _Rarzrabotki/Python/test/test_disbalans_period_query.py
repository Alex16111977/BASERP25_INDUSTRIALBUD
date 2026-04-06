# -*- coding: utf-8 -*-
"""
Test: zapros disbalansa za period (gruden 2025).
Emulyaciya logiki АнализДокументов() dlya А_ОбработкаДисбалансаПоПодразделениям.

Algoritm:
1. Zapros Oboroty po vsem dokumentam za period
2. Gruppirovat' v kode po Registratoru
3. Dlya kazhdogo: proverit' balans org (summa=0) + disbalans dept
4. Najti sushch. Vzaimozachety
5. Opredelit' StatyaDisbalansa
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

# === Krok 1: Zapros vsekh oborotov za period ===
print("=== Krok 1: Zapros oborotov ===")
q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Об.Регистратор КАК Документ,
    Об.Подразделение КАК Подразделение,
    Об.Организация КАК Организация,
    Об.СуммаПриход КАК Дебет,
    Об.СуммаРасход КАК Кредит,
    Об.СуммаПриход - Об.СуммаРасход КАК Контроль
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(
    &НачалоПериода, &ОкончаниеПериода, Регистратор, ) КАК Об
ГДЕ НЕ Об.Статья В (
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))"""
q.SetParameter("НачалоПериода", NACH)
q.SetParameter("ОкончаниеПериода", KON)

tab = q.Execute().Выгрузить()
print(f"  Vsego strok: {tab.Количество()}")

# === Krok 2: Gruppirovat po Registratoru ===
print("\n=== Krok 2: Gruppirovka po Registratoru ===")

# Sobiraem danye v slovar: dok_xml -> [{dept, dept_ref, org_ref, kontrol}]
docs = defaultdict(list)
for i in range(tab.Количество()):
    row = tab.Получить(i)
    dok_ref = row.Документ
    dok_xml = conn.XMLСтрока(dok_ref)
    dept_ref = row.Подразделение
    dept_name = dept_ref.Наименование if conn.ЗначениеЗаполнено(dept_ref) else "?"
    org_ref = row.Организация
    kontrol = float(row.Контроль)
    docs[dok_xml].append({
        'dok_ref': dok_ref,
        'dept_ref': dept_ref,
        'dept_name': dept_name,
        'org_ref': org_ref,
        'kontrol': kontrol,
    })

print(f"  Unikalnykh dokumentov: {len(docs)}")

# === Krok 3: Filtraciya - balans org OK + disbalans dept ===
print("\n=== Krok 3: Dokumenty s disbalansom ===")

disbalans_docs = []
for dok_xml, rows in docs.items():
    total_kontrol = sum(r['kontrol'] for r in rows)
    # Balans po org dolzhen shoditsya (=0)
    if abs(total_kontrol) > 0.01:
        continue
    # Dolzhno byt 2+ podrazdelenij s kontrol <> 0
    depts_with_disbalans = [r for r in rows if abs(r['kontrol']) > 0.01]
    if len(depts_with_disbalans) < 2:
        continue

    # Najti debitor i kreditor
    debitor = max(depts_with_disbalans, key=lambda x: x['kontrol'])
    kreditor = min(depts_with_disbalans, key=lambda x: x['kontrol'])

    if debitor['kontrol'] <= 0 or kreditor['kontrol'] >= 0:
        continue

    disbalans_docs.append({
        'dok_ref': rows[0]['dok_ref'],
        'dok_xml': dok_xml,
        'org_ref': rows[0]['org_ref'],
        'debitor_dept': debitor['dept_name'],
        'debitor_ref': debitor['dept_ref'],
        'kreditor_dept': kreditor['dept_name'],
        'kreditor_ref': kreditor['dept_ref'],
        'summa': abs(debitor['kontrol']),
    })

print(f"  Dokumentov s disbalansom: {len(disbalans_docs)}")
print()

# Pokzhat pervye 20
for i, d in enumerate(disbalans_docs[:20]):
    print(f"  {i+1:3d}. {str(d['dok_ref']):60s}")
    print(f"       Deb: {d['debitor_dept']:25s}  Kred: {d['kreditor_dept']:25s}  Summa: {d['summa']:12.2f}")

if len(disbalans_docs) > 20:
    print(f"  ... i eshche {len(disbalans_docs) - 20}")

# === Krok 4: Proverit na izvestnykh keysakh ===
print("\n=== Proverka izvestnykh keysov ===")
known = {
    "PeremTov 001591": ("001591", "Глобино-2", "КМД-2", 5810.40),
    "PeremTov 001666": ("001666", "Глобино-2", "КМД-2", 126602.90),
}
for label, (num, exp_deb, exp_kred, exp_sum) in known.items():
    found = [d for d in disbalans_docs if num in str(d['dok_ref'])]
    if found:
        d = found[0]
        ok_deb = exp_deb in d['debitor_dept']
        ok_kred = exp_kred in d['kreditor_dept']
        ok_sum = abs(d['summa'] - exp_sum) < 0.01
        status = "OK" if (ok_deb and ok_kred and ok_sum) else "FAIL"
        print(f"  [{status}] {label}: Deb={d['debitor_dept']}, Kred={d['kreditor_dept']}, Sum={d['summa']:.2f}")
    else:
        print(f"  [MISS] {label}: ne najden v rezultatakh")

print("\n=== GOTOVO ===")
