# -*- coding: utf-8 -*-
"""
Тест: РаспределениеДоходовПоНаправлениямДеятельности — баг з порожнім Підрозділом
Документ: 000Ц-000002 від 31.12.2025 (UUID: a6f0a6dd-fd21-11f0-a2e2-c54425f51b91)

Перевірка:
1. РАСХОД (ПрибылиИУбытки) має Підрозділ = порожній → баг
2. Після виправлення: для кожного (Організація, Підрозділ)
   SUM(Приход ДоходыТекущегоПериода) == SUM(Расход ПрибылиИУбытки)
"""
import win32com.client
import sys

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN_ERP)

DOC_UUID = "a6f0a6dd-fd21-11f0-a2e2-c54425f51b91"

query = conn.NewObject("Query")
query.Text = """
ВЫБРАТЬ
    ВЫБОР КОГДА Д.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА "Приход" ИНАЧЕ "Расход" КОНЕЦ КАК ВидДвиж,
    Д.Подразделение КАК Подразделение,
    Д.Подразделение.Наименование КАК ПодразделениеНаим,
    Д.Статья.Наименование КАК СтатьяНаим,
    Д.Сумма КАК Сумма
ИЗ
    РегистрНакопления.ПрочиеАктивыПассивы КАК Д
ГДЕ
    Д.Регистратор = &Док
УПОРЯДОЧИТЬ ПО ВидДвиж, ПодразделениеНаим, СтатьяНаим
"""
doc_ref = conn.Documents.РаспределениеДоходовПоНаправлениямДеятельности.ПолучитьСсылку(
    conn.NewObject("УникальныйИдентификатор", DOC_UUID))
query.SetParameter("Док", doc_ref)

result = query.Execute()
sel = result.Choose()

print("=" * 100)
print("Рухи ПрочиеАктивыПассивы по документу РаспределениеДоходовПоНаправлениямДеятельности")
print("=" * 100)

# Collect data
rashod_pribyli = []  # РАСХОД ПрибылиИУбытки
prihod_dohody = []   # ПРИХОД ДоходыТекущегоПериода
all_rows = []

while sel.Next():
    vid = str(sel.ВидДвиж)
    podrazd = str(sel.ПодразделениеНаим) if conn.ValueIsFilled(sel.Подразделение) else "(ПОРОЖНІЙ)"
    statya = str(sel.СтатьяНаим)
    summa = float(sel.Сумма)

    row = {"vid": vid, "podrazd": podrazd, "statya": statya, "summa": summa}
    all_rows.append(row)

    if vid == "Расход" and "Прибыл" in statya:
        rashod_pribyli.append(row)
    elif vid == "Приход" and "Доход" in statya and "текущ" in statya.lower():
        prihod_dohody.append(row)

# Print all movements
for r in all_rows:
    print(f"  {r['vid']:10s} | {r['podrazd']:40s} | {r['statya']:30s} | {r['summa']:>15,.2f}")

print()
print("-" * 100)

# Check 1: РАСХОД ПрибылиИУбытки — Підрозділ порожній?
empty_podrazd = [r for r in rashod_pribyli if r["podrazd"] == "(ПОРОЖНІЙ)"]
print(f"\nРАСХОД ПрибылиИУбытки: {len(rashod_pribyli)} рядків, з них порожній підрозділ: {len(empty_podrazd)}")

if empty_podrazd:
    print(">>> БАГ ПІДТВЕРДЖЕНО: РАСХОД ПрибылиИУбытки має порожній Підрозділ!")
    for r in empty_podrazd:
        print(f"    Сума: {r['summa']:>15,.2f}")
else:
    print(">>> ВИПРАВЛЕНО: всі РАСХОД ПрибылиИУбытки мають заповнений Підрозділ")

# Check 2: Баланс по підрозділах
print("\n--- Баланс по підрозділах ---")
balance = {}
for r in prihod_dohody:
    key = r["podrazd"]
    balance.setdefault(key, {"prihod": 0, "rashod": 0})
    balance[key]["prihod"] += r["summa"]

for r in rashod_pribyli:
    key = r["podrazd"]
    balance.setdefault(key, {"prihod": 0, "rashod": 0})
    balance[key]["rashod"] += r["summa"]

ok = True
for podrazd, vals in sorted(balance.items()):
    diff = vals["prihod"] - vals["rashod"]
    status = "OK" if abs(diff) < 0.01 else "ДИСБАЛАНС"
    if abs(diff) >= 0.01:
        ok = False
    print(f"  {podrazd:40s} | Приход(Доходи): {vals['prihod']:>15,.2f} | Расход(Прибутки): {vals['rashod']:>15,.2f} | Різниця: {diff:>15,.2f} | {status}")

print()
if ok:
    print("РЕЗУЛЬТАТ: Баланс СХОДИТЬСЯ по всіх підрозділах")
else:
    print("РЕЗУЛЬТАТ: Баланс НЕ СХОДИТЬСЯ — потрібне виправлення!")

sys.exit(0 if ok and not empty_podrazd else 1)
