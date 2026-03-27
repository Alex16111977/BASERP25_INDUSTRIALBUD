# -*- coding: utf-8 -*-
"""
Создание документов ВводОстатковСПодотчетниками по 10 сотрудникам
"""
import urllib.request, json, base64

url = "http://localhost/BaseERP/hs/mcp"
creds = base64.b64encode("Администратор:24043".encode("utf-8")).decode()
HEADERS = {"Content-Type": "application/json", "Authorization": "Basic " + creds}

ORG_UUID      = "6bee36b2-53f0-11e6-80d3-000c29bbac23"
VALUTA_UUID   = "86433645-6ea2-11f0-a2de-bccbaebe2890"
STATYA_UUID   = "11258552-c7dd-11e8-a200-000c299fb278"
HOZOP         = "ХозяйственныеОперации.ВводОстатковЗадолженностиПодотчетников"

# Список: (ФИО, сумма, код физлица если известен)
EMPLOYEES = [
    ("Гірієв Борис Зезікович",           4200.00),
    ("Задригун Іван Іванович",            4200.00),
    ("Манжула Максим Анатолійович",        608.58),
    ("Міщенко Анатолій Анатолійович",     5600.00),
    ("Огоновський Дмитро Анатолійович",   4200.00),
    ("Присада Дмитро Васильович",         5600.00),
    ("Присада Максим Васильович",         2800.00),
    ("Сидоренко Володимир Миколайович",   1800.00),
    ("Сніжко Юрій Тимофійович",           2800.00),
    ("Хурсін Тарас Євгенійович",          4600.00),
]

def call(name, args):
    payload = {"jsonrpc":"2.0","id":1,"method":"tools/call",
               "params":{"name": name, "arguments": args}}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=60) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        text = result["result"]["content"][0]["text"]
        try:
            return json.loads(text), result["result"].get("isError", False)
        except:
            return text, result["result"].get("isError", False)

results = []

for fio, summa in EMPLOYEES:
    print(f"\n{'='*60}")
    print(f"Обрабатываю: {fio} | {summa}")

    # Шаг 1: найти UUID физлица
    ref, err = call("find_catalog_ref", {"catalog_type": "ФизическиеЛица", "name": fio})
    if err or not ref.get("found"):
        print(f"  ❌ Физлицо не найдено: {ref}")
        results.append({"fio": fio, "status": "ОШИБКА", "detail": "Физлицо не найдено"})
        continue
    person_uuid = ref["uuid"]
    print(f"  ✅ UUID физлица: {person_uuid}")

    # Шаг 2: найти подразделение из декабря 2025
    q = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 "
        "ПРЕДСТАВЛЕНИЕ(р.Подразделение) КАК Подразделение "
        "ИЗ РегистрНакопления.ДенежныеСредстваУПодотчетныхЛиц КАК р "
        f"ГДЕ р.Период >= ДАТАВРЕМЯ(2025,12,1,0,0,0) "
        f"И р.Период < ДАТАВРЕМЯ(2026,1,1,0,0,0) "
        f"И р.ПодотчетноеЛицо.Код = \"{ref['code']}\""
    )
    rows, err2 = call("execute_query", {"query_text": q})
    podrazdel_name = ""
    podrazdel_uuid = ""
    if not err2 and rows and rows[0].get("Подразделение"):
        podrazdel_name = rows[0]["Подразделение"]
        # Получить UUID подразделения
        pref, perr = call("find_catalog_ref", {"catalog_type": "СтруктураПредприятия", "name": podrazdel_name})
        if not perr and pref.get("found"):
            podrazdel_uuid = pref["uuid"]
            print(f"  ✅ Подразделение: {podrazdel_name} | UUID: {podrazdel_uuid}")
        else:
            print(f"  ⚠️  Подразделение '{podrazdel_name}' не найдено по имени, пробую без него")
    else:
        print(f"  ⚠️  Подразделение не найдено в декабре 2025 (нет движений)")

    # Шаг 3: создать документ
    doc_data = {
        "Дата": "2025-11-30T12:00:00",
        "Организация": ORG_UUID,
        "Валюта": VALUTA_UUID,
        "ХозяйственнаяОперация": HOZOP,
        "ОтражатьВБУиНУ": True,
        "ОтражатьВОперативномУчете": True,
        "ОтражатьВУУ": True,
        "РасчетыСПодотчетниками": [{
            "ПодотчетноеЛицо": person_uuid,
            "Валюта": VALUTA_UUID,
            "Сумма": summa,
            "СуммаРегл": summa,
            "СуммаУпр": summa,
            "СтатьяДвиженияДенежныхСредств": STATYA_UUID
        }]
    }
    if podrazdel_uuid:
        doc_data["Подразделение"] = podrazdel_uuid

    doc, derr = call("create_document", {"doc_type": "ВводОстатковСПодотчетниками", "data": doc_data, "post": False})
    if derr or not doc.get("success"):
        print(f"  ❌ Ошибка создания: {doc}")
        results.append({"fio": fio, "status": "ОШИБКА", "detail": str(doc)})
    else:
        print(f"  ✅ Документ создан: {doc['number']} | ref: {doc['ref']}")
        results.append({
            "fio": fio,
            "summa": summa,
            "number": doc["number"],
            "ref": doc["ref"],
            "podrazdel": podrazdel_name,
            "status": "OK"
        })

print(f"\n{'='*60}")
print("ИТОГ:")
for r in results:
    if r["status"] == "OK":
        print(f"  ✅ {r['fio']:45s} {r['summa']:10.2f} грн | {r['number']} | {r['podrazdel']}")
    else:
        print(f"  ❌ {r['fio']:45s} ОШИБКА: {r['detail']}")
