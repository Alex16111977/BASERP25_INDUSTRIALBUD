# -*- coding: utf-8 -*-
import urllib.request, json, base64

url = "http://localhost/BaseERP/hs/mcp"
creds = base64.b64encode("Администратор:24043".encode("utf-8")).decode()
HEADERS = {"Content-Type": "application/json", "Authorization": "Basic " + creds}

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

# Коды физлиц из предыдущего запуска
EMPLOYEES = [
    ("ВОРОНЦОВ ОЛЕКСАНДР ВОЛОДИМИРОВИЧ", "0000000074", "8797edf0-1b1c-11f1-a2e6-8a1890236e3d"),
    ("Пронін Роман Вікторович",           "0000000770", "8797edf1-1b1c-11f1-a2e6-8a1890236e3d"),
    ("Шкурат Ігор Миколайович",           "0000000249", "8efb0cfe-1b1c-11f1-a2e6-8a1890236e3d"),
]

for fio, code, doc_ref in EMPLOYEES:
    print(f"\n{fio}")
    # Ищем в любом регистре за весь декабрь + ноябрь 2025
    q = (
        "ВЫБРАТЬ ПЕРВЫЕ 3 "
        "ПРЕДСТАВЛЕНИЕ(р.Подразделение) КАК Подразделение, "
        "ПРЕДСТАВЛЕНИЕ(р.Регистратор) КАК Регистратор "
        "ИЗ РегистрНакопления.ДенежныеСредстваУПодотчетныхЛиц КАК р "
        f"ГДЕ р.ПодотчетноеЛицо.Код = \"{code}\" "
        "И р.Период >= ДАТАВРЕМЯ(2025,11,1,0,0,0) "
        "И р.Период < ДАТАВРЕМЯ(2026,1,1,0,0,0) "
        "УПОРЯДОЧИТЬ ПО р.Период УБЫВ"
    )
    rows, err = call("execute_query", {"query_text": q})
    if not err and rows:
        for r in rows:
            print(f"  Подр: {r['Подразделение']} | Рег: {r['Регистратор']}")
    else:
        print(f"  Нет данных")
