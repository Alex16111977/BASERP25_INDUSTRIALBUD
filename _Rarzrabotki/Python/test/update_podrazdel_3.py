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

UPDATES = [
    ("8797edf0-1b1c-11f1-a2e6-8a1890236e3d", "0Ц-00000016", "ВОРОНЦОВ ОЛЕКСАНДР ВОЛОДИМИРОВИЧ", "Астарта. Тищенки"),
    ("8797edf1-1b1c-11f1-a2e6-8a1890236e3d", "0Ц-00000017", "Пронін Роман Вікторович",           "Строительство"),
    ("8efb0cfe-1b1c-11f1-a2e6-8a1890236e3d", "0Ц-00000018", "Шкурат Ігор Миколайович",           "Астарта. Тищенки"),
]

for doc_ref, number, fio, podrazdel_name in UPDATES:
    # Получить UUID подразделения
    pref, perr = call("find_catalog_ref", {"catalog_type": "СтруктураПредприятия", "name": podrazdel_name})
    if perr or not pref.get("found"):
        print(f"  ❌ {number} {fio} — подразделение не найдено: {pref}")
        continue
    podrazdel_uuid = pref["uuid"]

    res, err = call("update_document", {
        "doc_ref": doc_ref,
        "data": {"Подразделение": podrazdel_uuid},
        "post": False
    })
    if err or not res.get("success"):
        print(f"  ❌ {number} | {fio} — ОШИБКА: {res}")
    else:
        print(f"  ✅ {number} | {fio} | {podrazdel_name}")
