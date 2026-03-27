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

# UUID Проніна: 19d332f4-1117-4361-9afa-48445bf81030
q = (
    "ВЫБРАТЬ ПЕРВЫЕ 3 "
    "д.Номер, д.Подразделение.Наименование КАК Подразделение, д.Дата "
    "ИЗ Документ.АвансовыйОтчет КАК д "
    "ГДЕ д.ФизЛицо.Код = \"0000000770\" "
    "И д.Дата >= ДАТАВРЕМЯ(2025,11,1,0,0,0) "
    "УПОРЯДОЧИТЬ ПО д.Дата УБЫВ"
)
rows, err = call("execute_query", {"query_text": q})
print("Пронін - ФизЛицо.Код:")
if not err and rows:
    for r in rows: print(f"  {r}")
else:
    # Пробуем через Сотрудник
    q2 = (
        "ВЫБРАТЬ ПЕРВЫЕ 3 "
        "д.Номер, д.Подразделение.Наименование КАК Подразделение, д.Дата "
        "ИЗ Документ.АвансовыйОтчет КАК д "
        "ГДЕ д.Сотрудник.ФизическоеЛицо.Код = \"0000000770\" "
        "И д.Дата >= ДАТАВРЕМЯ(2025,11,1,0,0,0) "
        "УПОРЯДОЧИТЬ ПО д.Дата УБЫВ"
    )
    rows2, err2 = call("execute_query", {"query_text": q2})
    print("Пронін - Сотрудник.ФизЛицо.Код:")
    if not err2 and rows2:
        for r in rows2: print(f"  {r}")
    else:
        print("  Нет:", rows2)
