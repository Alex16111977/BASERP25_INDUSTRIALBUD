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

# Ищем все движения Проніна без ограничения по дате
q = (
    "ВЫБРАТЬ ПЕРВЫЕ 5 "
    "ПРЕДСТАВЛЕНИЕ(р.Регистратор) КАК Рег, "
    "ПРЕДСТАВЛЕНИЕ(р.Подразделение) КАК Подр, "
    "р.Период "
    "ИЗ РегистрНакопления.ДенежныеСредстваУПодотчетныхЛиц КАК р "
    "ГДЕ р.ПодотчетноеЛицо.Наименование ПОДОБНО \"%Прон%\" "
    "УПОРЯДОЧИТЬ ПО р.Период УБЫВ"
)
rows, err = call("execute_query", {"query_text": q})
print("Пронін - все движения:")
if not err and rows:
    for r in rows: print(f"  {r}")
else:
    print("  Нет:", str(rows)[:400])
