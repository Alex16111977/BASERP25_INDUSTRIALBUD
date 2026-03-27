# -*- coding: utf-8 -*-
import urllib.request, json, base64

url = "http://localhost/BaseERP/hs/mcp"
creds = base64.b64encode("Администратор:24043".encode("utf-8")).decode()

# Запрос: подразделение Воропаєва за декабрь 2025
# Используем UUID физлица напрямую через условие по коду
query = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "ПРЕДСТАВЛЕНИЕ(р.ПодотчетноеЛицо) КАК ФИО, "
    "ПРЕДСТАВЛЕНИЕ(р.Подразделение) КАК Подразделение, "
    "ПРЕДСТАВЛЕНИЕ(р.СтатьяДвиженияДенежныхСредств) КАК Статья "
    "ИЗ РегистрНакопления.ДенежныеСредстваУПодотчетныхЛиц КАК р "
    "ГДЕ р.Период >= ДАТАВРЕМЯ(2025,12,1,0,0,0) "
    "И р.Период < ДАТАВРЕМЯ(2026,1,1,0,0,0) "
    "И р.ПодотчетноеЛицо.Код = \"0000000738\""
)

payload = {"jsonrpc":"2.0","id":1,"method":"tools/call",
    "params":{"name":"execute_query","arguments":{"query_text": query}}}
data = json.dumps(payload).encode("utf-8")
req = urllib.request.Request(url, data=data, headers={
    "Content-Type": "application/json", "Authorization": "Basic " + creds})
with urllib.request.urlopen(req) as resp:
    result = json.loads(resp.read().decode("utf-8"))
    text = result["result"]["content"][0]["text"]
    print("isError:", result["result"].get("isError"))
    print(text[:2000])
