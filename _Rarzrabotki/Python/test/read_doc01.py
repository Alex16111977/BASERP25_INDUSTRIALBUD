# -*- coding: utf-8 -*-
import urllib.request, json, base64

url = "http://localhost/BaseERP/hs/mcp"
creds = base64.b64encode("Администратор:24043".encode("utf-8")).decode()

def query(q):
    payload = {"jsonrpc":"2.0","id":1,"method":"tools/call",
        "params":{"name":"execute_query","arguments":{"query_text": q}}}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={
        "Content-Type": "application/json", "Authorization": "Basic " + creds})
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        text = result["result"]["content"][0]["text"]
        if not text.strip(): return []
        try:
            return json.loads(text)
        except:
            print("RAW:", text[:1000])
            return []

# Шапка документа 0Ц-00000001
print("=== ШАПКА 0Ц-00000001 ===")
rows = query(
    "ВЫБРАТЬ "
    "д.Организация.Наименование КАК Организация, "
    "д.Подразделение.Наименование КАК Подразделение, "
    "д.Валюта.Наименование КАК Валюта, "
    "д.ХозяйственнаяОперация, "
    "д.ОтражатьВБУиНУ, д.ОтражатьВОперативномУчете, д.ОтражатьВУУ "
    "ИЗ Документ.ВводОстатковСПодотчетниками КАК д "
    "ГДЕ д.Проведен = ИСТИНА"
)
for r in rows:
    print(json.dumps(r, ensure_ascii=False, indent=2))
