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
            print("ERR:", text[:500])
            return []

print("=== Шапка 0Ц-00000005 ===")
rows = query(
    "ВЫБРАТЬ "
    "д.Организация.Наименование КАК Орг, "
    "д.Подразделение.Наименование КАК Подр, "
    "ПРЕДСТАВЛЕНИЕ(д.ХозяйственнаяОперация) КАК ХозОп, "
    "д.ОтражатьВБУиНУ, д.ОтражатьВОперативномУчете, д.ОтражатьВУУ "
    "ИЗ Документ.ВводОстатковСПодотчетниками КАК д "
    "ГДЕ д.ПометкаУдаления = ЛОЖЬ И д.Проведен = ЛОЖЬ "
    "УПОРЯДОЧИТЬ ПО д.Дата УБЫВ"
)
for r in rows:
    print(r)

print("\n=== ТЧ 0Ц-00000005 ===")
rows2 = query(
    "ВЫБРАТЬ "
    "тч.ПодотчетноеЛицо.Наименование КАК ФИО, "
    "тч.Сумма, тч.Валюта.Наименование КАК Валюта, "
    "тч.СтатьяДвиженияДенежныхСредств.Наименование КАК Статья "
    "ИЗ Документ.ВводОстатковСПодотчетниками.РасчетыСПодотчетниками КАК тч "
    "ГДЕ тч.Ссылка.ПометкаУдаления = ЛОЖЬ И тч.Ссылка.Проведен = ЛОЖЬ "
    "УПОРЯДОЧИТЬ ПО тч.Ссылка.Дата УБЫВ"
)
for r in rows2:
    print(r)
