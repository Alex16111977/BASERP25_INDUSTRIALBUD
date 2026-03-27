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

# UUID документов из предыдущего создания
DOCS = [
    ("38f56d44-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000005", "Воропаєв Дмитро Миколайович"),
    ("bc67df02-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000006", "Гірієв Борис Зезікович"),
    ("bc67fd57-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000007", "Задригун Іван Іванович"),
    ("c40f75fb-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000008", "Манжула Максим Анатолійович"),
    ("ca1b5159-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000009", "Міщенко Анатолій Анатолійович"),
    ("ca1b544b-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000010", "Огоновський Дмитро Анатолійович"),
    ("d2e0ef95-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000011", "Присада Дмитро Васильович"),
    ("d2e0ef96-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000012", "Присада Максим Васильович"),
    ("dc3ea215-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000013", "Сидоренко Володимир Миколайович"),
    ("dc3ea216-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000014", "Сніжко Юрій Тимофійович"),
    ("e4df36b4-1b1b-11f1-a2e6-8a1890236e3d", "0Ц-00000015", "Хурсін Тарас Євгенійович"),
]

print("Обновляю Комментарий во всех документах...\n")
for ref, number, fio in DOCS:
    res, err = call("update_document", {
        "doc_ref": ref,
        "data": {"Комментарий": fio.strip()},
        "post": False
    })
    if err or not res.get("success"):
        print(f"  ❌ {number} {fio} — ОШИБКА: {res}")
    else:
        print(f"  ✅ {number} | {fio}")
