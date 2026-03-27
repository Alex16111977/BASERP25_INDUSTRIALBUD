# -*- coding: utf-8 -*-
"""
Тест: замена НеизвестногоПартнера в Ключе аналитики учета по партнерам.
"""
import urllib.request, json, base64, sys

url = "http://localhost/BaseERP/hs/mcp"
creds = base64.b64encode("Администратор:24043".encode("utf-8")).decode()

def mcp_call(tool_name, arguments):
    payload = {"jsonrpc":"2.0","id":1,"method":"tools/call",
        "params":{"name": tool_name, "arguments": arguments}}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={
        "Content-Type": "application/json", "Authorization": "Basic " + creds})
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        text = result["result"]["content"][0]["text"]
        if not text.strip(): return None
        try: return json.loads(text)
        except: return text

def query(q, params=None):
    args = {"query_text": q}
    if params: args["parameters"] = params
    return mcp_call("execute_query", args)

def find_catalog_ref(catalog_type, name=None):
    args = {"catalog_type": catalog_type}
    if name: args["name"] = name
    return mcp_call("find_catalog_ref", args)

def update_catalog_item(uuid, data):
    return mcp_call("update_catalog_item", {"catalog_ref": uuid, "data": data})

# ========================================================
# ШАГ 1: Найти ключ аналитики с НеизвестнымПартнером
# ========================================================
print("=" * 60)
print("ШАГ 1: Ищем ключи аналитики с НеизвестнымПартнером")
print("=" * 60)

keys = query(
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "К.Ссылка КАК Ссылка, "
    "К.Наименование КАК Наименование, "
    "К.Партнер.Наименование КАК ПартнерНаименование, "
    "К.Контрагент КАК Контрагент, "
    "ПРЕДСТАВЛЕНИЕ(К.Контрагент) КАК КонтрагентПредставление "
    "ИЗ Справочник.КлючиАналитикиУчетаПоПартнерам КАК К "
    "ГДЕ К.Партнер.Наименование = \"Неизвестный партнер\""
)

if not keys:
    print("Нет ключей аналитики с НеизвестнымПартнером")
    sys.exit(0)

key = keys[0]
print(f"  Ключ: {key['Наименование']}")
print(f"  Контрагент: {key['КонтрагентПредставление']}")
print(f"  Партнер: {key['ПартнерНаименование']}")

kontragent_name = key['КонтрагентПредставление']

# ========================================================
# ШАГ 2: Определяем тип контрагента и ищем партнера
# ========================================================
print()
print("=" * 60)
print(f"ШАГ 2: Ищем партнера для '{kontragent_name}'")
print("=" * 60)

# Контрагент может быть типа Контрагенты или Организации
# Ищем среди контрагентов
kontragent_ref = find_catalog_ref("Контрагенты", name=kontragent_name)
print(f"  Контрагент ref: {kontragent_ref}")

# Проверяем партнера у контрагента
if kontragent_ref and isinstance(kontragent_ref, dict) and kontragent_ref.get('found'):
    kontragent_partner = query(
        "ВЫБРАТЬ К.Партнер.Наименование КАК ПН "
        "ИЗ Справочник.Контрагенты КАК К "
        "ГДЕ К.Наименование = &Имя",
        {"Имя": kontragent_name}
    )
    if kontragent_partner:
        print(f"  Партнер контрагента: {kontragent_partner[0]['ПН']}")

# Ищем целевого партнера
target_partner = find_catalog_ref("Партнеры", name=kontragent_name)
print(f"  Целевой партнер: {target_partner}")

if not target_partner or not isinstance(target_partner, dict) or not target_partner.get('found'):
    print(f"  Партнер '{kontragent_name}' не найден — пропускаем тест")
    sys.exit(0)

# ========================================================
# ШАГ 3: Обновляем ключ аналитики
# ========================================================
print()
print("=" * 60)
print("ШАГ 3: Обновляем партнера в ключе аналитики")
print("=" * 60)

key_ref = find_catalog_ref("КлючиАналитикиУчетаПоПартнерам", name=key['Наименование'])
print(f"  Ключ ref: {key_ref}")

if key_ref and isinstance(key_ref, dict) and key_ref.get('found') and key_ref.get('uuid'):
    result = update_catalog_item(key_ref['uuid'], {"Партнер": target_partner['uuid']})
    print(f"  Результат: {result}")
else:
    print("  ОШИБКА: Не удалось найти UUID ключа")
    sys.exit(1)

# ========================================================
# ШАГ 4: Проверяем
# ========================================================
print()
print("=" * 60)
print("ШАГ 4: Проверяем результат")
print("=" * 60)

verify = query(
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "К.Наименование КАК Наименование, "
    "К.Партнер.Наименование КАК ПН "
    "ИЗ Справочник.КлючиАналитикиУчетаПоПартнерам КАК К "
    "ГДЕ К.Наименование = &Имя",
    {"Имя": key['Наименование']}
)

if verify:
    v = verify[0]
    print(f"  Ключ: {v['Наименование']}")
    print(f"  Партнер: {v['ПН']}")
    if v['ПН'] != "Неизвестный партнер":
        print("  ✓ ТЕСТ ПРОЙДЕН!")
    else:
        print("  ✗ ТЕСТ НЕ ПРОЙДЕН")
