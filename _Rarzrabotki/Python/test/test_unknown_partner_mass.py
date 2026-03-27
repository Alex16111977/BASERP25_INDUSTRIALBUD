# -*- coding: utf-8 -*-
"""
Массовый тест: замена НеизвестногоПартнера во всех справочниках.
Обрабатывает: ДоговорыКонтрагентов, КлючиАналитикиУчетаПоПартнерам.
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

def query(q, params=None, max_rows=200):
    args = {"query_text": q, "max_rows": max_rows}
    if params: args["parameters"] = params
    return mcp_call("execute_query", args)

def find_ref(catalog_type, name):
    return mcp_call("find_catalog_ref", {"catalog_type": catalog_type, "name": name})

def update_item(uuid, data):
    return mcp_call("update_catalog_item", {"catalog_ref": uuid, "data": data})

# Кеш партнеров чтобы не искать повторно
partner_cache = {}

def get_or_create_partner(kontragent_name):
    """Получить или создать партнера по имени контрагента."""
    if kontragent_name in partner_cache:
        return partner_cache[kontragent_name]

    # Ищем партнера
    ref = find_ref("Партнеры", kontragent_name)
    if ref and isinstance(ref, dict) and ref.get('found') and ref.get('uuid'):
        partner_cache[kontragent_name] = ref['uuid']
        return ref['uuid']

    # Партнер не найден — создаём
    print(f"    Создаём нового партнера: {kontragent_name}")
    result = mcp_call("create_catalog_item", {
        "catalog_type": "Партнеры",
        "data": {"Наименование": kontragent_name, "НаименованиеПолное": kontragent_name}
    })
    if result and isinstance(result, dict) and result.get('success'):
        uuid = result.get('ref')
        partner_cache[kontragent_name] = uuid
        return uuid
    else:
        print(f"    ОШИБКА создания партнера: {result}")
        return None

# =========================================
# 1. Договоры контрагентов
# =========================================
print("=" * 60)
print("Обработка ДоговорыКонтрагентов")
print("=" * 60)

contracts = query(
    "ВЫБРАТЬ "
    "Д.Ссылка КАК Ссылка, "
    "Д.Наименование КАК Наименование, "
    "Д.Контрагент.Наименование КАК КонтрагентНаименование "
    "ИЗ Справочник.ДоговорыКонтрагентов КАК Д "
    "ГДЕ Д.Партнер.Наименование = \"Неизвестный партнер\"",
    max_rows=500
)

if not contracts:
    print("  Нет договоров с НеизвестнымПартнером")
else:
    print(f"  Найдено: {len(contracts)} договоров")
    ok = 0
    err = 0
    for c in contracts:
        kontragent_name = c['КонтрагентНаименование']
        partner_uuid = get_or_create_partner(kontragent_name)
        if not partner_uuid:
            print(f"  SKIP: {c['Наименование']} — нет партнера для {kontragent_name}")
            err += 1
            continue

        ref = find_ref("ДоговорыКонтрагентов", c['Наименование'])
        if not ref or not isinstance(ref, dict) or not ref.get('uuid'):
            print(f"  SKIP: {c['Наименование']} — не найден UUID")
            err += 1
            continue

        result = update_item(ref['uuid'], {"Партнер": partner_uuid})
        if result and isinstance(result, dict) and result.get('success'):
            ok += 1
        else:
            print(f"  ERR: {c['Наименование']} — {result}")
            err += 1

    print(f"  Итого: {ok} успешно, {err} ошибок")

# =========================================
# 2. Ключи аналитики учета по партнерам
# =========================================
print()
print("=" * 60)
print("Обработка КлючиАналитикиУчетаПоПартнерам")
print("=" * 60)

keys = query(
    "ВЫБРАТЬ "
    "К.Ссылка КАК Ссылка, "
    "К.Наименование КАК Наименование, "
    "ПРЕДСТАВЛЕНИЕ(К.Контрагент) КАК КонтрагентПредставление "
    "ИЗ Справочник.КлючиАналитикиУчетаПоПартнерам КАК К "
    "ГДЕ К.Партнер.Наименование = \"Неизвестный партнер\"",
    max_rows=500
)

if not keys:
    print("  Нет ключей с НеизвестнымПартнером")
else:
    print(f"  Найдено: {len(keys)} ключей")
    ok = 0
    err = 0
    for k in keys:
        kontragent_name = k['КонтрагентПредставление']
        if not kontragent_name:
            print(f"  SKIP: {k['Наименование']} — нет контрагента")
            err += 1
            continue

        partner_uuid = get_or_create_partner(kontragent_name)
        if not partner_uuid:
            print(f"  SKIP: {k['Наименование']} — нет партнера")
            err += 1
            continue

        ref = find_ref("КлючиАналитикиУчетаПоПартнерам", k['Наименование'])
        if not ref or not isinstance(ref, dict) or not ref.get('uuid'):
            print(f"  SKIP: {k['Наименование']} — не найден UUID")
            err += 1
            continue

        result = update_item(ref['uuid'], {"Партнер": partner_uuid})
        if result and isinstance(result, dict) and result.get('success'):
            ok += 1
        else:
            print(f"  ERR: {k['Наименование']} — {result}")
            err += 1

    print(f"  Итого: {ok} успешно, {err} ошибок")

# =========================================
# 3. Проверка результата
# =========================================
print()
print("=" * 60)
print("Проверка остатков")
print("=" * 60)

remaining_contracts = query(
    "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол "
    "ИЗ Справочник.ДоговорыКонтрагентов КАК Д "
    "ГДЕ Д.Партнер.Наименование = \"Неизвестный партнер\""
)
print(f"  Договоры с НеизвестнымПартнером: {remaining_contracts[0]['Кол'] if remaining_contracts else '?'}")

remaining_keys = query(
    "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол "
    "ИЗ Справочник.КлючиАналитикиУчетаПоПартнерам КАК К "
    "ГДЕ К.Партнер.Наименование = \"Неизвестный партнер\""
)
print(f"  Ключи аналитики с НеизвестнымПартнером: {remaining_keys[0]['Кол'] if remaining_keys else '?'}")
