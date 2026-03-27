# -*- coding: utf-8 -*-
"""
Тест: замена НеизвестногоПартнера в договоре контрагента.
1. Находим договор с НеизвестнымПартнером
2. Определяем контрагента (Владелец договора)
3. Ищем/создаём правильного партнера по имени контрагента
4. Обновляем договор через MCP
5. Проверяем результат
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
        if not text.strip():
            return None
        try:
            return json.loads(text)
        except:
            return text

def query(q, params=None):
    args = {"query_text": q}
    if params:
        args["parameters"] = params
    return mcp_call("execute_query", args)

def find_catalog_ref(catalog_type, name=None, code=None):
    args = {"catalog_type": catalog_type}
    if name: args["name"] = name
    if code: args["code"] = code
    return mcp_call("find_catalog_ref", args)

def update_catalog_item(uuid, data):
    return mcp_call("update_catalog_item", {"catalog_ref": uuid, "data": data})

# ========================================================
# ШАГ 1: Найти договор с НеизвестнымПартнером
# ========================================================
print("=" * 60)
print("ШАГ 1: Ищем договоры с НеизвестнымПартнером")
print("=" * 60)

contracts = query(
    "ВЫБРАТЬ ПЕРВЫЕ 3 "
    "Д.Ссылка КАК Ссылка, "
    "Д.Наименование КАК Наименование, "
    "Д.Партнер.Наименование КАК ПартнерНаименование, "
    "Д.Контрагент КАК Контрагент, "
    "Д.Контрагент.Наименование КАК КонтрагентНаименование "
    "ИЗ Справочник.ДоговорыКонтрагентов КАК Д "
    "ГДЕ Д.Партнер.Наименование = \"Неизвестный партнер\" "
    "И Д.Наименование = \"ДОГОВІР-ЗАЯВКА на організацію послуг з перевезення\""
)

if not contracts:
    print("ОШИБКА: Не найдены договоры с НеизвестнымПартнером")
    sys.exit(1)

contract = contracts[0]
print(f"  Договор: {contract['Наименование']}")
print(f"  Контрагент: {contract['КонтрагентНаименование']}")
print(f"  Текущий партнер: {contract['ПартнерНаименование']}")

kontragent_name = contract['КонтрагентНаименование']

# ========================================================
# ШАГ 2: Проверяем — есть ли партнер с именем контрагента
# ========================================================
print()
print("=" * 60)
print(f"ШАГ 2: Ищем партнера с именем '{kontragent_name}'")
print("=" * 60)

partner_search = query(
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "П.Ссылка КАК Ссылка, "
    "П.Наименование КАК Наименование "
    "ИЗ Справочник.Партнеры КАК П "
    "ГДЕ П.Наименование = &Имя",
    {"Имя": kontragent_name}
)

if partner_search:
    print(f"  Найден партнер: {partner_search[0]['Наименование']}")
    # Проверяем — не занят ли этот партнер другим контрагентом
    check_busy = query(
        "ВЫБРАТЬ ПЕРВЫЕ 1 "
        "К.Ссылка КАК Ссылка, К.Наименование КАК Имя "
        "ИЗ Справочник.Контрагенты КАК К "
        "ГДЕ К.Партнер.Наименование = &ИмяПартнера "
        "И К.Наименование <> &ИмяКонтрагента",
        {"ИмяПартнера": kontragent_name, "ИмяКонтрагента": kontragent_name}
    )
    if check_busy:
        print(f"  ВНИМАНИЕ: Партнер занят другим контрагентом: {check_busy[0]['Имя']}")
        print("  Нужно создать нового партнера")
    else:
        print("  Партнер свободен — можно использовать")
else:
    print(f"  Партнер с именем '{kontragent_name}' НЕ найден")
    print("  Нужно создать нового партнера")

# ========================================================
# ШАГ 3: Получаем UUID контрагента и партнера
# ========================================================
print()
print("=" * 60)
print("ШАГ 3: Получаем UUID для обновления")
print("=" * 60)

# Ищем UUID контрагента
kontragent_ref = find_catalog_ref("Контрагенты", name=kontragent_name)
print(f"  Контрагент UUID: {kontragent_ref}")

# Ищем UUID партнера "Неизвестный партнер" для сравнения
unknown_ref = find_catalog_ref("Партнеры", name="Неизвестный партнер")
print(f"  НеизвестныйПартнер UUID: {unknown_ref}")

# Ищем или находим правильного партнера
target_partner_ref = find_catalog_ref("Партнеры", name=kontragent_name)
print(f"  Целевой партнер UUID: {target_partner_ref}")

# ========================================================
# ШАГ 4: Обновляем партнера в контрагенте (если нужно)
# ========================================================
print()
print("=" * 60)
print("ШАГ 4: Проверяем партнера у контрагента")
print("=" * 60)

kontragent_partner = query(
    "ВЫБРАТЬ "
    "К.Партнер.Наименование КАК ПартнерНаименование "
    "ИЗ Справочник.Контрагенты КАК К "
    "ГДЕ К.Наименование = &Имя",
    {"Имя": kontragent_name}
)

if kontragent_partner:
    current_partner = kontragent_partner[0]['ПартнерНаименование']
    print(f"  Текущий партнер контрагента: {current_partner}")
    if current_partner == "Неизвестный партнер" or not current_partner:
        print("  -> Нужно обновить партнера у контрагента")
        if target_partner_ref and isinstance(target_partner_ref, dict) and 'uuid' in target_partner_ref:
            if kontragent_ref and isinstance(kontragent_ref, dict) and 'uuid' in kontragent_ref:
                result = update_catalog_item(kontragent_ref['uuid'], {"Партнер": target_partner_ref['uuid']})
                print(f"  Результат обновления контрагента: {result}")
            else:
                print(f"  ОШИБКА: Не удалось получить UUID контрагента: {kontragent_ref}")
        else:
            print(f"  ОШИБКА: Не удалось получить UUID партнера: {target_partner_ref}")
    else:
        print("  Партнер уже назначен корректно")

# ========================================================
# ШАГ 5: Ищем UUID договора и обновляем
# ========================================================
print()
print("=" * 60)
print("ШАГ 5: Обновляем партнера в договоре")
print("=" * 60)

# Находим UUID договора
contract_uuid_data = query(
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "Д.Ссылка КАК Ссылка "
    "ИЗ Справочник.ДоговорыКонтрагентов КАК Д "
    "ГДЕ Д.Партнер.Наименование = \"Неизвестный партнер\" "
    "И Д.Наименование = \"ДОГОВІР-ЗАЯВКА на організацію послуг з перевезення\""
)

# Для обновления договора нужен UUID — получаем через find_catalog_ref
contract_ref = find_catalog_ref("ДоговорыКонтрагентов",
    name="ДОГОВІР-ЗАЯВКА на організацію послуг з перевезення")
print(f"  Договор ref: {contract_ref}")

if contract_ref and isinstance(contract_ref, dict) and 'uuid' in contract_ref:
    if target_partner_ref and isinstance(target_partner_ref, dict) and 'uuid' in target_partner_ref:
        result = update_catalog_item(contract_ref['uuid'], {"Партнер": target_partner_ref['uuid']})
        print(f"  Результат обновления договора: {result}")
    else:
        print(f"  ОШИБКА: Нет UUID партнера: {target_partner_ref}")
else:
    print(f"  ОШИБКА: Нет UUID договора: {contract_ref}")

# ========================================================
# ШАГ 6: Проверяем результат
# ========================================================
print()
print("=" * 60)
print("ШАГ 6: Проверяем результат")
print("=" * 60)

verify = query(
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "Д.Наименование КАК Наименование, "
    "Д.Партнер.Наименование КАК ПартнерНаименование, "
    "Д.Контрагент.Наименование КАК КонтрагентНаименование "
    "ИЗ Справочник.ДоговорыКонтрагентов КАК Д "
    "ГДЕ Д.Наименование = \"ДОГОВІР-ЗАЯВКА на організацію послуг з перевезення\""
)

if verify:
    v = verify[0]
    print(f"  Договор: {v['Наименование']}")
    print(f"  Контрагент: {v['КонтрагентНаименование']}")
    print(f"  Партнер: {v['ПартнерНаименование']}")
    if v['ПартнерНаименование'] != "Неизвестный партнер":
        print("  ✓ ТЕСТ ПРОЙДЕН: Партнер успешно заменён!")
    else:
        print("  ✗ ТЕСТ НЕ ПРОЙДЕН: Партнер всё ещё НеизвестныйПартнер")
else:
    print("  ОШИБКА: Договор не найден при проверке")
