# -*- coding: utf-8 -*-
"""
Тест 3: поиск Партнёра и Контрагента для подразделений дисбаланса
документа ІБ00-008058. Проверяет, что:
  - для каждого подразделения найдётся ровно 1 Партнёр с А_Подразделение = ...
  - для каждого Партнёра найдётся Контрагент
  - пара (КонтрагентДебитор, КонтрагентКредитор) однозначна для пары подразделений
"""
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
conn = v8.Connect(CONN)
print("Connected\n")

# Подразделения из ІБ00-008058 (4 штуки)
TARGET_NAMES = ["Строительство", "Спецтехника", "Логистика", "ЦО"]


def find_podr_by_name(name):
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 3
        С.Ссылка КАК Ссылка,
        С.Наименование КАК Наименование
    ИЗ Справочник.СтруктураПредприятия КАК С
    ГДЕ С.Наименование = &Имя И НЕ С.ПометкаУдаления
    """
    q.УстановитьПараметр("Имя", name)
    sel = q.Выполнить().Выбрать()
    results = []
    while sel.Следующий():
        results.append(sel.Ссылка)
    return results


def find_partner_by_podr(podr_ref):
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        П.Ссылка КАК Партнер,
        П.Наименование КАК Наименование,
        П.А_Подразделение КАК Подразделение
    ИЗ Справочник.Партнеры КАК П
    ГДЕ П.А_Подразделение = &Подр И НЕ П.ПометкаУдаления
    """
    q.УстановитьПараметр("Подр", podr_ref)
    sel = q.Выполнить().Выбрать()
    res = []
    while sel.Следующий():
        res.append({"ref": sel.Партнер, "name": sel.Наименование})
    return res


def find_contragents_by_partner(partner_ref):
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        К.Ссылка КАК Контрагент,
        К.Наименование КАК Наименование
    ИЗ Справочник.Контрагенты КАК К
    ГДЕ К.Партнер = &Партнер И НЕ К.ПометкаУдаления
    """
    q.УстановитьПараметр("Партнер", partner_ref)
    sel = q.Выполнить().Выбрать()
    res = []
    while sel.Следующий():
        res.append({"ref": sel.Контрагент, "name": sel.Наименование})
    return res


print("=" * 70)
print("Подразделения → Партнёр → Контрагент")
print("=" * 70)

podr_contragent = {}  # имя подр -> ref контрагента
for name in TARGET_NAMES:
    podrs = find_podr_by_name(name)
    if len(podrs) == 0:
        print(f"  [!] {name}: подразделение НЕ НАЙДЕНО")
        continue
    if len(podrs) > 1:
        print(f"  [!] {name}: найдено {len(podrs)} подразделений с таким именем — неоднозначность")
    podr_ref = podrs[0]

    partners = find_partner_by_podr(podr_ref)
    if len(partners) == 0:
        print(f"  [!] {name}: Партнёр с А_Подразделение НЕ НАЙДЕН")
        continue
    if len(partners) > 1:
        print(f"  [!] {name}: найдено {len(partners)} партнёров — неоднозначность")
    partner = partners[0]

    contragents = find_contragents_by_partner(partner["ref"])
    if len(contragents) == 0:
        print(f"  [!] {name}: Контрагент для партнёра '{partner['name']}' НЕ НАЙДЕН")
        continue

    c = contragents[0]
    podr_contragent[name] = c["ref"]
    print(f"  [+] {name:<20s} -> Партнёр '{partner['name']}' -> Контрагент '{c['name']}'")
    if len(contragents) > 1:
        print(f"       (у партнёра {len(contragents)} контрагентов, возьмём первого)")

print()
print("=" * 70)
print("Проверка уникальности пар (Дт, Кт) для жадного алгоритма")
print("=" * 70)
# 3 пары: Строительство<-ЦО, Спецтехника<-ЦО, Логистика<-ЦО
pairs = [("Строительство", "ЦО"), ("Спецтехника", "ЦО"), ("Логистика", "ЦО")]
seen = set()
ok = True
for dt, kt in pairs:
    if dt not in podr_contragent or kt not in podr_contragent:
        print(f"  [!] Пара ({dt}, {kt}): нет контрагента")
        ok = False
        continue
    key = (str(podr_contragent[dt]), str(podr_contragent[kt]))
    if key in seen:
        print(f"  [!] Пара ({dt}, {kt}): ДУБЛИКАТ контрагентов")
        ok = False
    else:
        seen.add(key)
        print(f"  [+] Пара ({dt}, {kt}): уникальные контрагенты")

print()
print(f"ИТОГ: {'OK' if ok else 'FAIL'}")
