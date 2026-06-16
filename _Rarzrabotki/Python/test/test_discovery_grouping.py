# -*- coding: utf-8 -*-
# Task 4: НайтиПоСсылкам -> определить контрагент-контекст каждого источника, сгруппировать
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()
P = c.neizv(erp)
arr = erp.NewObject("Массив")
arr.Добавить(P)
res = erp.НайтиПоСсылкам(arr)


def try_attr(obj, name):
    try:
        v = getattr(obj, name)
        return v
    except Exception:
        return None


def контрагент_контекст(данные, мета):
    # Контрагенты сами себе контекст
    if мета == "Справочник.Контрагенты":
        return данные
    # Прямой реквизит Контрагент (документы, ключи, ОР, договоры)
    k = try_attr(данные, "Контрагент")
    if k is not None and erp.ЗначениеЗаполнено(k):
        # Контрагент может быть типа Организации — нас интересует только Контрагенты
        return k
    return None


# Сначала разберём, что такое row.Данные по типам метаданных
print("=== типы row.Данные по метаданным ===")
seen_types = {}
n = res.Количество()
for i in range(n):
    row = res.Получить(i)
    try:
        мета = row.Метаданные.ПолноеИмя()
    except Exception:
        мета = "?"
    данные = row.Данные
    tname = str(type(данные).__name__)
    # для COM объектов попробуем узнать, ссылка это или запись
    descr = ""
    try:
        descr = "ref?" + ("да" if erp.ЗначениеЗаполнено(данные) else "нет")
    except Exception:
        descr = "не-ссылка"
    key = (мета, descr)
    seen_types[key] = seen_types.get(key, 0) + 1
for (m, d), cnt in sorted(seen_types.items(), key=lambda x: -x[1]):
    print(f"  {cnt:5d}  {m:45} | Данные: {d}")

# Группировка по контрагенту
print("\n=== группировка по контрагенту ===")
groups = {}
no_ctx = {}
for i in range(n):
    row = res.Получить(i)
    try:
        мета = row.Метаданные.ПолноеИмя()
    except Exception:
        мета = "?"
    данные = row.Данные
    k = контрагент_контекст(данные, мета)
    if k is None or not erp.ЗначениеЗаполнено(k):
        no_ctx[мета] = no_ctx.get(мета, 0) + 1
        continue
    try:
        key = c.uid(erp, k)
        nm = k.Наименование
    except Exception:
        key, nm = "?", "?"
    g = groups.setdefault(key, {"наим": nm, "всего": 0, "мета": {}})
    g["всего"] += 1
    g["мета"][мета] = g["мета"].get(мета, 0) + 1

for key, v in sorted(groups.items(), key=lambda x: -x[1]["всего"]):
    print(f"  {v['наим']:45} всего={v['всего']:4}  {v['мета']}")
print("\nГрупп(контрагентов):", len(groups))
print("Без контрагент-контекста:", no_ctx)
