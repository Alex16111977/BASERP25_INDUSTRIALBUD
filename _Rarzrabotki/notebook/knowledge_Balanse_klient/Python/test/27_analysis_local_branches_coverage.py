# -*- coding: utf-8 -*-
"""
СКРИПТ 27 (Phase 1) — Coverage локальных веток для ВозвратОплатыКлиенту в 3 ManagerModule

ЧТО ДЕЛАЕТ:
    Через regex проверяет наличие ветки ВозвратОплатыКлиенту в функции
    ТекстЗапросаТаблицаПрочиеАктивыПассивы для:
    - СписаниеБезналичныхДенежныхСредств
    - РасходныйКассовыйОрдер
    - ОперацияПоПлатежнойКарте
"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
from _common import save_csv

CFG = "C:/Configuration_downloads/BASERP25"
DOCS = (
    "СписаниеБезналичныхДенежныхСредств",
    "РасходныйКассовыйОрдер",
    "ОперацияПоПлатежнойКарте",
)


def analyze(doc_name):
    path = f"{CFG}/Documents/{doc_name}/Ext/ManagerModule.bsl"
    if not os.path.exists(path):
        return {"Документ": doc_name, "Найдено": False, "Покрытие": "файл не найден"}
    with open(path, encoding="utf-8") as f:
        src = f.read()

    # Найти функцию ТекстЗапросаТаблицаПрочиеАктивыПассивы
    m = re.search(
        r"Функция ТекстЗапросаТаблицаПрочиеАктивыПассивы\b[^\n]*\n(.+?)(?=\n(?:Функция |Процедура ))",
        src, re.DOTALL,
    )
    if not m:
        return {"Документ": doc_name, "Найдено": False, "Покрытие": "функция не найдена"}
    body = m.group(1)
    body_len = len(body.splitlines())

    has_branch = "ВозвратОплатыКлиенту" in body
    has_zad_kl = "ЗадолженностьКлиентов" in body
    has_poluch_av = "ПолученныеАвансы" in body

    return {
        "Документ": doc_name,
        "Найдено": True,
        "Строк_в_функции": body_len,
        "Имеет_ВозвратОплатыКлиенту": has_branch,
        "Имеет_ЗадолженностьКлиентов": has_zad_kl,
        "Имеет_ПолученныеАвансы": has_poluch_av,
        "Покрытие": (
            "ПОЛНОЕ" if has_branch and has_zad_kl else
            "ЧАСТИЧНОЕ" if has_zad_kl else
            "ОТСУТСТВУЕТ"
        ),
    }


print("=" * 100)
print("СКРИПТ 27 — Coverage локальных веток для ВозвратОплатыКлиенту")
print("=" * 100)

rows = []
for d in DOCS:
    r = analyze(d)
    rows.append(r)
    print(f"\n  {d}:")
    for k, v in r.items():
        if k != "Документ":
            print(f"      {k}: {v}")

path = save_csv("27_local_branches", rows,
                ["Документ", "Найдено", "Строк_в_функции",
                 "Имеет_ВозвратОплатыКлиенту", "Имеет_ЗадолженностьКлиентов",
                 "Имеет_ПолученныеАвансы", "Покрытие"])
print(f"\nАртефакт: {path}")
