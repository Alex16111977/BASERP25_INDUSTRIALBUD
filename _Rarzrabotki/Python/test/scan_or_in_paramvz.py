# -*- coding: utf-8 -*-
"""
Сканер: проверить во всех документах с операциями ВозвратОплатыКлиенту /
ВозвратДенежныхСредствОтПоставщика, есть ли в их ManagerModule.bsl функция
ПараметрыВзаиморасчеты со структурой ОбъектРасчетов = "Объект.ОбъектРасчетов".

Если нет — это ошибка типового для конкретной хоз. операции.
"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

DOCS_DIR = r"C:\Configuration_downloads\BASERP25\Documents"

ВОЗВРАТЫ = [
    "ВозвратОплатыКлиенту",
    "ВозвратДенежныхСредствОтПоставщика",
    "ВозвратДенежныхСредствВДругуюОрганизацию",
]

# Найти все документы которые упоминают эти операции в Manager или ObjectModule
print("=" * 110, flush=True)
print("СКАН: ПараметрыВзаиморасчеты в документах с возвратными операциями", flush=True)
print("=" * 110, flush=True)

# Список документов где есть возвратные операции
доки_с_возвратом = []
for имя_папки in os.listdir(DOCS_DIR):
    путь = os.path.join(DOCS_DIR, имя_папки, "Ext", "ManagerModule.bsl")
    if not os.path.exists(путь):
        continue
    try:
        with open(путь, encoding='utf-8') as f:
            content = f.read()
    except Exception:
        try:
            with open(путь, encoding='cp1251') as f:
                content = f.read()
        except Exception:
            continue
    # Какие возвратные операции встречаются
    возвраты_в_док = [op for op in ВОЗВРАТЫ if op in content]
    if not возвраты_в_док:
        continue
    доки_с_возвратом.append((имя_папки, путь, возвраты_в_док, content))

print(f"\nНайдено документов с возвратными операциями: {len(доки_с_возвратом)}\n", flush=True)

# Проверка каждого
results = []
for имя, путь, возвраты, content in доки_с_возвратом:
    # Есть ли функция ПараметрыВзаиморасчеты
    has_func = re.search(r"Функция\s+ПараметрыВзаиморасчеты", content) is not None
    # Есть ли строка СтруктураПараметров.ОбъектРасчетов = "Объект.ОбъектРасчетов" (или подобная)
    or_pattern_in_header = re.search(r'СтруктураПараметров\.ОбъектРасчетов\s*=\s*"Объект\.ОбъектРасчетов"', content) is not None
    # Альтернативно — общий поиск ОбъектРасчетов = в функции
    any_or_path = re.search(r'СтруктураПараметров\.ОбъектРасчетов\s*=\s*"[^"]+"', content) is not None

    status = ""
    if not has_func:
        status = "НЕТ ФУНКЦИИ ПараметрыВзаиморасчеты"
    elif or_pattern_in_header:
        status = "✓ ЕСТЬ структура ОбъектРасчетов=Объект.ОбъектРасчетов"
    elif any_or_path:
        # Найдём что именно присвоено
        match = re.search(r'СтруктураПараметров\.ОбъектРасчетов\s*=\s*"([^"]+)"', content)
        путь_ор = match.group(1) if match else "?"
        status = f"⚠ ОбъектРасчетов = \"{путь_ор}\" (другой путь)"
    else:
        status = "❌ НЕТ СтруктураПараметров.ОбъектРасчетов в коде ВООБЩЕ"

    # Проверка: упоминание возвратных операций в контексте функции ПараметрыВзаиморасчеты
    # Найти область функции ПараметрыВзаиморасчеты
    blocks_for_returns = {}
    for op in возвраты:
        # ищем "Если ... ХозяйственныеОперации." + op в коде функции
        if has_func:
            # Найдём блок функции (грубо)
            func_match = re.search(
                r"Функция\s+ПараметрыВзаиморасчеты.*?(КонецФункции)",
                content,
                re.DOTALL,
            )
            if func_match:
                func_block = func_match.group(0)
                if op in func_block:
                    blocks_for_returns[op] = True
                else:
                    blocks_for_returns[op] = False
            else:
                blocks_for_returns[op] = "?"
        else:
            blocks_for_returns[op] = "—"

    results.append({
        'имя': имя,
        'статус': status,
        'возвраты': возвраты,
        'в_парамвз': blocks_for_returns,
    })

# Вывод
for r in results:
    print(f"📄 {r['имя']}", flush=True)
    print(f"   {r['статус']}", flush=True)
    print(f"   Возвратные операции в коде: {', '.join(r['возвраты'])}", flush=True)
    if r['в_парамвз']:
        for op, есть in r['в_парамвз'].items():
            mark = "✓" if есть is True else ("✗" if есть is False else "?")
            print(f"     {mark} {op} в ф-ции ПараметрыВзаиморасчеты", flush=True)
    print("", flush=True)

# Сводка
print("\n" + "=" * 110, flush=True)
print("СВОДКА: документы с ошибкой (нет СтруктураПараметров.ОбъектРасчетов = Объект.ОбъектРасчетов)", flush=True)
print("=" * 110, flush=True)
bad = [r for r in results if "ЕСТЬ структура" not in r['статус']]
print(f"Всего проблемных: {len(bad)} из {len(results)}", flush=True)
for r in bad:
    print(f"  ❌ {r['имя']:<45} | {r['статус']}", flush=True)
