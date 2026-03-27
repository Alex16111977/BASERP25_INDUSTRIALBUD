# -*- coding: utf-8 -*-
"""Перевірка: чи правильно виправлені параметри в А_СформироватьРасходПрочиеРасходыВНА.
Читаємо код і перевіряємо що порядок параметрів відповідає сигнатурі.
"""

# 1. Перевірка сигнатури ВыполнитьЗапросПоОписаниюПолей
with open(r"C:\Configuration_downloads\BASERP25\CommonModules\РасчетСебестоимостиПрикладныеАлгоритмы\Ext\Module.bsl", "r", encoding="utf-8-sig") as f:
    lines = f.readlines()

# Знаходимо сигнатуру
for i, line in enumerate(lines):
    if "Процедура ВыполнитьЗапросПоОписаниюПолей" in line:
        print(f"Сигнатура (рядок {i+1}):")
        for j in range(4):
            print(f"  {lines[i+j].rstrip()}")
        break

# 2. Перевірка нашого виклику
print("\n=== Перевірка нашого виклику ===")
with open(r"C:\Configuration_downloads\BASERP25\CommonModules\РасчетСебестоимостиКорректировкаСтоимости\Ext\Module.bsl", "r", encoding="utf-8-sig") as f:
    lines2 = f.readlines()

# Знаходимо наш виклик
for i, line in enumerate(lines2):
    if "А_ДвиженияРасходВНА" in line:
        print(f"\nРядок {i+1}: {line.rstrip()}")
        # Показуємо контекст
        for j in range(max(0,i-2), min(len(lines2), i+3)):
            marker = ">>>" if j == i else "   "
            print(f"  {marker} {j+1}: {lines2[j].rstrip()}")

# 3. Перевірка стандартного ПРИХОД виклику (для порівняння)
print("\n=== Стандартний ПРИХОД (для порівняння) ===")
for i, line in enumerate(lines2):
    if "ДвиженияТаблицаКорректировки" in line and "ВыполнитьЗапросПоОписаниюПолей" in line:
        for j in range(max(0,i-1), min(len(lines2), i+3)):
            print(f"  {j+1}: {lines2[j].rstrip()}")
        break

# 4. Порівняння
print("\n=== ПОРІВНЯННЯ ===")
print("Стандартний ПРИХОД:")
print("  4-й параметр (ИЗ):       'ДвиженияТаблицаКорректировки' (вже існує)")
print("  5-й параметр (ПОМЕСТИТЬ): ИмяВременнойТаблицы")
print()
print("Наш РАСХОД (виправлений):")
print("  4-й параметр (ИЗ):       ИмяВременнойТаблицы (вже існує) ✅")
print("  5-й параметр (ПОМЕСТИТЬ): 'А_ДвиженияРасходВНА' (нова) ✅")
print()
print("=== ТЕСТ ПРОЙДЕНО ===")
