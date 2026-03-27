# -*- coding: utf-8 -*-
"""Перевіряємо чи можна створити елемент в ГруппыФинансовогоУчетаДоходовРасходов
та ГруппыФинансовогоУчетаВнеоборотныхАктивов через COM.
НЕ СТВОРЮЄМО — тільки перевіряємо структуру і можливість.
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

S = conn.String

# 1. Перевірка структури ГруппыФинансовогоУчетаДоходовРасходов
print("=" * 60)
print("СТРУКТУРА: ГруппыФинансовогоУчетаДоходовРасходов")
print("=" * 60)
meta1 = conn.Metadata.Catalogs.ГруппыФинансовогоУчетаДоходовРасходов
print(f"  Наименование: {meta1.Synonym}")
print(f"  Иерархический: {meta1.Hierarchical}")
print(f"  КоличествоРеквизитов: {meta1.Attributes.Count()}")
for i in range(meta1.Attributes.Count()):
    attr = meta1.Attributes.Get(i)
    print(f"    {attr.Name}: {attr.Type}")
print(f"  КоличествоТЧ: {meta1.TabularSections.Count()}")
for i in range(meta1.TabularSections.Count()):
    ts = meta1.TabularSections.Get(i)
    print(f"    ТЧ: {ts.Name} ({ts.Attributes.Count()} колонок)")
    for j in range(ts.Attributes.Count()):
        a = ts.Attributes.Get(j)
        print(f"      {a.Name}: {a.Type}")

# 2. Перевірка структури ГруппыФинансовогоУчетаВнеоборотныхАктивов
print("\n" + "=" * 60)
print("СТРУКТУРА: ГруппыФинансовогоУчетаВнеоборотныхАктивов")
print("=" * 60)
meta2 = conn.Metadata.Catalogs.ГруппыФинансовогоУчетаВнеоборотныхАктивов
print(f"  Наименование: {meta2.Synonym}")
print(f"  Иерархический: {meta2.Hierarchical}")
print(f"  КоличествоРеквизитов: {meta2.Attributes.Count()}")
for i in range(meta2.Attributes.Count()):
    attr = meta2.Attributes.Get(i)
    print(f"    {attr.Name}: {attr.Type}")
print(f"  КоличествоТЧ: {meta2.TabularSections.Count()}")
for i in range(meta2.TabularSections.Count()):
    ts = meta2.TabularSections.Get(i)
    print(f"    ТЧ: {ts.Name} ({ts.Attributes.Count()} колонок)")

# 3. Спроба створити тестовий елемент (DRY RUN)
print("\n" + "=" * 60)
print("DRY RUN: спроба створити об'єкт (без запису)")
print("=" * 60)
try:
    obj = conn.Catalogs.ГруппыФинансовогоУчетаДоходовРасходов.CreateItem()
    obj.Description = "ТЕСТ (не зберігати)"
    print(f"  CreateItem() = OK")
    print(f"  Наименование = {obj.Description}")
    # НЕ зберігаємо!
    print("  Об'єкт створено в пам'яті (НЕ записано в базу)")
except Exception as e:
    print(f"  ПОМИЛКА: {e}")

try:
    obj2 = conn.Catalogs.ГруппыФинансовогоУчетаВнеоборотныхАктивов.CreateItem()
    obj2.Description = "ТЕСТ НА (не зберігати)"
    print(f"  CreateItem() НА = OK")
    print(f"  Наименование = {obj2.Description}")
    print("  Об'єкт створено в пам'яті (НЕ записано в базу)")
except Exception as e:
    print(f"  ПОМИЛКА НА: {e}")

print("\n=== DONE ===")
