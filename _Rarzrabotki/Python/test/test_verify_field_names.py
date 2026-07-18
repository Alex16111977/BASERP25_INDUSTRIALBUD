# -*- coding: utf-8 -*-
"""Перевірка: чи правильні імена полів у ПрочиеРасходы та яка фільтрація потрібна.
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# 1. Поля регістру ПрочиеРасходы
print("=== Поля регістру ПрочиеРасходы ===")
meta = conn.Metadata.AccumulationRegisters.ПрочиеРасходы
print("Виміри:")
for i in range(meta.Dimensions.Count()):
    d = meta.Dimensions.Get(i)
    print(f"  {d.Name}")
print("Ресурси:")
for i in range(meta.Resources.Count()):
    r = meta.Resources.Get(i)
    print(f"  {r.Name}")

# 2. Перевірка: чи є поле СтатьяРасходовСписания в регістрі?
print("\n=== Перевірка полів ===")
has_stat_rash = False
has_stat_rash_spis = False
for i in range(meta.Dimensions.Count()):
    d = meta.Dimensions.Get(i)
    if d.Name == "СтатьяРасходов":
        has_stat_rash = True
    if d.Name == "СтатьяРасходовСписания":
        has_stat_rash_spis = True
for i in range(meta.Attributes.Count()):
    a = meta.Attributes.Get(i)
    if a.Name == "СтатьяРасходовСписания":
        has_stat_rash_spis = True

print(f"  СтатьяРасходов в регістрі: {has_stat_rash}")
print(f"  СтатьяРасходовСписания в регістрі: {has_stat_rash_spis}")

# 3. Перевірка запиту з правильним полем
print("\n=== Тест запиту з СтатьяРасходов (правильне поле) ===")
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ
    ПР.СтатьяРасходов,
    ПР.АналитикаРасходов,
    ПР.СуммаОстаток
ИЗ
    РегистрНакопления.ПрочиеРасходы.Остатки(,
        СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА)
        И СтатьяРасходов.ВариантРаспределенияРасходовРегл = ЗНАЧЕНИЕ(Перечисление.ВариантыРаспределенияРасходов.НаВнеоборотныеАктивы)
    ) КАК ПР
ГДЕ ПР.СуммаОстаток > 0"""
try:
    res = q.Execute()
    sel = res.Select()
    cnt = 0
    while sel.Next():
        cnt += 1
        if cnt <= 5:
            print(f"  #{cnt} Стаття={S(sel.СтатьяРасходов)} Аналітика={S(sel.АналитикаРасходов)} Сума={sel.СуммаОстаток}")
    print(f"  Всього: {cnt}")
except Exception as e:
    print(f"  ПОМИЛКА: {e}")

# 4. Перевірка умови ГДЕ для фільтрації
print("\n=== Умова ГДЕ для фільтрації ===")
print("Правильна умова:")
print("  СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА)")
print("  І СтатьяРасходов.ВариантРаспределенияРасходовРегл = ЗНАЧЕНИЕ(Перечисление.ВариантыРаспределенияРасходов.НаВнеоборотныеАктивы)")
print()
print("НЕ правильна (помилка):")
print("  СтатьяРасходовСписания.ТипРасходов = ... <-- поля СтатьяРасходовСписания НЕМАЄ в регістрі!")

print("\n=== DONE ===")
