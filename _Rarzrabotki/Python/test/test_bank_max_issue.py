# -*- coding: utf-8 -*-
"""Діагностика: перевірити частоту А_Подразделение для договору 5/0725/18.
МАКСИМУМ() обирає за внутрішнім ID, а не за частотою.
"""
import win32com.client
import datetime
from collections import Counter

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

# 1. Знайти договір
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Спр.Ссылка ИЗ Справочник.ДоговорыКонтрагентов КАК Спр
    ГДЕ Спр.Наименование ПОДОБНО &Наим"""
q.SetParameter("Наим", "%5/0725/18%")
res = q.Execute()
sel = res.Select()
sel.Next()
contract_ref = sel.Ссылка
print(f"Договір: {conn.String(contract_ref)}")
print(f"  Поточне А_ПодразделениеБанк: '{conn.String(contract_ref.А_ПодразделениеБанк)}'")
print(f"  Поточне А_НаправлениеДеятельностиБанк: '{conn.String(contract_ref.А_НаправлениеДеятельностиБанк)}'")

# 2. Всі банковські документи для цього договору
print("\n" + "=" * 70)
print("ВСІ БАНКОВСЬКІ ДОКУМЕНТИ ДЛЯ ЦЬОГО ДОГОВОРУ")
print("=" * 70)

podrazd_counter = Counter()
napr_counter = Counter()

for doc_type in ["СписаниеБезналичныхДенежныхСредств", "ПоступлениеБезналичныхДенежныхСредств"]:
    q2 = conn.NewObject("Запрос")
    q2.Text = """ВЫБРАТЬ
        Док.Номер КАК Номер,
        Док.Дата КАК ДатаДок,
        Док.А_Подразделение КАК А_Подразделение,
        Док.А_НаправлениеДеятельности КАК А_НаправлениеДеятельности
    ИЗ
        Документ.""" + doc_type + """.РасшифровкаПлатежа КАК РП
        ЛЕВОЕ СОЕДИНЕНИЕ Документ.""" + doc_type + """ КАК Док
        ПО РП.Ссылка = Док.Ссылка
    ГДЕ
        Док.Проведен
        И РП.А_Договор = &Договор
        И Док.Дата >= &Дата1"""
    q2.SetParameter("Договор", contract_ref)
    q2.SetParameter("Дата1", datetime.datetime(2025, 12, 1))

    res2 = q2.Execute()
    sel2 = res2.Select()
    print(f"\n  --- {doc_type} ---")
    count = 0
    while sel2.Next():
        count += 1
        p = conn.String(sel2.А_Подразделение)
        n = conn.String(sel2.А_НаправлениеДеятельности)
        if conn.ValueIsFilled(sel2.А_Подразделение):
            podrazd_counter[p] += 1
        if conn.ValueIsFilled(sel2.А_НаправлениеДеятельности):
            napr_counter[n] += 1
        if count <= 5:
            print(f"    #{count} Док #{sel2.Номер} от {sel2.ДатаДок}")
            print(f"       А_Подразделение: '{p}'")
            print(f"       А_НаправлениеДеятельности: '{n}'")
    if count > 5:
        print(f"    ... ще {count - 5} документів")
    print(f"    Всього: {count}")

# 3. Результат
print("\n" + "=" * 70)
print("ЧАСТОТА ПІДРОЗДІЛІВ (А_Подразделение)")
print("=" * 70)
for podrazd, cnt in podrazd_counter.most_common():
    marker = " <-- НАЙЧАСТІШЕ" if cnt == podrazd_counter.most_common(1)[0][1] else ""
    print(f"  '{podrazd}': {cnt} документів{marker}")

print("\n" + "=" * 70)
print("ЧАСТОТА НАПРЯМКІВ (А_НаправлениеДеятельности)")
print("=" * 70)
for napr, cnt in napr_counter.most_common():
    marker = " <-- НАЙЧАСТІШЕ" if cnt == napr_counter.most_common(1)[0][1] else ""
    print(f"  '{napr}': {cnt} документів{marker}")

print("\n=== ВИСНОВОК ===")
if podrazd_counter:
    most_common_p = podrazd_counter.most_common(1)[0]
    print(f"  Найчастіше Подразделение: '{most_common_p[0]}' ({most_common_p[1]} разів)")
if napr_counter:
    most_common_n = napr_counter.most_common(1)[0]
    print(f"  Найчастіше Направление: '{most_common_n[0]}' ({most_common_n[1]} разів)")

print("\n=== DONE ===")
