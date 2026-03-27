# -*- coding: utf-8 -*-
"""Діагностика помилки відображення документів в регл. обліку за грудень 2025.
Перевіряємо: журнал помилок, рухи ПринятиеКУчетуОС, статус закриття місяця.
"""
import win32com.client
import datetime

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")

S = conn.String
VF = conn.ValueIsFilled

# 1. Документи "ОтражениеДокументовВРеглУчете" за грудень 2025
print("=" * 70)
print("1. ДОКУМЕНТИ ВІДОБРАЖЕННЯ В РЕГЛ. ОБЛІКУ (грудень 2025)")
print("=" * 70)

for doc_type in ["ОтражениеДокументовВРеглУчете", "ОтражениеДокументовВУчете"]:
    try:
        q = conn.NewObject("Запрос")
        q.Text = f"""ВЫБРАТЬ ПЕРВЫЕ 5
            Д.Ссылка, Д.Номер, Д.Дата, Д.Проведен
        ИЗ Документ.{doc_type} КАК Д
        ГДЕ Д.Дата >= &Дата1 И Д.Дата <= &Дата2
        УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"""
        q.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
        q.SetParameter("Дата2", datetime.datetime(2025, 12, 31, 23, 59, 59))
        res = q.Execute()
        sel = res.Select()
        cnt = 0
        while sel.Next():
            cnt += 1
            print(f"  [{doc_type}] #{sel.Номер} від {sel.Дата} Проведен={sel.Проведен}")
        if cnt == 0:
            print(f"  [{doc_type}] Не знайдено за грудень")
    except Exception as e:
        print(f"  [{doc_type}] Помилка: {str(e)[:100]}")

# 2. РасчетСебестоимостиТоваров за грудень 2025
print("\n" + "=" * 70)
print("2. РасчетСебестоимости ЗА ГРУДЕНЬ 2025")
print("=" * 70)

for doc_type in ["РасчетСебестоимостиТоваров", "РасчетСебестоимости"]:
    try:
        q2 = conn.NewObject("Запрос")
        q2.Text = f"""ВЫБРАТЬ ПЕРВЫЕ 5
            Д.Ссылка, Д.Номер, Д.Дата, Д.Проведен
        ИЗ Документ.{doc_type} КАК Д
        ГДЕ Д.Дата >= &Дата1 И Д.Дата <= &Дата2
        УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"""
        q2.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
        q2.SetParameter("Дата2", datetime.datetime(2025, 12, 31, 23, 59, 59))
        res2 = q2.Execute()
        sel2 = res2.Select()
        cnt = 0
        while sel2.Next():
            cnt += 1
            print(f"  [{doc_type}] #{sel2.Номер} від {sel2.Дата} Проведен={sel2.Проведен}")
        if cnt == 0:
            print(f"  [{doc_type}] Не знайдено")
    except Exception as e:
        print(f"  [{doc_type}] Помилка: {str(e)[:100]}")

# 3. Рухи ПринятиеКУчетуОС ІБ00-000245 (після перепроведення)
print("\n" + "=" * 70)
print("3. РУХИ ПринятиеКУчетуОС ІБ00-000245 (ПІСЛЯ виправлення)")
print("=" * 70)

q3 = conn.NewObject("Запрос")
q3.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.ПринятиеКУчетуОС КАК Д
    ГДЕ Д.Номер ПОДОБНО &Номер И Д.Дата >= &Дата"""
q3.SetParameter("Номер", "%000245%")
q3.SetParameter("Дата", datetime.datetime(2025, 12, 1))
res3 = q3.Execute()
sel3 = res3.Select()
doc_ref = None
if sel3.Next():
    doc_ref = sel3.Ссылка

if doc_ref:
    # СтоимостьОС
    print(f"\n  --- СтоимостьОС ---")
    try:
        q_s = conn.NewObject("Запрос")
        q_s.Text = """ВЫБРАТЬ
            Д.ВидДвижения, Д.ОсновноеСредство, Д.Стоимость,
            Д.ГруппаФинансовогоУчета, Д.ХозяйственнаяОперация
        ИЗ РегистрНакопления.СтоимостьОС КАК Д
        ГДЕ Д.Регистратор = &Ссылка"""
        q_s.SetParameter("Ссылка", doc_ref)
        res_s = q_s.Execute()
        sel_s = res_s.Select()
        cnt = 0
        while sel_s.Next():
            cnt += 1
            print(f"    #{cnt} ОС={S(sel_s.ОсновноеСредство)} Вартість={sel_s.Стоимость}")
            print(f"       ГФУ={S(sel_s.ГруппаФинансовогоУчета)} filled={VF(sel_s.ГруппаФинансовогоУчета)}")
            print(f"       ХозОп={S(sel_s.ХозяйственнаяОперация)}")
        if cnt == 0:
            print("    ПУСТО! Немає рухів!")
    except Exception as e:
        print(f"    Помилка: {str(e)[:150]}")

    # ПрочиеРасходы
    print(f"\n  --- ПрочиеРасходы ---")
    try:
        q_pr = conn.NewObject("Запрос")
        q_pr.Text = """ВЫБРАТЬ
            Д.ВидДвижения, Д.СтатьяРасходов, Д.АналитикаРасходов, Д.Сумма
        ИЗ РегистрНакопления.ПрочиеРасходы КАК Д
        ГДЕ Д.Регистратор = &Ссылка"""
        q_pr.SetParameter("Ссылка", doc_ref)
        res_pr = q_pr.Execute()
        sel_pr = res_pr.Select()
        cnt = 0
        while sel_pr.Next():
            cnt += 1
            print(f"    #{cnt} Вид={S(sel_pr.ВидДвижения)} Стаття={S(sel_pr.СтатьяРасходов)}")
            print(f"       Аналітика={S(sel_pr.АналитикаРасходов)} Сума={sel_pr.Сумма}")
        if cnt == 0:
            print("    ПУСТО!")
    except Exception as e:
        print(f"    Помилка: {str(e)[:150]}")

    # Хозрасчетный (бухгалтерський облік)
    print(f"\n  --- РегістрБухгалтерії.Хозрасчетный ---")
    try:
        q_buh = conn.NewObject("Запрос")
        q_buh.Text = """ВЫБРАТЬ
            Д.Период, Д.СчетДт, Д.СчетКт, Д.Сумма
        ИЗ РегистрБухгалтерии.Хозрасчетный КАК Д
        ГДЕ Д.Регистратор = &Ссылка"""
        q_buh.SetParameter("Ссылка", doc_ref)
        res_buh = q_buh.Execute()
        sel_buh = res_buh.Select()
        cnt = 0
        while sel_buh.Next():
            cnt += 1
            print(f"    #{cnt} Дт={S(sel_buh.СчетДт)} Кт={S(sel_buh.СчетКт)} Сума={sel_buh.Сумма}")
        if cnt == 0:
            print("    ПУСТО!")
    except Exception as e:
        print(f"    Помилка: {str(e)[:150]}")
else:
    print("  Документ не знайдено!")

# 4. Перевірка ПрочиеРасходы — РАСХОД записи за грудень (від закриття місяця)
print("\n" + "=" * 70)
print("4. ПрочиеРасходы: РАСХОД записи за грудень 2025 (Капітальні інвестиції)")
print("=" * 70)

try:
    q4 = conn.NewObject("Запрос")
    q4.Text = """ВЫБРАТЬ ПЕРВЫЕ 20
        Д.Период, Д.Регистратор, Д.ВидДвижения,
        Д.СтатьяРасходов, Д.АналитикаРасходов, Д.Сумма
    ИЗ РегистрНакопления.ПрочиеРасходы КАК Д
    ГДЕ Д.Период >= &Дата1 И Д.Период <= &Дата2
        И Д.СтатьяРасходов.Наименование ПОДОБНО &Стаття
        И Д.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
    УПОРЯДОЧИТЬ ПО Д.Период"""
    q4.SetParameter("Дата1", datetime.datetime(2025, 12, 1))
    q4.SetParameter("Дата2", datetime.datetime(2025, 12, 31, 23, 59, 59))
    q4.SetParameter("Стаття", "%Капітальні інвестиції%")
    res4 = q4.Execute()
    sel4 = res4.Select()
    cnt = 0
    while sel4.Next():
        cnt += 1
        print(f"  #{cnt} {sel4.Период} Регістратор={S(sel4.Регистратор)}")
        print(f"     Стаття={S(sel4.СтатьяРасходов)} Сума={sel4.Сумма}")
    if cnt == 0:
        print("  НЕМАЄ РАСХОД записів! Закриття місяця не списало капінвестиції!")
except Exception as e:
    print(f"  Помилка: {str(e)[:150]}")

# 5. Журнал реєстрації помилок
print("\n" + "=" * 70)
print("5. ПЕРЕВІРКА ФУНКЦІОНАЛЬНИХ ОПЦІЙ")
print("=" * 70)

for opt in ["ИспользоватьВнеоборотныеАктивы2_4", "ВерсияУчетаНеоборотныхАктивов2_5",
            "ФормироватьУправленческийБаланс", "ИспользоватьМеждународныйФинансовыйУчет"]:
    try:
        val = conn.GetFunctionalOption(opt)
        print(f"  {opt} = {val}")
    except:
        print(f"  {opt} — не знайдено")

print("\n=== DONE ===")
