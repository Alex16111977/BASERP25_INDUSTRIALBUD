# -*- coding: utf-8 -*-
"""Тест v2: правильне допроведення РАСХОД в ПрочиеРасходы."""
import win32com.client, datetime
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Отримати значення Расход через запит
q_rt = conn.NewObject("Запрос")
q_rt.Text = "ВЫБРАТЬ ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Знач"
sel_rt = q_rt.Execute().Select()
sel_rt.Next()
РАСХОД = sel_rt.Знач
print(f"Расход XML = {conn.XMLString(РАСХОД)}")

# Знайти документ
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.ВнутреннееПотреблениеТоваров КАК Д
ГДЕ Д.Номер ПОДОБНО &Н И Д.Дата >= &Д1 И Д.Дата <= &Д2"""
q.SetParameter("Н", "%000Ц-000009%")
q.SetParameter("Д1", datetime.datetime(2025, 12, 11))
q.SetParameter("Д2", datetime.datetime(2025, 12, 11, 23, 59, 59))
res = q.Execute().Select()
res.Next()
doc_ref = res.Ссылка
print(f"Документ: {S(doc_ref)}")

# Прочитати ПРИХОД записи
q2 = conn.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ ПР.Период, ПР.Организация, ПР.Подразделение,
    ПР.НаправлениеДеятельности, ПР.НалоговоеНазначение,
    ПР.СтатьяРасходов, ПР.АналитикаРасходов,
    ПР.Сумма, ПР.СуммаБезНДС, ПР.СуммаРегл, ПР.СуммаРеглБезНДС,
    ПР.НДСРегл, ПР.СуммаУпр, ПР.СуммаУпрБезНДС,
    ПР.ХозяйственнаяОперация, ПР.АналитикаУчетаНоменклатуры
ИЗ РегистрНакопления.ПрочиеРасходы КАК ПР
ГДЕ ПР.Регистратор = &Рег И ПР.Активность
    И ПР.СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА)
    И ПР.СтатьяРасходов.ВариантРаспределенияРасходовРегл = ЗНАЧЕНИЕ(Перечисление.ВариантыРаспределенияРасходов.НаВнеоборотныеАктивы)"""
q2.SetParameter("Рег", doc_ref)
tbl = q2.Execute().Unload()

# Порахувати ПРИХОД і РАСХОД
prihod = []
rashod_cnt = 0
for i in range(tbl.Count()):
    row = tbl.Get(i)
    # Всі записи тут — ПРИХОД (фільтр в запиті не перевіряє ВидДвижения,
    # а RegisterRecords повертає тільки ПРИХОД бо РАСХОД ще не створені)
    prihod.append(row)

print(f"\nПРИХОД записів: {len(prihod)}")
for i, row in enumerate(prihod):
    print(f"  #{i+1} {S(row.АналитикаРасходов)} = {row.Сумма}")

# Прочитати набір записів
nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor.Filter.Регистратор.Set(doc_ref)
nabor.Read()
print(f"\nНабір ДО: {nabor.Count()} записів")

# Видалити старі РАСХОД (якщо є від попередніх запусків)
indices_to_del = []
for i in range(nabor.Count()):
    rec = nabor.Get(i)
    if conn.XMLString(rec.RecordType) == "Expense":
        stat = rec.СтатьяРасходов
        if conn.ValueIsFilled(stat):
            try:
                if S(stat.ТипРасходов) == "Формирование стоимости необоротных активов":
                    indices_to_del.append(i)
            except:
                pass

if indices_to_del:
    print(f"Видаляємо {len(indices_to_del)} старих РАСХОД")
    for idx in reversed(indices_to_del):
        nabor.Delete(idx)

# Додати РАСХОД для кожного ПРИХОД
for row in prihod:
    new_rec = nabor.Add()
    new_rec.RecordType = РАСХОД
    new_rec.Period = row.Период
    new_rec.Организация = row.Организация
    new_rec.Подразделение = row.Подразделение
    new_rec.НаправлениеДеятельности = row.НаправлениеДеятельности
    new_rec.НалоговоеНазначение = row.НалоговоеНазначение
    new_rec.СтатьяРасходов = row.СтатьяРасходов
    new_rec.АналитикаРасходов = row.АналитикаРасходов
    new_rec.Сумма = row.Сумма
    new_rec.СуммаБезНДС = row.СуммаБезНДС if row.СуммаБезНДС else 0
    new_rec.СуммаРегл = row.СуммаРегл if row.СуммаРегл else 0
    new_rec.СуммаРеглБезНДС = row.СуммаРеглБезНДС if row.СуммаРеглБезНДС else 0
    new_rec.НДСРегл = row.НДСРегл if row.НДСРегл else 0
    new_rec.СуммаУпр = row.СуммаУпр if row.СуммаУпр else 0
    new_rec.СуммаУпрБезНДС = row.СуммаУпрБезНДС if row.СуммаУпрБезНДС else 0
    new_rec.ХозяйственнаяОперация = row.ХозяйственнаяОперация
    new_rec.АналитикаУчетаНоменклатуры = row.АналитикаУчетаНоменклатуры
    print(f"  + РАСХОД: {S(row.АналитикаРасходов)} = {row.Сумма} (XML: {conn.XMLString(new_rec.RecordType)})")

print(f"Набір ПІСЛЯ: {nabor.Count()} записів")

# Записати
try:
    nabor.Write(True)  # True = замінити ВСІ записи регістратора
    print("ЗАПИСАНО УСПІШНО!")
except Exception as e:
    print(f"ПОМИЛКА: {e}")
    exit()

# Перевірка
print("\n=== ПЕРЕВІРКА ===")
nabor2 = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor2.Filter.Регистратор.Set(doc_ref)
nabor2.Read()
p = 0
r = 0
for i in range(nabor2.Count()):
    rec = nabor2.Get(i)
    xml = conn.XMLString(rec.RecordType)
    if xml == "Receipt":
        p += 1
    else:
        r += 1
    if conn.ValueIsFilled(rec.СтатьяРасходов):
        try:
            tp = S(rec.СтатьяРасходов.ТипРасходов)
            if "необоротных" in tp:
                print(f"  {xml}: {S(rec.АналитикаРасходов)} = {rec.Сумма}")
        except:
            pass

print(f"\nВсього: ПРИХОД={p} РАСХОД={r}")
print(f"{'OK!' if p >= r and r > 0 else 'НЕ OK!'}")

# Залишок
q4 = conn.NewObject("Запрос")
q4.Text = """ВЫБРАТЬ ПР.АналитикаРасходов, ПР.СуммаОстаток
ИЗ РегистрНакопления.ПрочиеРасходы.Остатки(,
    АналитикаРасходов В (ВЫБРАТЬ РАЗЛИЧНЫЕ Т.АналитикаРасходов ИЗ РегистрНакопления.ПрочиеРасходы КАК Т ГДЕ Т.Регистратор = &Рег И Т.СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА))
    И СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА)
) КАК ПР"""
q4.SetParameter("Рег", doc_ref)
try:
    sel4 = q4.Execute().Select()
    while sel4.Next():
        print(f"  Залишок: {S(sel4.АналитикаРасходов)} = {sel4.СуммаОстаток}")
except Exception as e:
    print(f"  Помилка залишку: {e}")

print("\n=== DONE ===")
