# -*- coding: utf-8 -*-
"""Тест: списання ПрочиеРасходы ВНА для ВСІХ проведених АмортизацияОС.
Обробляє кожен документ послідовно (від найстарішого).
"""
import win32com.client, datetime
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# DRY_RUN = True  -> тільки показує
# DRY_RUN = False -> реально записує
DRY_RUN = False

# Отримати значення перерахувань
q0 = conn.NewObject("Запрос")
q0.Text = "ВЫБРАТЬ ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Р"
res0 = q0.Execute().Select()
res0.Next()
VD_RASHOD = res0.Р

# Знайти всі проведені АмортизацияОС
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Д.Ссылка, Д.Дата, Д.Номер, Д.Организация
ИЗ Документ.АмортизацияОС КАК Д
ГДЕ Д.Проведен = ИСТИНА
УПОРЯДОЧИТЬ ПО Д.Дата"""
docs = q.Execute().Unload()
print(f"Знайдено документів: {docs.Count()}")

for d in range(docs.Count()):
    doc_row = docs.Get(d)
    doc_ref = doc_row.Ссылка
    doc_date = doc_row.Дата
    doc_org = doc_row.Организация
    print(f"\n{'='*60}")
    print(f"Документ: {S(doc_ref)}")

    # Запит залишків ВНА на момент кінця місяця документа
    q3 = conn.NewObject("Запрос")
    q3.Text = """ВЫБРАТЬ
        О.Организация, О.Подразделение, О.НаправлениеДеятельности,
        О.НалоговоеНазначение, О.СтатьяРасходов, О.АналитикаРасходов,
        О.СуммаОстаток, О.СуммаБезНДСОстаток,
        О.СуммаРеглОстаток, О.СуммаРеглБезНДСОстаток,
        О.НДСРеглОстаток, О.СуммаУпрОстаток, О.СуммаУпрБезНДСОстаток
    ИЗ РегистрНакопления.ПрочиеРасходы.Остатки(
        &МВ,
        Организация = &Орг
        И СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА)
        И СтатьяРасходов.ВариантРаспределенияРасходовРегл = ЗНАЧЕНИЕ(Перечисление.ВариантыРаспределенияРасходов.НаВнеоборотныеАктивы)
    ) КАК О
    ГДЕ О.СуммаОстаток > 0"""

    # Граница = КонецМесяца(Дата), Включая
    end_month = doc_date.replace(day=28)
    # Простий спосіб: кінець місяця = перший день наступного - 1 секунда
    if doc_date.month == 12:
        end_month = datetime.datetime(doc_date.year, 12, 31, 23, 59, 59)
    else:
        end_month = datetime.datetime(doc_date.year, doc_date.month + 1, 1) - datetime.timedelta(seconds=1)

    boundary = conn.NewObject("Граница", end_month, conn.BoundaryType.Including)
    q3.SetParameter("МВ", boundary)
    q3.SetParameter("Орг", doc_org)

    try:
        tbl3 = q3.Execute().Unload()
    except Exception as e:
        print(f"  ПОМИЛКА запиту: {e}")
        continue

    if tbl3.Count() == 0:
        print(f"  Немає залишків ВНА (вже обнулено)")
        continue

    total = sum(tbl3.Get(i).СуммаОстаток for i in range(tbl3.Count()))
    print(f"  Залишків ВНА: {tbl3.Count()} рядків, сума: {total}")

    # Прочитати набір записів
    nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
    nabor.Filter.Регистратор.Set(doc_ref)
    nabor.Read()
    initial_count = nabor.Count()

    # Додати РАСХОД
    for i in range(tbl3.Count()):
        row = tbl3.Get(i)
        new_rec = nabor.Add()
        new_rec.RecordType = VD_RASHOD
        new_rec.Period = doc_date
        new_rec.Организация = row.Организация
        new_rec.Подразделение = row.Подразделение
        new_rec.НаправлениеДеятельности = row.НаправлениеДеятельности
        new_rec.НалоговоеНазначение = row.НалоговоеНазначение
        new_rec.СтатьяРасходов = row.СтатьяРасходов
        new_rec.АналитикаРасходов = row.АналитикаРасходов
        new_rec.Сумма = row.СуммаОстаток
        new_rec.СуммаБезНДС = row.СуммаБезНДСОстаток if row.СуммаБезНДСОстаток else 0
        new_rec.СуммаРегл = row.СуммаРеглОстаток if row.СуммаРеглОстаток else 0
        new_rec.СуммаРеглБезНДС = row.СуммаРеглБезНДСОстаток if row.СуммаРеглБезНДСОстаток else 0
        new_rec.НДСРегл = row.НДСРеглОстаток if row.НДСРеглОстаток else 0
        new_rec.СуммаУпр = row.СуммаУпрОстаток if row.СуммаУпрОстаток else 0
        new_rec.СуммаУпрБезНДС = row.СуммаУпрБезНДСОстаток if row.СуммаУпрБезНДСОстаток else 0

    print(f"  Записів: {initial_count} -> {nabor.Count()} (+{tbl3.Count()} РАСХОД)")

    if not DRY_RUN:
        try:
            nabor.Write(True)
            print(f"  ЗАПИСАНО!")
        except Exception as e:
            print(f"  ПОМИЛКА: {e}")
    else:
        print(f"  [DRY RUN]")

# Фінальна перевірка
print(f"\n{'='*60}")
print("ФІНАЛЬНА ПЕРЕВІРКА:")
q_check = conn.NewObject("Запрос")
q_check.Text = """ВЫБРАТЬ О.СуммаОстаток
ИЗ РегистрНакопления.ПрочиеРасходы.Остатки(,
    СтатьяРасходов.ТипРасходов = ЗНАЧЕНИЕ(Перечисление.ТипыРасходов.ФормированиеСтоимостиВНА)
    И СтатьяРасходов.ВариантРаспределенияРасходовРегл = ЗНАЧЕНИЕ(Перечисление.ВариантыРаспределенияРасходов.НаВнеоборотныеАктивы)
) КАК О
ГДЕ О.СуммаОстаток <> 0"""
res_check = q_check.Execute().Select()
has_balance = False
while res_check.Next():
    has_balance = True
    print(f"  Залишок: {res_check.СуммаОстаток}")
if not has_balance:
    print("  Всі залишки ВНА = 0 [OK]")
