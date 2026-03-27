# -*- coding: utf-8 -*-
"""Тест: створення РАСХОД в ПрочиеАктивыПассивы (Статья=ОсновныеСредства)
для всіх проведених АмортизацияОС.
Сума = СуммаРегл з ПрочиеРасходы (бухгалтерська амортизація, БЕЗ ПДВ).
"""
import win32com.client, datetime
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

DRY_RUN = True

# Значення перерахувань
q0 = conn.NewObject("Запрос")
q0.Text = """ВЫБРАТЬ
    ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Расход,
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ОсновныеСредства) КАК СтатьяОС"""
res0 = q0.Execute().Select()
res0.Next()
VD_RASHOD = res0.Расход
STATYA_OS = res0.СтатьяОС
print(f"Статья ОС: {S(STATYA_OS)}")

# Поточний залишок ОсновныеСредства
q_bal = conn.NewObject("Запрос")
q_bal.Text = """ВЫБРАТЬ АП.СуммаОстаток
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Остатки(,
    Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ОсновныеСредства)
) КАК АП"""
res_bal = q_bal.Execute().Select()
if res_bal.Next():
    print(f"Поточний залишок ОсновныеСредства: {res_bal.СуммаОстаток}")
else:
    print("Залишок ОсновныеСредства = 0")

# Знайти всі проведені АмортизацияОС
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Д.Ссылка, Д.Дата, Д.Организация
ИЗ Документ.АмортизацияОС КАК Д
ГДЕ Д.Проведен УПОРЯДОЧИТЬ ПО Д.Дата"""
docs = q.Execute().Unload()
print(f"\nДокументів АмортизацияОС: {docs.Count()}")

total_amort = 0

for d in range(docs.Count()):
    doc_row = docs.Get(d)
    doc_ref = doc_row.Ссылка
    doc_date = doc_row.Дата
    print(f"\n{'='*50}")
    print(f"Док: {S(doc_ref)}")

    # Запитати ПРИХОД ПрочиеРасходы (СуммаРегл > 0) — ВСІ записи
    q3 = conn.NewObject("Запрос")
    q3.Text = """ВЫБРАТЬ
        ПР.Период КАК Период,
        ПР.Организация КАК Организация,
        ПР.Подразделение КАК Подразделение,
        ПР.НаправлениеДеятельности КАК НаправлениеДеятельности,
        ПР.АналитикаРасходов КАК Аналитика,
        СУММА(ПР.СуммаРегл) КАК Сумма
    ИЗ РегистрНакопления.ПрочиеРасходы КАК ПР
    ГДЕ ПР.Регистратор = &Рег
        И ПР.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        И ПР.СуммаРегл > 0
    СГРУППИРОВАТЬ ПО
        ПР.Период, ПР.Организация, ПР.Подразделение,
        ПР.НаправлениеДеятельности, ПР.АналитикаРасходов"""
    q3.SetParameter("Рег", doc_ref)
    tbl3 = q3.Execute().Unload()

    if tbl3.Count() == 0:
        print("  Немає ПРИХОД з СуммаРегл > 0")
        continue

    doc_sum = sum(tbl3.Get(i).Сумма for i in range(tbl3.Count()))
    total_amort += doc_sum
    print(f"  Записів: {tbl3.Count()}, сума амортизації (регл): {doc_sum}")

    # Перевірити чи вже є рухи ПрочиеАктивыПассивы
    nabor = conn.AccumulationRegisters.ПрочиеАктивыПассивы.CreateRecordSet()
    nabor.Filter.Регистратор.Set(doc_ref)
    nabor.Read()
    existing = nabor.Count()
    print(f"  Існуючих записів ПрочиеАктивыПассивы: {existing}")

    # Додати РАСХОД
    for i in range(tbl3.Count()):
        row = tbl3.Get(i)
        new_rec = nabor.Add()
        new_rec.RecordType = VD_RASHOD
        new_rec.Period = row.Период
        new_rec.Организация = row.Организация
        new_rec.Подразделение = row.Подразделение
        new_rec.НаправлениеДеятельности = row.НаправлениеДеятельности
        new_rec.Статья = STATYA_OS
        new_rec.Аналитика = row.Аналитика
        new_rec.Сумма = row.Сумма

    print(f"  Після: {nabor.Count()} (+{tbl3.Count()} РАСХОД)")

    if not DRY_RUN:
        try:
            nabor.Write(True)
            print("  ЗАПИСАНО!")
        except Exception as e:
            print(f"  ПОМИЛКА: {e}")
    else:
        print("  [DRY RUN]")

print(f"\n{'='*50}")
print(f"ВСЬОГО амортизація (регл): {total_amort}")

if not DRY_RUN:
    print("\nЗАЛИШОК ПІСЛЯ:")
    res_bal2 = q_bal.Execute().Select()
    if res_bal2.Next():
        print(f"  ОсновныеСредства: {res_bal2.СуммаОстаток}")
    else:
        print("  ОсновныеСредства: 0")
