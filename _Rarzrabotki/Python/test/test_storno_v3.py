# -*- coding: utf-8 -*-
"""Тест v3: чистий тест РАСХОД в ПрочиеРасходы для документа 000Ц-000009.
Крок 1: Перевірити поточний стан
Крок 2: Якщо РАСХОД вже є — видалити (очистка)
Крок 3: Створити РАСХОД заново
Крок 4: Перевірити результат
"""
import win32com.client, datetime
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Знайти документ
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.ВнутреннееПотреблениеТоваров КАК Д
ГДЕ Д.Номер ПОДОБНО &Н И Д.Дата >= &Д1 И Д.Дата <= &Д2"""
q.SetParameter("Н", "%000Ц-000009%")
q.SetParameter("Д1", datetime.datetime(2025, 12, 11))
q.SetParameter("Д2", datetime.datetime(2025, 12, 11, 23, 59, 59))
res = q.Execute().Select()
if not res.Next():
    print("Документ не знайдено!")
    exit()
doc_ref = res.Ссылка
print(f"Документ: {S(doc_ref)}")

# Отримати значення Расход
q_rt = conn.NewObject("Запрос")
q_rt.Text = "ВЫБРАТЬ ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Знач"
sel_rt = q_rt.Execute().Select()
sel_rt.Next()
RASHOD = sel_rt.Знач
print(f"Расход XML = {conn.XMLString(RASHOD)}")

# === КРОК 1: Поточний стан ===
print("\n=== КРОК 1: Поточний стан ===")
nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor.Filter.Регистратор.Set(doc_ref)
nabor.Read()
print(f"Записів: {nabor.Count()}")
prihod_vna = []
rashod_vna = []
other = []
for i in range(nabor.Count()):
    rec = nabor.Get(i)
    xml_rt = conn.XMLString(rec.RecordType)
    is_vna = False
    try:
        stat = rec.СтатьяРасходов
        if conn.ValueIsFilled(stat):
            tp = S(stat.ТипРасходов)
            if "необоротных" in tp.lower():
                is_vna = True
    except:
        pass

    if is_vna and xml_rt == "Receipt":
        prihod_vna.append(i)
        print(f"  #{i+1} ПРИХОД ВНА: {S(rec.АналитикаРасходов)} = {rec.Сумма}")
    elif is_vna and xml_rt == "Expense":
        rashod_vna.append(i)
        print(f"  #{i+1} РАСХОД ВНА: {S(rec.АналитикаРасходов)} = {rec.Сумма}")
    else:
        other.append(i)
        print(f"  #{i+1} {xml_rt} інше: {S(rec.СтатьяРасходов)} = {rec.Сумма}")

print(f"\nПРИХОД ВНА: {len(prihod_vna)}, РАСХОД ВНА: {len(rashod_vna)}, Інше: {len(other)}")

# === КРОК 2: Видалити старі РАСХОД ВНА (якщо є) ===
if rashod_vna:
    print(f"\n=== КРОК 2: Видаляємо {len(rashod_vna)} старих РАСХОД ВНА ===")
    for idx in reversed(rashod_vna):
        nabor.Delete(idx)
    nabor.Write(True)
    print("Видалено і записано")

    # Перечитати
    nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
    nabor.Filter.Регистратор.Set(doc_ref)
    nabor.Read()
    print(f"Записів після очистки: {nabor.Count()}")
else:
    print("\n=== КРОК 2: РАСХОД ВНА немає, очистка не потрібна ===")

# === КРОК 3: Створити РАСХОД ===
print(f"\n=== КРОК 3: Створюємо РАСХОД ===")

# Зібрати ПРИХОД записи для сторно
prihod_data = []
for i in range(nabor.Count()):
    rec = nabor.Get(i)
    xml_rt = conn.XMLString(rec.RecordType)
    if xml_rt != "Receipt":
        continue
    is_vna = False
    try:
        stat = rec.СтатьяРасходов
        if conn.ValueIsFilled(stat):
            tp = S(stat.ТипРасходов)
            if "необоротных" in tp.lower():
                is_vna = True
    except:
        pass
    if is_vna:
        prihod_data.append({
            'Период': rec.Период,
            'Организация': rec.Организация,
            'Подразделение': rec.Подразделение,
            'НаправлениеДеятельности': rec.НаправлениеДеятельности,
            'НалоговоеНазначение': rec.НалоговоеНазначение,
            'СтатьяРасходов': rec.СтатьяРасходов,
            'АналитикаРасходов': rec.АналитикаРасходов,
            'Сумма': rec.Сумма,
            'СуммаБезНДС': rec.СуммаБезНДС,
            'СуммаРегл': rec.СуммаРегл,
            'СуммаРеглБезНДС': rec.СуммаРеглБезНДС,
            'НДСРегл': rec.НДСРегл,
            'СуммаУпр': rec.СуммаУпр,
            'СуммаУпрБезНДС': rec.СуммаУпрБезНДС,
            'ХозяйственнаяОперация': rec.ХозяйственнаяОперация,
            'АналитикаУчетаНоменклатуры': rec.АналитикаУчетаНоменклатуры,
        })

print(f"ПРИХОД ВНА для сторно: {len(prihod_data)}")

if not prihod_data:
    print("Нічого сторнувати!")
    exit()

# Додати РАСХОД записи
for data in prihod_data:
    new_rec = nabor.Add()
    new_rec.RecordType = RASHOD
    for key, val in data.items():
        setattr(new_rec, key, val if val is not None else 0)
    print(f"  + РАСХОД: {S(data['АналитикаРасходов'])} = {data['Сумма']} (RT XML={conn.XMLString(new_rec.RecordType)})")

print(f"Записів в наборі: {nabor.Count()}")

# Записати — True = замінити ВСІ записи цього регістратора
nabor.Write(True)
print("ЗАПИСАНО!")

# === КРОК 4: Перевірка ===
print(f"\n=== КРОК 4: Перевірка ===")
nabor2 = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor2.Filter.Регистратор.Set(doc_ref)
nabor2.Read()
p_vna = 0
r_vna = 0
for i in range(nabor2.Count()):
    rec = nabor2.Get(i)
    xml_rt = conn.XMLString(rec.RecordType)
    is_vna = False
    try:
        stat = rec.СтатьяРасходов
        if conn.ValueIsFilled(stat):
            tp = S(stat.ТипРасходов)
            if "необоротных" in tp.lower():
                is_vna = True
    except:
        pass
    if is_vna:
        if xml_rt == "Receipt":
            p_vna += 1
        elif xml_rt == "Expense":
            r_vna += 1
        print(f"  {xml_rt}: {S(rec.АналитикаРасходов)} = {rec.Сумма}")

print(f"\nПРИХОД ВНА: {p_vna}, РАСХОД ВНА: {r_vna}")
print(f"Баланс: {'✅ OK (ПРИХОД = РАСХОД)' if p_vna == r_vna and r_vna > 0 else '❌ НЕ OK!'}")

print("\n=== DONE ===")
