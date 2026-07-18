# -*- coding: utf-8 -*-
"""Допроведення: заповнити АналитикаРасходов в ПрочиеРасходы = ОбъектЭксплуатации,
потім створити ПрочиеАктивыПассивы РАСХОД з аналітикою по кожному ОС."""
import win32com.client, datetime, copy
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

doc_ref = conn.Documents.АмортизацияОС.FindByNumber("000Ц-000001",
    datetime.datetime(2025, 12, 31, 23, 59, 59))
print(f"Документ: {S(doc_ref)}")
org = doc_ref.GetObject().Организация
print(f"Організація: {S(org)}")

# === 1. Отримати маппінг: сума -> ОбъектЭксплуатации ===
print("\n=== Маппінг з ПараметрыАмортизацииОСБУ ===")
q1 = conn.NewObject("Запрос")
q1.Text = """ВЫБРАТЬ
    Парам.ОсновноеСредство КАК ОС,
    Парам.СтоимостьДляВычисленияАмортизации КАК Сумма
ИЗ РегистрСведений.ПараметрыАмортизацииОСБУ.СрезПоследних(
    ДАТАВРЕМЯ(2025,12,31,23,59,59)) КАК Парам
ГДЕ Парам.СтоимостьДляВычисленияАмортизации > 0"""
tbl1 = q1.Execute().Unload()
print(f"ОС з амортизацією: {tbl1.Count()}")

asset_map = {}  # key = round(сума, 2) -> list of ОбъектЭксплуатации
for i in range(tbl1.Count()):
    r = tbl1.Get(i)
    key = round(float(r.Сумма), 2)
    if key not in asset_map:
        asset_map[key] = []
    asset_map[key].append(r.ОС)

print(f"Унікальних сум: {len(asset_map)}")
for k, v in list(asset_map.items())[:5]:
    print(f"  {k} -> {S(v[0])} ({len(v)} шт)")

# === 2. Оновити ПрочиеРасходы — заповнити АналитикаРасходов ===
print("\n=== Оновлення ПрочиеРасходы ===")
nabor_pr = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor_pr.Filter.Регистратор.Set(doc_ref)
nabor_pr.Read()
print(f"Записів: {nabor_pr.Count()}")

available = {}
for k, v in asset_map.items():
    available[k] = list(v)

fixed = 0
not_found = 0
for i in range(nabor_pr.Count()):
    rec = nabor_pr.Get(i)
    vid = S(conn.XMLString(rec.RecordType))
    if vid != "Receipt" or rec.СуммаРегл <= 0:
        continue
    if conn.ValueIsFilled(rec.АналитикаРасходов):
        fixed_already = S(rec.АналитикаРасходов)
        continue

    key = round(float(rec.СуммаРегл), 2)
    if key in available and len(available[key]) > 0:
        obj = available[key].pop(0)
        rec.АналитикаРасходов = obj
        fixed += 1
        print(f"  #{i}: {rec.СуммаРегл} -> {S(obj)}")
    else:
        not_found += 1
        print(f"  #{i}: {rec.СуммаРегл} — НЕ ЗНАЙДЕНО!")

print(f"\nВиправлено: {fixed}, не знайдено: {not_found}")
if fixed > 0:
    try:
        nabor_pr.Write()
        print("ПрочиеРасходы ЗАПИСАНО!")
    except Exception as e:
        print(f"ПОМИЛКА: {e}")

# === 3. Створити ПрочиеАктивыПассивы ===
print("\n=== ПрочиеАктивыПассивы ===")
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ
    ПР.Период, ПР.Организация, ПР.Подразделение,
    ПР.НаправлениеДеятельности, ПР.АналитикаРасходов КАК Аналитика,
    СУММА(ПР.СуммаРегл) КАК Сумма
ИЗ РегистрНакопления.ПрочиеРасходы КАК ПР
ГДЕ ПР.Регистратор = &Рег
    И ПР.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
    И ПР.СуммаРегл > 0
СГРУППИРОВАТЬ ПО
    ПР.Период, ПР.Организация, ПР.Подразделение,
    ПР.НаправлениеДеятельности, ПР.АналитикаРасходов"""
q.SetParameter("Рег", doc_ref)
tbl = q.Execute().Unload()

q0 = conn.NewObject("Запрос")
q0.Text = """ВЫБРАТЬ
    ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Р,
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ОсновныеСредства) КАК С"""
r0 = q0.Execute().Select()
r0.Next()

nabor_ap = conn.AccumulationRegisters.ПрочиеАктивыПассивы.CreateRecordSet()
nabor_ap.Filter.Регистратор.Set(doc_ref)
nabor_ap.Clear()

total = 0
with_anal = 0
for i in range(tbl.Count()):
    row = tbl.Get(i)
    rec = nabor_ap.Add()
    rec.RecordType = r0.Р
    rec.Period = row.Период
    rec.Организация = row.Организация
    rec.Подразделение = row.Подразделение
    rec.НаправлениеДеятельности = row.НаправлениеДеятельности
    rec.Статья = r0.С
    rec.Аналитика = row.Аналитика
    rec.Сумма = row.Сумма
    total += row.Сумма
    has_a = conn.ValueIsFilled(row.Аналитика)
    if has_a:
        with_anal += 1
    anal = S(row.Аналитика) if has_a else "<порожня>"
    print(f"  {S(row.Подразделение)} | {anal} | {row.Сумма}")

print(f"\nВсього: {total}, з аналітикою: {with_anal}/{tbl.Count()}")
try:
    nabor_ap.Write()
    print("ПрочиеАктивыПассивы ЗАПИСАНО!")
except Exception as e:
    print(f"ПОМИЛКА: {e}")
