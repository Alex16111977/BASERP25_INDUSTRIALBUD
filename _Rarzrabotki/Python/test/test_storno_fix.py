# -*- coding: utf-8 -*-
"""Фікс: відкатити тестові записи і знайти правильний спосіб створити РАСХОД."""
import win32com.client, datetime
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# 1. Знайти документ
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

# 2. Прочитати набір і подивитись RecordType
nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor.Filter.Регистратор.Set(doc_ref)
nabor.Read()
print(f"\nЗаписів: {nabor.Count()}")
for i in range(nabor.Count()):
    rec = nabor.Get(i)
    print(f"  #{i+1} RecordType={rec.RecordType} Сума={rec.Сумма} Стаття={S(rec.СтатьяРасходов)}")

# 3. Видалити зайві записи (залишити тільки перші 2 оригінальні ПРИХОД)
print(f"\n--- Відновлення: залишаємо тільки перші 2 записи ---")
while nabor.Count() > 2:
    nabor.Delete(nabor.Count() - 1)
print(f"Записів після очистки: {nabor.Count()}")
nabor.Write(True)  # True = замінити все
print("Записано (замінено)")

# 4. Перевірити RecordType значення
print(f"\n--- Перевірка RecordType ---")
nabor2 = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor2.Filter.Регистратор.Set(doc_ref)
nabor2.Read()
for i in range(nabor2.Count()):
    rec = nabor2.Get(i)
    rt = rec.RecordType
    print(f"  RecordType={rt}, type={type(rt)}, XMLString={conn.XMLString(rt)}")

# 5. Спробувати створити РАСХОД правильно
print(f"\n--- Спроба створити РАСХОД ---")
new_rec = nabor2.Add()
# Спробуємо різні способи
try:
    # Спосіб 1: через XMLСтрока
    расход = conn.XMLValue(conn.NewObject("XMLТипЗнч", "AccumulationRecordType"), "Expense")
    new_rec.RecordType = расход
    print(f"  Спосіб XMLValue: OK, RecordType={new_rec.RecordType}")
except Exception as e:
    print(f"  Спосіб XMLValue: ПОМИЛКА {e}")

# Видалити тестовий запис
nabor2.Delete(nabor2.Count() - 1)

# Спосіб 2: через eval
try:
    q_eval = conn.NewObject("Запрос")
    q_eval.Text = "ВЫБРАТЬ ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК Знач"
    res_eval = q_eval.Execute().Select()
    res_eval.Next()
    rashod_value = res_eval.Знач
    print(f"  Значення Расход через запит: {rashod_value}, type={type(rashod_value)}, XML={conn.XMLString(rashod_value)}")

    new_rec2 = nabor2.Add()
    new_rec2.RecordType = rashod_value
    print(f"  Спосіб запит: RecordType={new_rec2.RecordType}, XML={conn.XMLString(new_rec2.RecordType)}")

    # Перевіримо чи це дійсно Расход
    new_rec2.Период = nabor2.Get(0).Период
    new_rec2.Организация = nabor2.Get(0).Организация
    new_rec2.Подразделение = nabor2.Get(0).Подразделение
    new_rec2.СтатьяРасходов = nabor2.Get(0).СтатьяРасходов
    new_rec2.АналитикаРасходов = nabor2.Get(0).АналитикаРасходов
    new_rec2.Сумма = nabor2.Get(0).Сумма

    print(f"  Новий запис: RecordType XML = {conn.XMLString(new_rec2.RecordType)}")

    # Видалити тестовий
    nabor2.Delete(nabor2.Count() - 1)
except Exception as e:
    print(f"  Спосіб запит: ПОМИЛКА {e}")

print("\n=== DONE ===")
