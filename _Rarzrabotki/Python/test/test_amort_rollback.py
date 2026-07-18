# -*- coding: utf-8 -*-
"""Відкат: видалити РАСХОД записи, які були помилково додані як ПРИХОД.
Повернути набір записей АмортизацияОС 000Ц-000001 від 31.12.2025 до стандартного стану (28 записів).
"""
import win32com.client, datetime
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# 1. Знайти документ
q = conn.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка
ИЗ Документ.АмортизацияОС КАК Д
ГДЕ Д.Проведен = ИСТИНА
УПОРЯДОЧИТЬ ПО Д.Дата"""
res = q.Execute().Select()
res.Next()
doc_ref = res.Ссылка
print(f"Документ: {S(doc_ref)}")

# 2. Прочитати набір записей
nabor = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor.Filter.Регистратор.Set(doc_ref)
nabor.Read()
print(f"Записів в наборі: {nabor.Count()}")

# 3. Залишити тільки перші 28 (оригінальні)
# Видалити записи з індексу 28 і далі (додані нами)
to_delete = nabor.Count() - 28
if to_delete > 0:
    for i in range(to_delete):
        nabor.Delete(28)  # завжди видаляємо 28-й (наступний після оригінальних)
    print(f"Видалено {to_delete} записів. Залишилось: {nabor.Count()}")
    nabor.Write(True)
    print("Записано!")
else:
    print("Нічого видаляти — вже 28 або менше записів.")

# 4. Перевірка
nabor2 = conn.AccumulationRegisters.ПрочиеРасходы.CreateRecordSet()
nabor2.Filter.Регистратор.Set(doc_ref)
nabor2.Read()
print(f"Перевірка: записів = {nabor2.Count()}")
