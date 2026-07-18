# -*- coding: utf-8 -*-
"""
Интеграционный тест: вызов ЗаполнитьНачисленияИзБазЗП() через COM.
Использует существующий документ НачислениеЗарплаты за декабрь 2025.
"""
import win32com.client
import datetime
import sys

sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("=" * 60)
print("Подключение к ERP...")
conn = v8.Connect(CONN_ERP)
print("[OK] Подключение к ERP успешно")

# Найдём последний документ НачислениеЗарплаты
q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 "
    "  Док.Ссылка КАК Ссылка "
    "ИЗ "
    "  Документ.НачислениеЗарплаты КАК Док "
    "УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
)

res = q.Execute().Choose()
if not res.Next():
    print("[ОШИБКА] Не найден документ НачислениеЗарплаты за декабрь 2025!")
    conn = None
    sys.exit(1)

doc_ref = res.Ссылка
print(f"[OK] Найден документ: {doc_ref}")

# Получаем объект документа
doc_obj = doc_ref.GetObject()
print(f"  Дата: {doc_obj.Дата}")
print(f"  Номер: {doc_obj.Номер}")
print(f"  А_Обработан: {doc_obj.А_Обработан}")

# Сбрасываем А_Обработан чтобы процедура отработала
doc_obj.А_Обработан = False
print()
print("=" * 60)
print("Вызов ЗаполнитьНачисленияИзБазЗП()...")
print("=" * 60)

try:
    doc_obj.ЗаполнитьНачисленияИзБазЗП()
    print("[OK] Процедура выполнена успешно!")
except Exception as e:
    print(f"[ОШИБКА] {e}")
    conn = None
    sys.exit(1)

# Проверяем результаты
print()
print("=" * 60)
print("РЕЗУЛЬТАТЫ:")
print("=" * 60)
print(f"  А_Начисления:    {doc_obj.А_Начисления.Count()} строк")
print(f"  А_Удержания:     {doc_obj.А_Удержания.Count()} строк")
print(f"  А_НачисленияБух: {doc_obj.А_НачисленияБух.Count()} строк")
print(f"  Начисления:      {doc_obj.Начисления.Count()} строк")
print(f"  Удержания:       {doc_obj.Удержания.Count()} строк")
print(f"  ФизическиеЛица:  {doc_obj.ФизическиеЛица.Count()} строк")
print(f"  А_Обработан:     {doc_obj.А_Обработан}")

# Проверим первые 3 строки Начисления
print()
print("Примеры строк Начисления:")
for i in range(min(3, doc_obj.Начисления.Count())):
    row = doc_obj.Начисления.Get(i)
    print(f"  [{i}] Сотрудник={row.Сотрудник}, Начисление={row.Начисление}, Результат={row.Результат}")

# Записываем документ
print()
print("Запись документа...")
try:
    doc_obj.Write(conn.Documents.НачислениеЗарплаты.GetRef().Metadata().WriteMode.Запись)
except Exception:
    # Пробуем другой способ записи
    try:
        write_mode = conn.Enums.РежимЗаписиДокумента.Запись if hasattr(conn, 'Enums') else None
        if write_mode:
            doc_obj.Write(write_mode)
        else:
            doc_obj.Write()
        print("[OK] Документ записан")
    except Exception as e2:
        print(f"[ПРЕДУПРЕЖДЕНИЕ] Запись: {e2}")
        print("  (Документ заполнен в памяти, но не записан)")

print()
print("=" * 60)
print("ТЕСТ ЗАВЕРШЁН")
print("=" * 60)

conn = None
