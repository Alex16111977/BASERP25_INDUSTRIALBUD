"""Очистить ТЧ всех А_РасшифровкаЛистов от служебных листов и провести.

Удаляет строки с ИмяЛиста IN ('расчет 2025', 'расчет 2026', 'Метрики', 'PL_Свод').
Идемпотентен — повторный запуск не трогает уже очищенные документы (только перепроводит).
"""
import io, sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp, uuid_str


TO_REMOVE = {"расчет 2024", "расчет 2025", "расчет 2026", "Метрики", "PL_Свод"}


conn = connect_erp()
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ Ссылка, Номер, Дата, Проведен
ИЗ Документ.А_РасшифровкаЛистов
ГДЕ НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата
"""
tz = q.Выполнить().Выгрузить()
print(f"Документов А_РасшифровкаЛистов: {tz.Количество()}\n")

for i in range(tz.Количество()):
    row = tz.Получить(i)
    obj = row.Ссылка.ПолучитьОбъект()
    n_before = obj.Расшифровка.Количество()
    # Найти индексы к удалению
    to_del = []
    for j in range(n_before):
        if str(obj.Расшифровка.Получить(j).ИмяЛиста) in TO_REMOVE:
            to_del.append(j)
    if not to_del:
        # Просто перепровести (если документ не проведен)
        if not obj.Проведен:
            obj.Записать(conn.РежимЗаписиДокумента.Проведение)
            print(f"  №{str(obj.Номер).strip()}  {obj.Дата.strftime('%Y-%m-%d')}  строк={n_before} (cleanup не нужен, проведён)")
        else:
            print(f"  №{str(obj.Номер).strip()}  {obj.Дата.strftime('%Y-%m-%d')}  строк={n_before} (нечего удалять)")
        continue
    # Удалить с конца
    for j in reversed(to_del):
        obj.Расшифровка.Удалить(j)
    n_after = obj.Расшифровка.Количество()
    obj.Записать(conn.РежимЗаписиДокумента.Проведение)
    print(f"  №{str(obj.Номер).strip()}  {obj.Дата.strftime('%Y-%m-%d')}  удалено={len(to_del)}  было={n_before} → стало={n_after}")

print("\nDone.")
