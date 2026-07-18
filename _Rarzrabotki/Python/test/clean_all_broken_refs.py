# -*- coding: utf-8 -*-
"""
Универсальная очистка битых ссылок регистраторов во ВСЕХ регистрах накопления и регистрах сведений.
Проходит по метаданным, находит записи с Регистратор.Ссылка ЕСТЬ NULL, очищает наборы записей.
"""
import win32com.client
import sys

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("=" * 70)
print("Универсальная очистка битых ссылок регистраторов")
print("=" * 70)

print("\nПодключение к BaseERP...")
conn = v8.Connect(CONN_ERP)
print("OK\n")

query = conn.NewObject("Запрос")
total_cleaned = 0
total_errors = 0
results = []

# ============================================================
# 1. Регистры накопления
# ============================================================
print("=" * 70)
print("РЕГИСТРЫ НАКОПЛЕНИЯ")
print("=" * 70)

metadata = conn.Метаданные
accum_regs = metadata.РегистрыНакопления

for i in range(accum_regs.Количество()):
    reg_meta = accum_regs.Получить(i)
    reg_name = reg_meta.Имя
    full_name = f"РегистрНакопления.{reg_name}"

    # Проверяем количество осиротевших
    try:
        query.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ {full_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
        result = query.Execute().Выбрать()
        result.Следующий()
        orphan_count = result.Кол
    except:
        continue

    if orphan_count == 0:
        continue

    print(f"\n  {full_name}: {orphan_count} битых записей")

    # Получить уникальные битые регистраторы
    try:
        query.Text = f"ВЫБРАТЬ РАЗЛИЧНЫЕ Регистратор КАК Рег ИЗ {full_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
        result = query.Execute().Выбрать()

        orphan_refs = []
        while result.Следующий():
            orphan_refs.append(result.Рег)

        print(f"    Уникальных регистраторов: {len(orphan_refs)}")

        # Очистка
        reg_obj = getattr(conn.РегистрыНакопления, reg_name)
        cleaned = 0
        errors = 0
        for ref in orphan_refs:
            try:
                record_set = reg_obj.СоздатьНаборЗаписей()
                record_set.Отбор.Регистратор.Установить(ref)
                record_set.Записать()
                cleaned += 1
                if cleaned % 200 == 0:
                    print(f"    Очищено: {cleaned}...")
            except Exception as e:
                errors += 1
                if errors <= 2:
                    print(f"    ОШИБКА: {str(e)[:80]}")

        print(f"    Результат: {cleaned} OK, {errors} ошибок")
        total_cleaned += cleaned
        total_errors += errors
        results.append((full_name, orphan_count, cleaned, errors))

    except Exception as e:
        print(f"    ОШИБКА обработки: {str(e)[:80]}")

# ============================================================
# 2. Регистры сведений (только подчинённые регистратору)
# ============================================================
print("\n" + "=" * 70)
print("РЕГИСТРЫ СВЕДЕНИЙ (подчинённые регистратору)")
print("=" * 70)

info_regs = metadata.РегистрыСведений

for i in range(info_regs.Количество()):
    reg_meta = info_regs.Получить(i)
    reg_name = reg_meta.Имя
    full_name = f"РегистрСведений.{reg_name}"

    # Проверяем только регистры подчинённые регистратору
    try:
        mode = str(reg_meta.РежимЗаписи)
    except:
        continue

    if "ПодчинениеРегистратору" not in mode:
        continue

    # Проверяем количество осиротевших
    try:
        query.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ {full_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
        result = query.Execute().Выбрать()
        result.Следующий()
        orphan_count = result.Кол
    except:
        continue

    if orphan_count == 0:
        continue

    print(f"\n  {full_name}: {orphan_count} битых записей")

    try:
        query.Text = f"ВЫБРАТЬ РАЗЛИЧНЫЕ Регистратор КАК Рег ИЗ {full_name} ГДЕ Регистратор.Ссылка ЕСТЬ NULL"
        result = query.Execute().Выбрать()

        orphan_refs = []
        while result.Следующий():
            orphan_refs.append(result.Рег)

        print(f"    Уникальных регистраторов: {len(orphan_refs)}")

        reg_obj = getattr(conn.РегистрыСведений, reg_name)
        cleaned = 0
        errors = 0
        for ref in orphan_refs:
            try:
                record_set = reg_obj.СоздатьНаборЗаписей()
                record_set.Отбор.Регистратор.Установить(ref)
                record_set.Записать()
                cleaned += 1
                if cleaned % 200 == 0:
                    print(f"    Очищено: {cleaned}...")
            except Exception as e:
                errors += 1
                if errors <= 2:
                    print(f"    ОШИБКА: {str(e)[:80]}")

        print(f"    Результат: {cleaned} OK, {errors} ошибок")
        total_cleaned += cleaned
        total_errors += errors
        results.append((full_name, orphan_count, cleaned, errors))

    except Exception as e:
        print(f"    ОШИБКА обработки: {str(e)[:80]}")

# ============================================================
# Итоги
# ============================================================
print("\n" + "=" * 70)
print("ИТОГИ")
print("=" * 70)

if results:
    print(f"\n{'Регистр':<55} {'Было':>8} {'Очищено':>8} {'Ошибки':>7}")
    print("-" * 80)
    for name, count, cleaned, errors in results:
        print(f"  {name:<53} {count:>8} {cleaned:>8} {errors:>7}")
    print("-" * 80)
    print(f"  {'ИТОГО':<53} {sum(r[1] for r in results):>8} {total_cleaned:>8} {total_errors:>7}")
else:
    print("\nБитых ссылок не найдено!")

print(f"\nГОТОВО")
