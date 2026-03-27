# -*- coding: utf-8 -*-
"""
Диагностика: какие поля количества есть в ДвиженияССубконто корреспондентного регистра.
Проверяем: Количество, КоличествоДт, КоличествоКт
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'
DATE_FROM = datetime(2025, 12, 1)
DATE_TO = datetime(2025, 12, 31, 23, 59, 59)


def make_time(dt):
    return pywintypes.Time(dt)


def connect_buhbud():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Подключение к BuhBud")
    return conn


def test_field(conn, field_name, label):
    """Проверяет доступность поля в ДвиженияССубконто"""
    print(f"\n--- Проверка поля: {field_name} ---")
    query = conn.NewObject("Запрос")
    query.Текст = (
        f"ВЫБРАТЬ ПЕРВЫЕ 3\n"
        f"  ПРЕДСТАВЛЕНИЕ(Д.Регистратор) КАК Рег,\n"
        f"  Д.Сумма,\n"
        f"  Д.{field_name} КАК ПолеКол\n"
        f"ИЗ\n"
        f"  РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(&Н, &К, , , ) КАК Д\n"
        f"ГДЕ Д.Сумма <> 0"
    )
    query.УстановитьПараметр("Н", make_time(DATE_FROM))
    query.УстановитьПараметр("К", make_time(DATE_TO))
    try:
        result = query.Выполнить().Выбрать()
        count = 0
        while result.Следующий():
            count += 1
            print(f"  [{count}] {label}={result.ПолеКол}  Сумма={result.Сумма}  Рег={str(result.Рег)[:60]}")
        print(f"  [OK] Поле '{field_name}' СУЩЕСТВУЕТ! Записей: {count}")
        return True
    except Exception as e:
        err_msg = str(e)
        if "Поле" in err_msg or "Field" in err_msg or "не найдено" in err_msg:
            print(f"  [НЕТ] Поле '{field_name}' НЕ СУЩЕСТВУЕТ: {err_msg[:100]}")
        else:
            print(f"  [ОШИБКА] {err_msg[:150]}")
        return False


def test_raw_table_fields(conn):
    """Проверяем поля в raw таблице (не виртуальной)"""
    print("\n=== RAW таблица РегистрБухгалтерии.Хозрасчетный ===")
    for field in ["Количество", "КоличествоДт", "КоличествоКт"]:
        query = conn.NewObject("Запрос")
        query.Текст = (
            f"ВЫБРАТЬ ПЕРВЫЕ 2\n"
            f"  ПРЕДСТАВЛЕНИЕ(Д.Регистратор) КАК Рег,\n"
            f"  Д.Сумма,\n"
            f"  Д.{field} КАК ПолеКол\n"
            f"ИЗ\n"
            f"  РегистрБухгалтерии.Хозрасчетный КАК Д\n"
            f"ГДЕ Д.Сумма <> 0"
        )
        query.УстановитьПараметр("Н", make_time(DATE_FROM))
        query.УстановитьПараметр("К", make_time(DATE_TO))
        try:
            result = query.Выполнить().Выбрать()
            count = 0
            while result.Следующий():
                count += 1
                print(f"  RAW [{field}]={result.ПолеКол}  Сумма={result.Сумма}")
            print(f"  [OK] RAW поле '{field}' СУЩЕСТВУЕТ")
        except Exception as e:
            print(f"  [НЕТ] RAW поле '{field}': {str(e)[:100]}")


def test_specific_document(conn):
    """Проверяем конкретный документ ІБ00-007653 — все числовые поля"""
    print("\n=== Документ ІБ00-007653 — все числовые ресурсы ===")

    # Сначала попробуем с Количество, КоличествоДт, КоличествоКт
    for fields_str, label in [
        ("Д.Сумма, Д.КоличествоДт, Д.КоличествоКт", "КоличествоДт/Кт"),
        ("Д.Сумма, Д.Количество", "Количество"),
    ]:
        print(f"\n--- {label} ---")
        query = conn.NewObject("Запрос")
        query.Текст = (
            f"ВЫБРАТЬ\n"
            f"  ПРЕДСТАВЛЕНИЕ(Д.Регистратор) КАК Рег,\n"
            f"  Д.СчетДт,\n"
            f"  Д.СчетКт,\n"
            f"  {fields_str}\n"
            f"ИЗ\n"
            f"  РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(&Н, &К, , , ) КАК Д\n"
            f"ГДЕ\n"
            f"  ПРЕДСТАВЛЕНИЕ(Д.Регистратор) ПОДОБНО ""%007653%"""
        )
        query.УстановитьПараметр("Н", make_time(DATE_FROM))
        query.УстановитьПараметр("К", make_time(DATE_TO))
        try:
            result = query.Выполнить().Выбрать()
            count = 0
            while result.Следующий():
                count += 1
                line = f"  [{count}] Дт:{result.СчетДт} Кт:{result.СчетКт} Сумма={result.Сумма}"
                if "КоличествоДт" in fields_str:
                    line += f" КолДт={result.КоличествоДт} КолКт={result.КоличествоКт}"
                else:
                    line += f" Кол={result.Количество}"
                print(line)
            print(f"  [OK] {label}: {count} записей")
        except Exception as e:
            print(f"  [ОШИБКА] {label}: {str(e)[:150]}")


def test_document_000633(conn):
    """Проверяем документ ІБ00-000633 (Передача МБП)"""
    print("\n=== Документ ІБ00-000633 (Передача МБП) — все ресурсы ===")

    for fields_str, label in [
        ("Д.Сумма, Д.КоличествоДт, Д.КоличествоКт", "КоличествоДт/Кт"),
    ]:
        query = conn.NewObject("Запрос")
        query.Текст = (
            f"ВЫБРАТЬ\n"
            f"  ПРЕДСТАВЛЕНИЕ(Д.Регистратор) КАК Рег,\n"
            f"  Д.СчетДт,\n"
            f"  Д.СчетКт,\n"
            f"  {fields_str}\n"
            f"ИЗ\n"
            f"  РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(&Н, &К, , , ) КАК Д\n"
            f"ГДЕ\n"
            f"  ПРЕДСТАВЛЕНИЕ(Д.Регистратор) ПОДОБНО ""%000633%"""
        )
        query.УстановитьПараметр("Н", make_time(DATE_FROM))
        query.УстановитьПараметр("К", make_time(DATE_TO))
        try:
            result = query.Выполнить().Выбрать()
            count = 0
            while result.Следующий():
                count += 1
                line = f"  [{count}] Дт:{result.СчетДт} Кт:{result.СчетКт} Сумма={result.Сумма}"
                line += f" КолДт={result.КоличествоДт} КолКт={result.КоличествоКт}"
                print(line)
            if count == 0:
                print(f"  [!] НЕ НАЙДЕН — документ 000633 не попал в выборку")
            else:
                print(f"  [OK] {count} записей")
        except Exception as e:
            print(f"  [ОШИБКА] {str(e)[:150]}")


def main():
    print("=" * 80)
    print("  ДИАГНОСТИКА ПОЛЕЙ КОЛИЧЕСТВА в Хозрасчетный")
    print("=" * 80)

    conn = None
    try:
        conn = connect_buhbud()

        # 1. Какие поля есть в ДвиженияССубконто
        print("\n=== ДвиженияССубконто — проверка полей ===")
        test_field(conn, "Количество", "Кол")
        test_field(conn, "КоличествоДт", "КолДт")
        test_field(conn, "КоличествоКт", "КолКт")

        # 2. Какие поля есть в raw таблице
        test_raw_table_fields(conn)

        # 3. Конкретный документ 007653
        test_specific_document(conn)

        # 4. Конкретный документ 000633
        test_document_000633(conn)

    except Exception as e:
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn = None
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        print("\n[OK] Завершено")


if __name__ == "__main__":
    main()
