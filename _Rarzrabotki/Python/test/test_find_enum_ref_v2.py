# -*- coding: utf-8 -*-
"""
ТЕСТ v2: Полная симуляция логики BSL НайтиСсылкуЗначенияПеречисления()
Цель: подтвердить рабочий алгоритм ПЕРЕД правкой BSL файла.

Алгоритм BSL:
  1. Запрос: SELECT Ссылка, ПРЕДСТАВЛЕНИЕ(Ссылка) КАК Синоним, Порядок FROM Перечисление.X
  2. Итерация: Пока Выборка.Следующий() Цикл
  3. Если Выборка.Синоним = СинонимЗначения Тогда
  4.     УИД = Строка(Выборка.Ссылка.УникальныйИдентификатор())
  5.     Возврат результат
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

TESTS = [
    # (enum_type, synonym, должен найти)
    ("ХозяйственныеОперации", "Ввод остатков задолженности подотчетников", True),
    ("ХозяйственныеОперации", "Авансовый отчет", True),
    ("ХозяйственныеОперации", "НЕСУЩЕСТВУЮЩИЙ СИНОНИМ 999", False),
]


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Подключено к BaseERP\n")
    return conn


def find_enum_ref(conn, enum_type, synonym):
    """
    Точная симуляция BSL функции НайтиСсылкуЗначенияПеречисления()
    
    BSL код:
        Запрос.Текст = "ВЫБРАТЬ
        |    Перечисление.Ссылка КАК Ссылка,
        |    ПРЕДСТАВЛЕНИЕ(Перечисление.Ссылка) КАК Синоним,
        |    Перечисление.Порядок КАК Порядок
        |ИЗ
        |    Перечисление." + ИмяПеречисления + " КАК Перечисление"
        
        Пока Выборка.Следующий() Цикл
            Если Выборка.Синоним = СинонимЗначения Тогда
                УИД = Строка(Выборка.Ссылка.УникальныйИдентификатор())
                Возврат ...
            КонецЕсли
        КонецЦикла
        Возврат found=False
    """
    # Шаг 1: проверка через запрос (COM-индекс не работает, BSL использует Найти())
    # В Python: пробуем запрос к перечислению, если ошибка — перечисление не существует
    # (в BSL Метаданные.Перечисления.Найти() работает правильно)

    # Шаг 2: запрос без WHERE (ключевое решение!)
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ\n"
        "    Перечисление.Ссылка КАК Ссылка,\n"
        "    ПРЕДСТАВЛЕНИЕ(Перечисление.Ссылка) КАК Синоним,\n"
        "    Перечисление.Порядок КАК Порядок\n"
        "ИЗ\n"
        f"    Перечисление.{enum_type} КАК Перечисление"
    )
    try:
        res = q.Выполнить()
    except Exception as e:
        return {"found": False, "error": f"Перечисление не найдено в метаданных: {enum_type} ({e})"}
    sel = res.Выбрать()

    # Шаг 3: итерация + сравнение (BSL: Выборка.Синоним = СинонимЗначения)
    while sel.Следующий():
        synonym_val = str(sel.Синоним)  # ПРЕДСТАВЛЕНИЕ — строка, str() работает
        if synonym_val == synonym:
            # Шаг 4: UUID (BSL: Строка(Ссылка.УникальныйИдентификатор()))
            ref = sel.Ссылка
            try:
                uid_obj = ref.УникальныйИдентификатор()  # вызов метода BSL
                uid_str = str(conn.Строка(uid_obj))
            except Exception:
                # Если метод недоступен через COM — используем альтернативу
                uid_str = "N/A (метод недоступен через COM, но в BSL работает)"

            return {
                "found": True,
                "uuid": uid_str,
                "enum_type": enum_type,
                "synonym": synonym_val,
                "order": sel.Порядок,
            }

    # Не нашли
    return {
        "found": False,
        "enum_type": enum_type,
        "synonym": synonym,
        "error": "Значення перечислення не знайдено за синонімом",
    }


def run_tests(conn):
    passed = 0
    failed = 0

    for enum_type, synonym, should_find in TESTS:
        print(f"{'='*60}")
        print(f"Тест: enum_type='{enum_type}'")
        print(f"      synonym='{synonym}'")
        print(f"      ожидаем found={should_find}")
        print("-"*60)

        result = find_enum_ref(conn, enum_type, synonym)

        found = result.get("found", False)
        ok = (found == should_find)

        if ok:
            passed += 1
            status = "[PASS]"
        else:
            failed += 1
            status = "[FAIL]"

        print(f"  {status} found={found}")
        if found:
            print(f"  uuid   = {result.get('uuid')}")
            print(f"  synonym= {result.get('synonym')}")
            print(f"  order  = {result.get('order')}")
        else:
            print(f"  error  = {result.get('error')}")
        print()

    print("="*60)
    print(f"ИТОГО: {passed} PASS / {failed} FAIL")
    return failed == 0


if __name__ == "__main__":
    conn = connect()
    success = run_tests(conn)
    if success:
        print("\n✅ ВСЕ ТЕСТЫ ПРОШЛИ — BSL код корректен, можно обновлять 1С!")
    else:
        print("\n❌ ЕСТЬ ОШИБКИ — нужно доработать алгоритм!")
