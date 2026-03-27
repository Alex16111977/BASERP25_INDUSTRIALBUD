# -*- coding: utf-8 -*-
"""
Тест: find_enum_ref — поиск UUID значения перечисления по синониму
База: BaseERP
Цель: найти рабочий вариант запроса для НайтиСсылкуЗначенияПеречисления()
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ENUM_NAME = "ХозяйственныеОперации"
SYNONYM = "Ввод остатков задолженности подотчетников"


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Подключено к BaseERP")
    return conn


def test1_all_values(conn):
    """Тест 1: Все значения — Ссылка как строка (Представление)"""
    print("\n" + "="*60)
    print("ТЕСТ 1: Все значения ХозяйственныеОперации (первые 5)")
    print("="*60)
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 5\n"
        "    Перечисление.Ссылка КАК Ссылка,\n"
        "    Перечисление.Порядок КАК Порядок\n"
        "ИЗ\n"
        "    Перечисление.ХозяйственныеОперации КАК Перечисление\n"
        "УПОРЯДОЧИТЬ ПО Порядок"
    )
    res = q.Выполнить()
    sel = res.Выбрать()
    while sel.Следующий():
        ref = sel.Ссылка
        presentation = str(ref)  # str(COMObject) = синоним/представление
        print(f"  Порядок={sel.Порядок}  Представление={presentation}")


def test2_via_metadata(conn):
    """Тест 2: Перебор через метаданные объектной модели"""
    print("\n" + "="*60)
    print(f"ТЕСТ 2: Поиск через метаданные (Имя + Синоним)")
    print("="*60)
    try:
        meta_enums = conn.Метаданные.Перечисления
        meta_enum = meta_enums[ENUM_NAME]
        vals = meta_enum.ЗначенияПеречисления
        count = vals.Количество()
        print(f"  Всего значений: {count}")
        enum_manager = conn.Перечисления[ENUM_NAME]
        found = None
        for i in range(count):
            mv = vals[i]
            prog_name = str(mv.Имя)
            synonym_val = str(mv.Синоним)
            enum_val = enum_manager[prog_name]
            pres = str(enum_val)
            print(f"  [{i}] Имя={prog_name}  Синоним={synonym_val}  Представление={pres}")
            if synonym_val == SYNONYM or pres == SYNONYM:
                found = {"prog_name": prog_name, "synonym": synonym_val, "presentation": pres}
        if found:
            print(f"\n  [НАЙДЕНО] {found}")
        else:
            print(f"\n  [НЕ НАЙДЕНО] Ищем: {SYNONYM}")
    except Exception as e:
        print(f"  [ОШИБКА] {e}")


def test3_filter_in_python(conn):
    """Тест 3: ПРЕДСТАВЛЕНИЕ в SELECT + фильтр в Python через conn.Строка()"""
    print("\n" + "="*60)
    print("ТЕСТ 3: ПРЕДСТАВЛЕНИЕ в SELECT + фильтр в Python")
    print("="*60)
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ\n"
        "    Перечисление.Ссылка КАК Ссылка,\n"
        "    ПРЕДСТАВЛЕНИЕ(Перечисление.Ссылка) КАК Синоним,\n"
        "    Перечисление.Порядок КАК Порядок\n"
        "ИЗ\n"
        "    Перечисление.ХозяйственныеОперации КАК Перечисление"
    )
    res = q.Выполнить()
    sel = res.Выбрать()
    found = None
    total = 0
    while sel.Следующий():
        total += 1
        synonym_val = str(sel.Синоним)  # ПРЕДСТАВЛЕНИЕ - строковый тип, str() работает!
        ref = sel.Ссылка
        if total <= 3:  # Показываем первые 3 для диагностики
            print(f"  [{total}] Синоним='{synonym_val}'  Порядок={sel.Порядок}")
        if synonym_val == SYNONYM:
            # UUID через XMLСтрока
            try:
                xml_str = str(conn.XMLСтрока(ref))
                print(f"  [НАЙДЕНО] XMLСтрока={xml_str}  Порядок={sel.Порядок}")
                found = xml_str
            except Exception as e:
                print(f"  [НАЙДЕНО, XMLСтрока ошибка: {e}]")
    print(f"  Всего проверено: {total}")
    if not found:
        print(f"  [НЕ НАЙДЕНО]  Искали: '{SYNONYM}'")


def test4_подобно(conn):
    """Тест 4: WHERE ПРЕДСТАВЛЕНИЕ ПОДОБНО &Синоним"""
    print("\n" + "="*60)
    print("ТЕСТ 4: WHERE ПРЕДСТАВЛЕНИЕ ПОДОБНО &Синоним")
    print("="*60)
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ\n"
        "    Перечисление.Ссылка КАК Ссылка,\n"
        "    Перечисление.Порядок КАК Порядок\n"
        "ИЗ\n"
        "    Перечисление.ХозяйственныеОперации КАК Перечисление\n"
        "ГДЕ\n"
        "    ПРЕДСТАВЛЕНИЕ(Перечисление.Ссылка) ПОДОБНО &Синоним"
    )
    q.УстановитьПараметр("Синоним", SYNONYM)
    try:
        res = q.Выполнить()
        sel = res.Выбрать()
        found = False
        while sel.Следующий():
            found = True
            print(f"  [OK] Представление={str(sel.Ссылка)}  Порядок={sel.Порядок}")
        if not found:
            print("  [НЕ НАЙДЕНО]")
    except Exception as e:
        print(f"  [ОШИБКА] {e}")


def test5_uuid_via_string(conn):
    """Тест 5: Получение UUID через Строка(Ссылка.УникальныйИдентификатор())"""
    print("\n" + "="*60)
    print("ТЕСТ 5: UUID через Строка(Ссылка.УникальныйИдентификатор())")
    print("="*60)
    q = conn.NewObject("Запрос")
    # Главное открытие: запрос через BSL может использовать Строка(Ссылка.УникальныйИдентификатор())
    # в SELECT, но через COM нам нужно прочитать Ссылка.УникальныйИдентификатор() напрямую
    # Или использовать XMLСтрока на OID
    # Пробуем: прочитать OID через COM-атрибут Ссылка
    q.Текст = (
        "ВЫБРАТЬ\n"
        "    Перечисление.Ссылка КАК Ссылка,\n"
        "    ПРЕДСТАВЛЕНИЕ(Перечисление.Ссылка) КАК Синоним,\n"
        "    Перечисление.Порядок КАК Порядок\n"
        "ИЗ\n"
        "    Перечисление.ХозяйственныеОперации КАК Перечисление"
    )
    res = q.Выполнить()
    sel = res.Выбрать()
    while sel.Следующий():
        synonym_val = str(sel.Синоним)
        if synonym_val == SYNONYM:
            ref = sel.Ссылка
            # Пробуем разные способы UUID через COM
            print(f"  Найдено: '{synonym_val}'  Порядок={sel.Порядок}")
            # Способ 1: conn.Строка(ref.УникальныйИдентификатор())
            try:
                uid_obj = ref.УникальныйИдентификатор()
                uid_str = str(conn.Строка(uid_obj))
                print(f"  [1] Строка(УникальныйИдентификатор())={uid_str}")
            except Exception as e:
                print(f"  [1] Ошибка: {e}")
            # Способ 2: XMLСтрока на ссылку
            try:
                xml_str = str(conn.XMLСтрока(ref))
                print(f"  [2] XMLСтрока(Ссылка)={xml_str}")
            except Exception as e:
                print(f"  [2] Ошибка: {e}")
            # Способ 3: GUIDString через запрос с функцией в 1С BSL
            try:
                q2 = conn.NewObject("Запрос")
                q2.Текст = (
                    "ВЫБРАТЬ\n"
                    "    &Ссылка.УникальныйИдентификатор() КАК УИД"
                )
                q2.УстановитьПараметр("Ссылка", ref)
                r2 = q2.Выполнить().Wыбрать()
                if r2.Следующий():
                    print(f"  [3] Через параметр запроса: {str(r2.УИД)}")
            except Exception as e:
                print(f"  [3] Ошибка: {e}")
            break


if __name__ == "__main__":
    conn = connect()
    test1_all_values(conn)
    test2_via_metadata(conn)
    test3_filter_in_python(conn)
    test4_подобно(conn)
    test5_uuid_via_string(conn)
    print("\n[ГОТОВО]")
