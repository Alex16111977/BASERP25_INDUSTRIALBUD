# -*- coding: utf-8 -*-
"""
Тест: Как получить UUID значения перечисления в 1С?
Проверяем разные способы через COM
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ENUM_NAME = "ХозяйственныеОперации"
SYNONYM = "Авансовый отчет"  # Порядок=0, проще искать


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Подключено к BaseERP\n")
    return conn


def test_uuid_methods(conn):
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1\n"
        "    Перечисление.Ссылка КАК Ссылка,\n"
        "    ПРЕДСТАВЛЕНИЕ(Перечисление.Ссылка) КАК Синоним,\n"
        "    Перечисление.Порядок КАК Порядок\n"
        "ИЗ\n"
        "    Перечисление.ХозяйственныеОперации КАК Перечисление\n"
        "ГДЕ\n"
        "    Перечисление.Порядок = 0"
    )
    res = q.Выполнить()
    sel = res.Выбрать()
    sel.Следующий()
    
    ref = sel.Ссылка
    print(f"Синоним: {str(sel.Синоним)}  Порядок: {sel.Порядок}")
    print(f"type(ref): {type(ref)}")
    
    # Способ 1: XMLСтрока
    try:
        v = str(conn.XMLСтрока(ref))
        print(f"[1] XMLСтрока(ref) = '{v}'")
    except Exception as e:
        print(f"[1] XMLСтрока ОШИБКА: {e}")

    # Способ 2: Строка(ref)
    try:
        v = str(conn.Строка(ref))
        print(f"[2] Строка(ref) = '{v}'")
    except Exception as e:
        print(f"[2] Строка ОШИБКА: {e}")

    # Способ 3: Через XML-сериализацию
    try:
        ЗаписьXML = conn.NewObject("ЗаписьXML")
        ЗаписьXML.УстановитьСтроку()
        conn.СериализаторXDTO.ЗаписатьXML(ЗаписьXML, ref, 2)  # 2 = Явное
        xml_str = str(ЗаписьXML.Закрыть())
        print(f"[3] СериализаторXDTO.ЗаписатьXML = '{xml_str}'")
    except Exception as e:
        print(f"[3] СериализаторXDTO ОШИБКА: {e}")

    # Способ 4: Через ОбщегоНазначения.ЗначениеВСтрокуXML (БСП)
    try:
        v = str(conn.ОбщегоНазначения.ЗначениеВСтрокуXML(ref))
        print(f"[4] ОбщегоНазначения.ЗначениеВСтрокуXML = '{v}'")
    except Exception as e:
        print(f"[4] ОбщегоНазначения ОШИБКА: {e}")

    # Способ 5: Через метаданные
    try:
        meta_val = conn.Метаданные.Перечисления[ENUM_NAME]
        print(f"[5] meta_val.ЗначенияПеречисления[0].Имя = '{str(meta_val.ЗначенияПеречисления[0].Имя)}'")
    except Exception as e:
        print(f"[5] Метаданные ОШИБКА: {e}")

    # Способ 6: ref как COM object - смотрим доступные атрибуты
    try:
        attrs = [a for a in dir(ref) if not a.startswith('_')]
        print(f"[6] dir(ref) = {attrs[:20]}")
    except Exception as e:
        print(f"[6] dir(ref) ОШИБКА: {e}")

    # Способ 7: Запрос с BSL функцией через параметр
    try:
        q2 = conn.NewObject("Запрос")
        q2.Текст = (
            "ВЫБРАТЬ\n"
            "    &РефПеречисления КАК РефПеречисления\n"
        )
        q2.УстановитьПараметр("РефПеречисления", ref)
        r2 = q2.Выполнить().Выбрать()
        r2.Следующий()
        ref2 = r2.РефПеречисления
        # Теперь пробуем XMLСтрока на результат
        xml2 = str(conn.XMLСтрока(ref2))
        print(f"[7] XMLСтрока через запрос = '{xml2}'")
    except Exception as e:
        print(f"[7] Запрос с параметром ОШИБКА: {e}")

    # Способ 8: ОбменДаннымиXML (если есть в БСП)
    try:
        ser = conn.NewObject("СериализаторXDTO", conn.ФабрикаXDTO)
        ЗаписьXML = conn.NewObject("ЗаписьXML")
        ЗаписьXML.УстановитьСтроку()
        ser.ЗаписатьXML(ЗаписьXML, ref)
        xml_str = str(ЗаписьXML.Закрыть())
        print(f"[8] СериализаторXDTO (NewObject) = '{xml_str}'")
    except Exception as e:
        print(f"[8] СериализаторXDTO (NewObject) ОШИБКА: {e}")


if __name__ == "__main__":
    conn = connect()
    test_uuid_methods(conn)
    print("\n[ГОТОВО]")
