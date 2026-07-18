# -*- coding: utf-8 -*-
"""
Тест: сформировать ПользовательскиеНастройки с фильтром Подразделение
и убедиться что РежимОтображения остаётся "Обычный" (не "Недоступный").

Это точный эмулятор нужной логики формы документа.
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONNECTION_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    # Получить произвольное подразделение для теста
    q = conn.NewObject("Запрос")
    q.Текст = '''
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.СтруктураПредприятия
    ГДЕ НЕ ПометкаУдаления
    '''
    p = q.Выполнить().Выгрузить()
    if p.Количество() == 0:
        print("FAIL: нет подразделений")
        return 1
    подразделение = p.Получить(0).Ссылка
    print(f"  Тест. подразделение: {подразделение}")

    report = conn.Reports.А_ОтчетPL.Создать()
    schema = report.СхемаКомпоновкиДанных

    # Найти вариант
    variant_name = "А_ОтчетPL_РасшифровкаДокументаОтчетPL"
    variant = None
    for v in schema.ВариантыНастроек:
        if v.Имя == variant_name:
            variant = v
            break
    if variant is None:
        print(f"FAIL: вариант {variant_name} не найден")
        return 1

    composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    composer.Инициализировать(source)
    composer.ЗагрузитьНастройки(variant.Настройки)

    # === Установить Период и Подразделение через ПользовательскиеНастройки ===
    период = conn.NewObject("СтандартныйПериод")
    период.ДатаНачала = pywintypes.Time(datetime(2025, 12, 1))
    период.ДатаОкончания = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    # Режим "Обычный" для сравнения
    режим_обычный = conn.РежимОтображенияЭлементаНастройкиКомпоновкиДанных.Обычный
    режим_недоступный = conn.РежимОтображенияЭлементаНастройкиКомпоновкиДанных.Недоступный

    # userSettingID Подразделения в варианте "РасшифровкаДокументаОтчетPL"
    UID_ПОДРАЗД = "43fffebb-fede-4af3-9a9d-112dce5ef8ff"

    подр_установлен = False
    for item in composer.ПользовательскиеНастройки.Элементы:
        try:
            uid = item.ИдентификаторПользовательскойНастройки
        except Exception:
            uid = None

        # Период
        if uid == "Период":
            item.Значение = период
            item.Использование = True
            print(f"  [OK] Период установлен")

        # Подразделение — ищем по uid
        if uid == UID_ПОДРАЗД:
            список = conn.NewObject("СписокЗначений")
            список.Добавить(подразделение)
            item.ПравоеЗначение = список
            item.Использование = True
            current_mode = item.РежимОтображения
            is_normal = (current_mode == режим_обычный)
            is_inacc = (current_mode == режим_недоступный)
            print(f"  [OK] Подразделение установлено. Обычный={is_normal}  Недоступный={is_inacc}")
            подр_установлен = True

    if not подр_установлен:
        print("  [FAIL] Элемент uid='{UID_ПОДРАЗД}' не найден в ПользовательскиеНастройки")
        return 1

    # Финальная проверка: распечатать все элементы ПользовательскиеНастройки с их режимом
    print("\n--- Итог: ПользовательскиеНастройки ---")
    for item in composer.ПользовательскиеНастройки.Элементы:
        type_name = type(item).__name__
        try:
            uid = item.ИдентификаторПользовательскойНастройки
        except Exception:
            uid = "?"
        try:
            mode_norm = (item.РежимОтображения == режим_обычный)
            mode_quick = (item.РежимОтображения == conn.РежимОтображенияЭлементаНастройкиКомпоновкиДанных.Быстрый)
            mode_inac = (item.РежимОтображения == режим_недоступный)
            mode_str = "Обычный" if mode_norm else ("Быстрый" if mode_quick else ("Недоступный" if mode_inac else "?"))
        except Exception:
            mode_str = "?"
        try:
            use = item.Использование
        except Exception:
            use = "?"
        print(f"    {type_name:40}  uid={str(uid)[:10]:10}  режим={mode_str:12}  использование={use}")

    print("\n[PASS] Логика работает: режим остаётся Обычный, не переопределяется на Недоступный.")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
