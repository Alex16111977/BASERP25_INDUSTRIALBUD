# -*- coding: utf-8 -*-
"""
Проверка исправленной логики: получаем uid фильтра из Настройки.Отбор,
потом устанавливаем значение в ПользовательскиеНастройки по этому uid.
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    q = conn.NewObject("Запрос")
    q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.СтруктураПредприятия ГДЕ НЕ ПометкаУдаления'
    p = q.Выполнить().Выгрузить()
    подразделение = p.Получить(0).Ссылка
    print(f"  Подразделение: получено")

    report = conn.Reports.А_ОтчетPL.Создать()
    schema = report.СхемаКомпоновкиДанных

    variant_name = "А_ОтчетPL_РасшифровкаДокументаОтчетPL"
    variant = None
    for v in schema.ВариантыНастроек:
        if v.Имя == variant_name:
            variant = v
            break

    composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    composer.Инициализировать(source)
    composer.ЗагрузитьНастройки(variant.Настройки)

    # Шаг 1: находим uid из Настройки.Отбор (здесь ЛевоеЗначение доступно)
    uid_подразд = ""
    uid_напр = ""
    print(f"  Настройки.Отбор.Элементы.Количество() = {composer.Настройки.Отбор.Элементы.Количество()}")
    for idx in range(composer.Настройки.Отбор.Элементы.Количество()):
        el = composer.Настройки.Отбор.Элементы.Получить(idx)
        # В 1C Строка(ЛевоеЗначение) вернёт имя поля
        try:
            name = conn.Строка(el.ЛевоеЗначение)
        except Exception:
            try:
                name = conn.XMLСтрока(el.ЛевоеЗначение)
            except Exception:
                name = ""
        uid = el.ИдентификаторПользовательскойНастройки
        print(f"  [{idx}] ЛевоеЗначение='{name}'  uid={uid}")
        if name == "Подразделение":
            uid_подразд = uid
        elif name == "НаправлениеДеятельности":
            uid_напр = uid

    # Шаг 2: устанавливаем в Компоновщик.ПользовательскиеНастройки по этому uid
    период = conn.NewObject("СтандартныйПериод")
    период.ДатаНачала = pywintypes.Time(datetime(2025, 12, 1))
    период.ДатаОкончания = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    установлено = {"Период": False, "Подразделение": False}
    for item in composer.ПользовательскиеНастройки.Элементы:
        uid = item.ИдентификаторПользовательскойНастройки
        if uid == "Период":
            item.Значение = период
            item.Использование = True
            установлено["Период"] = True
        elif uid == uid_подразд:
            list_v = conn.NewObject("СписокЗначений")
            list_v.Добавить(подразделение)
            item.ПравоеЗначение = list_v
            item.Использование = True
            установлено["Подразделение"] = True

    print(f"  Установлено: {установлено}")
    if all(установлено.values()):
        print("[PASS] Всё ОК, логика работает.")
        return 0
    else:
        print("[FAIL] Не все значения установлены")
        return 1


if __name__ == "__main__":
    import sys
    sys.exit(main())
