# -*- coding: utf-8 -*-
"""
Диагностический тест: читаем реальные настройки фильтров варианта «Основной»
через СКД API — смотрим что СКД показывает как ВидНастройки/РежимРедактирования
для фильтра Подразделение.

Сравниваем с рабочим отчётом АнализДоходовРасходов.
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def inspect_filters(conn, report_name, variant_name=None):
    print(f"\n=== Отчёт: {report_name} (вариант: {variant_name or 'по-умолчанию'}) ===")
    report_mgr = getattr(conn.Reports, report_name)
    schema = report_mgr.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")

    composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    composer.Инициализировать(source)

    # Найти нужный вариант
    variant_settings = None
    if variant_name:
        for v in schema.ВариантыНастроек:
            if v.Имя == variant_name:
                variant_settings = v.Настройки
                break
    if variant_settings is None:
        variant_settings = schema.НастройкиПоУмолчанию

    composer.ЗагрузитьНастройки(variant_settings)
    settings = composer.Настройки

    # Итерируем фильтр верхнего уровня
    for item in settings.Отбор.Элементы:
        try:
            left = str(item.ЛевоеЗначение) if hasattr(item, "ЛевоеЗначение") else "?"
        except:
            left = "?"
        try:
            view_mode = item.ВидНастройки
        except:
            view_mode = "?"
        try:
            use = item.Использование
        except:
            use = "?"
        try:
            user_setting_id = item.ИдентификаторПользовательскойНастройки
        except:
            user_setting_id = "?"
        print(f"  Filter: {left:40} use={use}  viewMode={view_mode}  uid={user_setting_id}")


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    # Наш отчёт
    inspect_filters(conn, "А_ОтчетPL", "Основной")
    inspect_filters(conn, "А_ОтчетPL", "А_ОтчетPL_РасшифровкаДокументаОтчетPL")

    # Стандартный рабочий отчёт для сравнения
    try:
        inspect_filters(conn, "АнализДоходовРасходов")
    except Exception as e:
        print(f"АДР не доступен: {e}")

    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
