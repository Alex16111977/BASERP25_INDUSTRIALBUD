# -*- coding: utf-8 -*-
"""
Через СКД API читаем фактический РежимОтображения фильтра Подразделение
варианта "Основной" отчёта А_ОтчетPL.

Также смотрим, показывается ли фильтр в Пользовательских настройках.
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    schema = conn.Reports.А_ОтчетPL.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
    composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    composer.Инициализировать(source)
    composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)

    print("\n--- Отбор на уровне Настроек варианта ---")
    for item in composer.Настройки.Отбор.Элементы:
        # Попытка через getattr множество имён свойств
        props = {}
        for name in ["ЛевоеЗначение", "ВидСравнения", "ПравоеЗначение",
                     "РежимОтображения", "РежимОтображенияПользовательскойНастройки",
                     "Использование", "ИдентификаторПользовательскойНастройки",
                     "ВидНастройкиПользовательскойНастройки"]:
            try:
                v = getattr(item, name)
                props[name] = v
            except AttributeError:
                pass
            except Exception as e:
                props[name] = f"err:{e}"
        print(f"  {props}")

    print("\n--- ПользовательскиеНастройки Компоновщика (с XMLСтрока для enum) ---")
    try:
        user = composer.ПользовательскиеНастройки
        print(f"  Элементов: {user.Элементы.Количество()}")
        for item in user.Элементы:
            try:
                uid = item.ИдентификаторПользовательскойНастройки
            except Exception:
                uid = "?"
            try:
                view_mode_enum = item.РежимОтображения
                view_mode_str = conn.XMLСтрока(view_mode_enum)
            except Exception as e:
                view_mode_str = f"err:{e}"
            print(f"    uid={uid:40}  РежимОтображения={view_mode_str}")
    except Exception as e:
        print(f"  FAIL: {e}")

    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
