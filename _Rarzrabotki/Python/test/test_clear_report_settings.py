# -*- coding: utf-8 -*-
"""
Очищает сохранённые пользовательские настройки отчёта А_ОтчетPL
для текущего пользователя (Администратор).

После этого при открытии отчёта будут загружены настройки из XML варианта
(viewMode=Normal вместо закэшированного Inaccessible).
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    # Попытаемся получить список сохранённых настроек через ХранилищеНастроекДанныхФорм
    # (в БСП пользовательские настройки отчёта обычно в ХранилищеПользовательскихНастроекОтчетов)
    storages = [
        "ХранилищеПользовательскихНастроекОтчетов",
        "ХранилищеВариантовОтчетов",
        "ХранилищеНастроекДанныхФорм",
        "ХранилищеСистемныхНастроек",
    ]

    for storage_name in storages:
        try:
            storage = getattr(conn, storage_name)
            print(f"\n=== {storage_name} ===")
            # Попробуем получить значения для пользователя Администратор
            # ключ объекта: "Отчет.А_ОтчетPL/<имя_варианта>"
            object_keys = [
                "Отчет.А_ОтчетPL",
                "Отчет.А_ОтчетPL/ОсновнаяСхемаКомпоновкиДанных",
                "Отчет.А_ОтчетPL/Основной",
                "Отчет.А_ОтчетPL/А_ОтчетPL_РасшифровкаДокументаОтчетPL",
            ]
            for okey in object_keys:
                try:
                    val = storage.Загрузить(okey, "", None, "Администратор")
                    if val is not None:
                        print(f"  Найдено для ключа {okey!r}: {type(val).__name__}")
                        storage.Удалить(okey, "", "Администратор")
                        print(f"    УДАЛЕНО")
                except Exception as e:
                    pass  # обычно "не найдено" = не ошибка
        except Exception as e:
            print(f"  {storage_name}: {e}")

    print("\nГотово. Попросите пользователя закрыть вкладку отчёта и открыть заново.")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
