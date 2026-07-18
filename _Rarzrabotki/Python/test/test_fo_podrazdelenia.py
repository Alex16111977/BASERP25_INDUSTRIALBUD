# -*- coding: utf-8 -*-
"""
Проверяем значение ФО "ИспользоватьПодразделения" через COM.
Также проверяем, что делает ОтчетыУТПереопределяемый с нашим фильтром при загрузке варианта.
"""
import win32com.client
import pythoncom

CONNECTION_STRING = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    fo_names = [
        "ИспользоватьПодразделения",
        "ИспользоватьНесколькоОрганизаций",
        "ИспользоватьНесколькоВалют",
        "БазоваяВерсия",
    ]
    for fo in fo_names:
        try:
            val = conn.ПолучитьФункциональнуюОпцию(fo)
            print(f"  ФО {fo}: {val}")
        except Exception as e:
            print(f"  ФО {fo}: ERROR {e}")

    # Также проверим используется ли учёт по направлениям
    try:
        v = conn.Справочники.НаправленияДеятельности.ИспользуетсяУчетПоНаправлениям()
        print(f"  ИспользуетсяУчетПоНаправлениям: {v}")
    except Exception as e:
        print(f"  ИспользуетсяУчетПоНаправлениям: ERROR {e}")

    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
