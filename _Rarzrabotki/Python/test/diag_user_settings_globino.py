# -*- coding: utf-8 -*-
"""Ищет per-user сохранённые настройки варианта 'Глобино-2' в ХранилищеВариантовОтчетов
и проверяет, есть ли там группировка/поле 'Регистратор' (источник Регистратора наверху).
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pythoncom, win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
erp = win32com.client.Dispatch("V83.COMConnector").Connect(CONN)

storage = erp.ХранилищеВариантовОтчетов
ВАР_КЛЮЧ = "125715be-2fdd-4f6b-b24d-82eba493329d"
кандидаты = [
    "Отчет.ВедомостьРасчетовСПартнерами",
    "Отчет.ВедомостьРасчетовСПартнерами/" + ВАР_КЛЮЧ,
    ВАР_КЛЮЧ,
]

def dump_struct(структура, уровень, найдено):
    for i in range(структура.Количество()):
        эл = структура.Получить(i)
        try:
            пг = эл.ПоляГруппировки
            поля = []
            for j in range(пг.Элементы.Количество()):
                s = str(пг.Элементы.Получить(j).Поле)
                поля.append(s)
                if "егистратор" in s:
                    найдено.append(s)
            print("  " * уровень + "• " + (", ".join(поля) if поля else "<детальные>"))
            dump_struct(эл.Структура, уровень + 1, найдено)
        except Exception:
            print("  " * уровень + "• [" + type(эл).__name__ + "]")

for ко in кандидаты:
    print(f"\n=== ПолучитьСписок('{ко}') ===")
    try:
        сп = storage.ПолучитьСписок(ко)
        print(f"  элементов: {сп.Количество()}")
        for i in range(сп.Количество()):
            эл = сп.Получить(i)
            print(f"    КлючНастроек='{эл.Значение}'  Представление='{эл.Представление}'")
    except Exception as e:
        msg = e.excepinfo[2] if (hasattr(e,'excepinfo') and e.excepinfo) else str(e)
        print(f"  ПолучитьСписок fail: {msg}")
        continue
    # пробуем загрузить каждый ключ для Администратора
    try:
        for i in range(сп.Количество()):
            кн = сп.Получить(i).Значение
            for польз in ("Администратор", ""):
                try:
                    знач = storage.Загрузить(ко, кн, None, польз)
                except Exception:
                    знач = None
                if знач is None:
                    continue
                тип = type(знач).__name__
                print(f"\n  >>> Загружено КлючОбъекта='{ко}' КлючНастроек='{кн}' Пользователь='{польз}' тип={тип}")
                найдено = []
                # НастройкиКомпоновкиДанных -> .Структура ; Пользовательские -> .Элементы
                try:
                    if hasattr(знач, "Структура"):
                        print("    --- Структура ---")
                        dump_struct(знач.Структура, 2, найдено)
                except Exception as e:
                    print("    (структура не прочитана):", e)
                try:
                    if hasattr(знач, "Элементы"):
                        print(f"    --- Элементы пользовательских настроек: {знач.Элементы.Количество()} ---")
                        for k in range(знач.Элементы.Количество()):
                            ук = знач.Элементы.Получить(k)
                            s = type(ук).__name__
                            print(f"      [{s}]")
                except Exception:
                    pass
                if найдено:
                    print(f"    !!! НАЙДЕН Регистратор в настройках: {найдено}")
    except Exception as e:
        print("  (load loop fail):", e)

print("\nDONE")
