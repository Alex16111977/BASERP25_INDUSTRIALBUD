# -*- coding: utf-8 -*-
"""Зондирует слоты автосохранения 'последнего вида' отчёта по варианту Глобино-2:
ХранилищеВариантовОтчетов и системное ХранилищеНастроекДанныхФорм, разные ключи/польз.
Ищем структуру с группировкой 'Регистратор'.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pythoncom, win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
erp = win32com.client.Dispatch("V83.COMConnector").Connect(CONN)

ВАР = "125715be-2fdd-4f6b-b24d-82eba493329d"
имяОтч = "Отчет.ВедомостьРасчетовСПартнерами"
объекты = [имяОтч, имяОтч + "/" + ВАР]
ключи = ["", ВАР, имяОтч + "/" + ВАР]
польз = ["", "Администратор"]

def найти_регистратор(знач, путь):
    """Рекурсивно ищем 'Регистратор' в структуре/полях значения настроек."""
    рез = []
    try:
        if hasattr(знач, "Структура"):
            def walk(s, lvl):
                for i in range(s.Количество()):
                    эл = s.Получить(i)
                    try:
                        пг = эл.ПоляГруппировки
                        поля = [str(пг.Элементы.Получить(j).Поле) for j in range(пг.Элементы.Количество())]
                        if lvl == 0:
                            рез.append("TOP:" + (",".join(поля) if поля else "<детальные>"))
                        for f in поля:
                            if "егистратор" in f:
                                рез.append("Регистратор@lvl%d" % lvl)
                        walk(эл.Структура, lvl + 1)
                    except Exception:
                        pass
            walk(знач.Структура, 0)
    except Exception:
        pass
    return рез

storages = []
for имя in ("ХранилищеВариантовОтчетов", "ХранилищеНастроекДанныхФорм", "СистемныеНастройки"):
    try:
        storages.append((имя, getattr(erp, имя)))
    except Exception:
        print(f"(нет хранилища {имя})")

for имяХр, st in storages:
    for ко in объекты:
        for кн in ключи:
            for п in польз:
                try:
                    зн = st.Загрузить(ко, кн, None, п)
                except Exception:
                    зн = None
                if зн is None:
                    continue
                тип = type(зн).__name__
                рег = найти_регистратор(зн, "")
                метка = "  <-- РЕГИСТРАТОР" if any("Регистратор@" in x for x in рег) else ""
                print(f"[{имяХр}] КО='{ко}' КН='{кн}' П='{п}' тип={тип} {рег}{метка}")

print("\nDONE")
