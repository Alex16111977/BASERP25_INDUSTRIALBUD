# -*- coding: utf-8 -*-
"""
Smoke: после загрузки нового ManagerModule у Документ.А_ОтчетPL метод
ДобавитьКомандыОтчетов должен РЕЗОЛВИТЬСЯ (именно он падал в ВариантыОтчетов.Модуль(7123)
"Метод объекта не обнаружен"). Свежее COM-соединение видит правку после /db-update.
"""
import sys, io
import win32com.client, pythoncom

if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

pythoncom.CoInitialize()
ERP = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def errmsg(e):
    return e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)


mgr = ERP.Документы.А_ОтчетPL

# 1) метод резолвится? (вызов с пустыми аргументами — упадёт на аргументах, НО не на "метод не обнаружен")
resolved = False
detail = ""
try:
    mgr.ДобавитьКомандыОтчетов(None, None)
    resolved = True
    detail = "вызвался без ошибки"
except Exception as e:
    m = errmsg(e)
    if ("не обнаружен" in m) or ("not found" in m.lower()):
        resolved = False
        detail = m
    else:
        resolved = True
        detail = "метод есть (упал на аргументах — ожидаемо): " + m

print(f"[1] Документы.А_ОтчетPL.ДобавитьКомандыОтчетов резолвится: {resolved}")
print(f"    {detail}")

# 2) полноценный вызов с реальной таблицей команд БСП (если получится через COM)
full = None
try:
    Команды = ERP.ВариантыОтчеты.КомандыОтчетов()
    n0 = Команды.Количество()
    mgr.ДобавитьКомандыОтчетов(Команды, None)
    full = Команды.Количество()
    print(f"[2] Полный вызов OK: команд было {n0}, стало {full} (структура подчинённости добавлена)")
except Exception as e:
    print(f"[2] Полный вызов через COM не выполнен (норм — главное п.1): {errmsg(e)}")

assert resolved, "FAIL: метод ДобавитьКомандыОтчетов всё ещё НЕ резолвится — фикс не применился"
print("\n########## SMOKE OK — 'Метод объекта не обнаружен' устранён ##########")
