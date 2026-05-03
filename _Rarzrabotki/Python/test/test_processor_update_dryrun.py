# -*- coding: utf-8 -*-
"""
Smoke (dry-run): проверяет что ОбновитьСтатьюВДоговорах не падает на пустом массиве.
НЕ ИЗМЕНЯЕТ боевые данные — массив договоров пустой.
"""

import sys
from pathlib import Path
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EPF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\ЗаполнениеСтатьиДенегВДоговорах.epf"
STATIA_UUID = "3e2ba4c7-7a7e-11eb-a208-000c299fb278"


def fail(msg):
    print(f"FAIL: {msg}")
    sys.exit(1)


def main():
    if not Path(EPF_PATH).exists():
        fail(f"EPF не собран: {EPF_PATH}")

    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)

    vn_obr = erp.ВнешниеОбработки.Создать(EPF_PATH, True)

    uid = erp.NewObject("УникальныйИдентификатор", STATIA_UUID)
    statia_ref = erp.Справочники.СтатьиДвиженияДенежныхСредств.ПолучитьСсылку(uid)

    pusto = erp.NewObject("Массив")

    try:
        result = vn_obr.ОбновитьСтатьюВДоговорах(pusto, statia_ref)
    except Exception as e:
        if hasattr(e, "excepinfo") and e.excepinfo:
            fail(f"ОбновитьСтатьюВДоговорах упала: {e.excepinfo[2]}")
        else:
            fail(f"ОбновитьСтатьюВДоговорах упала: {e}")

    if int(result.Обновлено) != 0:
        fail(f"С пустым массивом Обновлено должно быть 0, получено {result.Обновлено}")
    if int(result.Пропущено) != 0:
        fail(f"С пустым массивом Пропущено должно быть 0, получено {result.Пропущено}")
    if int(result.Ошибки.Количество()) != 0:
        fail(f"С пустым массивом Ошибки должны быть пустыми")

    print("OK: dry-run ОбновитьСтатьюВДоговорах прошёл (0 обновлено, 0 пропущено, 0 ошибок)")
    print("\nSUCCESS: smoke-тест прошёл")


if __name__ == "__main__":
    main()
