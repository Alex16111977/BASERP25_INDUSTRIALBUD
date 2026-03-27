# -*- coding: utf-8 -*-
"""
Test 1: Basic query of balances on account 3721 (Settlements with accountable persons)
in BuhBud (BAS Accounting) via COM.

Checks:
  - Query executes without errors
  - Returns > 0 rows
  - Shows FL_Name + amounts
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime
import traceback

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'
TEST_DATE_START = datetime(2026, 2, 1)
TEST_DATE_END = datetime(2026, 3, 1)


def make_time(dt):
    return pywintypes.Time(dt)


def connect_buhbud():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected to BuhBud")
    return conn


QUERY_TEXT = '''
ВЫБРАТЬ
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).Наименование КАК ФЛ_Наименование,
  Ост.Счет.Код КАК КодСчета,
  СУММА(Ост.СуммаНачальныйОстатокДт) КАК НачОстатокДт,
  СУММА(Ост.СуммаНачальныйОстатокКт) КАК НачОстатокКт,
  СУММА(Ост.СуммаОборотДт) КАК ОборотДт,
  СУММА(Ост.СуммаОборотКт) КАК ОборотКт,
  СУММА(Ост.СуммаКонечныйОстатокДт) КАК КонОстатокДт,
  СУММА(Ост.СуммаКонечныйОстатокКт) КАК КонОстатокКт
ИЗ
  РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(
    &НачалоПериода,
    &КонецПериода) КАК Ост
ГДЕ
  Ост.Счет.Код = "3721"
  И Ост.Субконто1 ССЫЛКА Справочник.ФизическиеЛица
  И (Ост.СуммаКонечныйОстатокДт <> 0 ИЛИ Ост.СуммаКонечныйОстатокКт <> 0
    ИЛИ Ост.СуммаОборотДт <> 0 ИЛИ Ост.СуммаОборотКт <> 0
    ИЛИ Ост.СуммаНачальныйОстатокДт <> 0 ИЛИ Ост.СуммаНачальныйОстатокКт <> 0)
СГРУППИРОВАТЬ ПО
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).Наименование,
  Ост.Счет.Код
'''.strip()


def test_basic_3721(conn):
    """Test: Basic balances query for account 3721"""
    print("\n" + "=" * 80)
    print("TEST: Balances on account 3721 (accountable persons)")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = QUERY_TEXT
    query.УстановитьПараметр("НачалоПериода", make_time(TEST_DATE_START))
    query.УстановитьПараметр("КонецПериода", make_time(TEST_DATE_END))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        total_nach = 0
        total_kon = 0
        while result.Следующий():
            count += 1
            name = str(result.ФЛ_Наименование)[:50] if result.ФЛ_Наименование else ""
            acct = str(result.КодСчета).strip()
            nach_dt = result.НачОстатокДт or 0
            nach_kt = result.НачОстатокКт or 0
            oborot_dt = result.ОборотДт or 0
            oborot_kt = result.ОборотКт or 0
            kon_dt = result.КонОстатокДт or 0
            kon_kt = result.КонОстатокКт or 0
            saldo_nach = nach_dt - nach_kt
            saldo_kon = kon_dt - kon_kt
            total_nach += saldo_nach
            total_kon += saldo_kon
            print(f"  [{count:3d}] {name:50s} | Rax:{acct} | Nach:{saldo_nach:>12.2f} | Dt:{oborot_dt:>12.2f} | Kt:{oborot_kt:>12.2f} | Kon:{saldo_kon:>12.2f}")

        print(f"\n  Total accountable persons: {count}")
        print(f"  Total opening balance:  {total_nach:>15.2f}")
        print(f"  Total closing balance:  {total_kon:>15.2f}")

        if count > 0:
            print(f"  [OK] Test PASSED! Found {count} accountable persons")
        else:
            print(f"  [FAIL] No accountable persons found!")
        return count > 0
    except Exception as e:
        print(f"  [FAIL] {e}")
        traceback.print_exc()
        return False


def main():
    print("=" * 80)
    print("  Test: Basic balances query for account 3721 (BuhBud)")
    print(f"  Period: {TEST_DATE_START:%d.%m.%Y} - {TEST_DATE_END:%d.%m.%Y}")
    print("=" * 80)

    conn = None
    try:
        conn = connect_buhbud()
        passed = test_basic_3721(conn)
    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")
        traceback.print_exc()
        passed = False
    finally:
        conn = None
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    print("\n" + "=" * 80)
    print(f"  RESULT: {'PASSED' if passed else 'FAILED'}")
    print("=" * 80)


if __name__ == "__main__":
    main()
