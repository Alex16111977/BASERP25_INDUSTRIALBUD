# -*- coding: utf-8 -*-
"""
Test 2: Query with ВЫРАЗИТЬ(Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод
for account 3721 in BuhBud (BAS Accounting) via COM.

Checks:
  - А_ИдКод is extracted successfully
  - Non-empty А_ИдКод count > 0
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime
import traceback

CONNECTION_STRING = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
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
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод КАК ФЛ_А_ИдКод,
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).Наименование КАК ФЛ_Наименование,
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).Код КАК ФЛ_Код,
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
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод,
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).Наименование,
  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).Код
'''.strip()


def test_subconto_a_idkod(conn):
    """Test: Extract A_IdKod from Subconto1 for account 3721"""
    print("\n" + "=" * 80)
    print("TEST: Extract A_IdKod from Subconto1 for account 3721")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = QUERY_TEXT
    query.УстановитьПараметр("НачалоПериода", make_time(TEST_DATE_START))
    query.УстановитьПараметр("КонецПериода", make_time(TEST_DATE_END))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        count_with_idkod = 0
        count_without_idkod = 0
        total_kon = 0
        while result.Следующий():
            count += 1
            a_idkod = str(result.ФЛ_А_ИдКод).strip() if result.ФЛ_А_ИдКод else ""
            name = str(result.ФЛ_Наименование)[:45] if result.ФЛ_Наименование else ""
            code = str(result.ФЛ_Код).strip() if result.ФЛ_Код else ""
            kon_dt = result.КонОстатокДт or 0
            kon_kt = result.КонОстатокКт or 0
            saldo = kon_dt - kon_kt
            total_kon += saldo

            has_id = len(a_idkod) > 0
            if has_id:
                count_with_idkod += 1
            else:
                count_without_idkod += 1

            id_status = "OK" if has_id else "NO ID"
            print(f"  [{count:3d}] [{id_status:5s}] Kod:{code:10s} | {name:45s} | Saldo:{saldo:>12.2f} | A_IdKod:{a_idkod[:36]}")

        print(f"\n  Total persons:     {count}")
        print(f"  With A_IdKod:      {count_with_idkod}")
        print(f"  Without A_IdKod:   {count_without_idkod}")
        print(f"  Total saldo:       {total_kon:>15.2f}")

        if count > 0 and count_with_idkod > 0:
            print(f"  [OK] Test PASSED! {count_with_idkod} persons have A_IdKod")
        elif count > 0:
            print(f"  [WARN] Persons found but NO A_IdKod extracted!")
        else:
            print(f"  [FAIL] No accountable persons found!")
        return count > 0 and count_with_idkod > 0
    except Exception as e:
        print(f"  [FAIL] {e}")
        traceback.print_exc()
        return False


def main():
    print("=" * 80)
    print("  Test: Extract A_IdKod from subconto for account 3721 (BuhBud)")
    print(f"  Period: {TEST_DATE_START:%d.%m.%Y} - {TEST_DATE_END:%d.%m.%Y}")
    print("=" * 80)

    conn = None
    try:
        conn = connect_buhbud()
        passed = test_subconto_a_idkod(conn)
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
