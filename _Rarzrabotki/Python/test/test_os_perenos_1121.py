# -*- coding: utf-8 -*-
"""
Тест: перевірка що запит з ИЛИ для рахунку 1121 (МНМА) працює в BuhBud.

Перевіряємо:
  - Запит вартості: Остатки(Дата, Счет В ИЕРАРХИИ(ОсновныеСредства) ИЛИ Счет = МалоценныеНеоборотныеМатериальныеАктивыИндивидуально, ,)
  - Запит зносу: Остатки(Дата, Счет В ИЕРАРХИИ(ИзносОС) ИЛИ Счет = ИзносДругихНеоборотныхМатериальныхАктивовИндивидуально, ,)

Очікуємо:
  - 1121: Дт = 149 202,85
  - 1321: Кт = 149 202,85
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime
import traceback

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'
TEST_DATE = datetime(2026, 3, 1)


def make_time(dt):
    return pywintypes.Time(dt)


def connect_buhbud():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Підключення до BuhBud встановлено")
    return conn


def test_a_cost_with_1121(conn):
    """Тест А: Запит вартості з ИЛИ — має повернути і 10x, і 1121"""
    print("\n" + "=" * 80)
    print("ТЕСТ А: Остатки вартості — ОсновніЗасоби ИЛИ МНМА (1121)")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ\n"
        "  Ост.Счет.Код КАК Счет,\n"
        "  СУММА(Ост.СуммаОстатокДт) КАК СуммаДт\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ОсновныеСредства))\n"
        "      ИЛИ Счет = ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.МалоценныеНеоборотныеМатериальныеАктивыИндивидуально),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "ГДЕ Ост.СуммаОстатокДт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  Ост.Счет.Код"
    )
    query.УстановитьПараметр("Дата", make_time(TEST_DATE))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        found_1121 = False
        sum_1121 = 0
        while result.Следующий():
            count += 1
            acct = str(result.Счет).strip()
            summ = result.СуммаДт or 0
            print(f"  [{count:3d}] Рахунок: {acct:6s} | Дт: {summ:>15.2f}")
            if acct == "1121":
                found_1121 = True
                sum_1121 = summ

        print(f"\n  Всього рахунків: {count}")
        if found_1121:
            print(f"  [OK] Рахунок 1121 ЗНАЙДЕНО! Сума Дт: {sum_1121:>15.2f}")
        else:
            print(f"  [ПОМИЛКА] Рахунок 1121 НЕ знайдено!")
        return found_1121
    except Exception as e:
        print(f"  [ПОМИЛКА] {e}")
        traceback.print_exc()
        return False


def test_b_depreciation_with_1321(conn):
    """Тест Б: Запит зносу з ИЛИ — має повернути і 131, і 1321"""
    print("\n" + "=" * 80)
    print("ТЕСТ Б: Остатки зносу — ИзносОС ИЛИ ИзносМНМА (1321)")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ\n"
        "  Ост.Счет.Код КАК Счет,\n"
        "  СУММА(Ост.СуммаОстатокКт) КАК СуммаКт\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ИзносОсновныхСредств))\n"
        "      ИЛИ Счет = ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ИзносДругихНеоборотныхМатериальныхАктивовИндивидуально),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "ГДЕ Ост.СуммаОстатокКт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  Ост.Счет.Код"
    )
    query.УстановитьПараметр("Дата", make_time(TEST_DATE))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        found_1321 = False
        sum_1321 = 0
        while result.Следующий():
            count += 1
            acct = str(result.Счет).strip()
            summ = result.СуммаКт or 0
            print(f"  [{count:3d}] Рахунок: {acct:6s} | Кт: {summ:>15.2f}")
            if acct == "1321":
                found_1321 = True
                sum_1321 = summ

        print(f"\n  Всього рахунків: {count}")
        if found_1321:
            print(f"  [OK] Рахунок 1321 ЗНАЙДЕНО! Сума Кт: {sum_1321:>15.2f}")
        else:
            print(f"  [ПОМИЛКА] Рахунок 1321 НЕ знайдено!")
        return found_1321
    except Exception as e:
        print(f"  [ПОМИЛКА] {e}")
        traceback.print_exc()
        return False


def test_c_full_query_with_1121(conn):
    """Тест В: Повний запит як у ObjectModule.bsl — перевірка що 1121 повертає дані з Субконто1"""
    print("\n" + "=" * 80)
    print("ТЕСТ В: Повний запит з Субконто1 — перевірка ОС по рахунку 1121")
    print("=" * 80)

    query = conn.NewObject("Запрос")
    query.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 10\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Код КАК ОСКод,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Наименование КАК ОСНаим,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).А_ИдКод КАК А_ИдКод,\n"
        "  Ост.Счет.Код КАК Счет,\n"
        "  СУММА(Ост.СуммаОстатокДт) КАК СуммаДт\n"
        "ИЗ\n"
        "  РегистрБухгалтерии.Хозрасчетный.Остатки(\n"
        "    &Дата,\n"
        "    Счет = ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.МалоценныеНеоборотныеМатериальныеАктивыИндивидуально),\n"
        "    ,\n"
        "  ) КАК Ост\n"
        "ГДЕ\n"
        "  Ост.Субконто1 ССЫЛКА Справочник.ОсновныеСредства\n"
        "  И Ост.СуммаОстатокДт > 0\n"
        "СГРУППИРОВАТЬ ПО\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Код,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).Наименование,\n"
        "  ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ОсновныеСредства).А_ИдКод,\n"
        "  Ост.Счет.Код"
    )
    query.УстановитьПараметр("Дата", make_time(TEST_DATE))

    try:
        result = query.Выполнить().Выбрать()
        count = 0
        total = 0
        while result.Следующий():
            count += 1
            code = str(result.ОСКод).strip() if result.ОСКод else ""
            name = str(result.ОСНаим)[:45] if result.ОСНаим else ""
            a_id = str(result.А_ИдКод).strip() if result.А_ИдКод else ""
            acct = str(result.Счет).strip() if result.Счет else ""
            summ = result.СуммаДт or 0
            total += summ
            print(f"  [{count:3d}] Код:{code:10s} | Рах:{acct} | А_ИдКод:{a_id}")
            print(f"         {name}")
            print(f"         Дт: {summ:>12.2f}")

        print(f"\n  Всього ОС по рахунку 1121: {count}")
        print(f"  Загальна сума Дт: {total:>15.2f}")
        if count > 0:
            print(f"  [OK] Тест В ПРАЦЮЄ!")
        else:
            print(f"  [ПОМИЛКА] Жодного ОС не знайдено по рахунку 1121!")
        return count > 0
    except Exception as e:
        print(f"  [ПОМИЛКА] {e}")
        traceback.print_exc()
        return False


def main():
    print("=" * 80)
    print("  Тест: Перевірка запитів з ИЛИ для рахунку 1121 (МНМА)")
    print(f"  Дата залишків: {TEST_DATE:%d.%m.%Y}")
    print("=" * 80)

    conn = None
    results = {}
    try:
        conn = connect_buhbud()
        results["A"] = test_a_cost_with_1121(conn)
        results["B"] = test_b_depreciation_with_1321(conn)
        results["C"] = test_c_full_query_with_1121(conn)
    except Exception as e:
        print(f"\n[КРИТИЧНА ПОМИЛКА] {e}")
        traceback.print_exc()
    finally:
        conn = None
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    print("\n" + "=" * 80)
    print("  ПІДСУМОК:")
    for test_name, passed in results.items():
        status = "PASSED" if passed else "FAILED"
        print(f"    Тест {test_name}: {status}")
    print("=" * 80)


if __name__ == "__main__":
    main()
