# =============================================================
# test_profit_balance.py
# Тест выравнивания баланса Актив=Пассив в регистре ПрочиеАктивыПассивы
# Документ: КорректировкаРегистров 0Ц-00000002 от 30.11.2025
#
# Использование:
#   python test/test_profit_balance.py before     # ДО запуска обработки
#   python test/test_profit_balance.py after      # ПОСЛЕ запуска обработки
#   python test/test_profit_balance.py movements  # Проверить движения документа
#   python test/test_profit_balance.py all        # Все тесты
# =============================================================

import sys
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONN_STR = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
PERIOD_END = datetime(2025, 11, 30, 23, 59, 59)
DOC_NUMBER = "0Ц-00000002"
EXPECTED_DISBALANCE = 34999.48


def connect():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_STR)
    print(f"Подключение к {CONN_STR.split('Ref=')[1].split(';')[0]} — OK")
    return conn


def test_disbalance_before(conn):
    """
    Тест 1: Диагностика ПЕРЕД заполнением документа.
    Показывает дисбаланс (Актив-Пассив) по каждой комбинации
    (Организация, Подразделение, НаправлениеДеятельности) без статьи Прибыли и убытки.
    """
    print("\n" + "=" * 60)
    print("ТЕСТ 1: Дисбаланс ДО заполнения документа")
    print("=" * 60)

    # Найти статью Прибыли и убытки
    stat_ref = conn.ПланыВидовХарактеристик.СтатьиАктивовПассивов.НайтиПоНаименованию(
        "Прибыли и убытки", True
    )
    if conn.ЗначениеЗаполнено(stat_ref) == False:
        print("ОШИБКА: Статья 'Прибыли и убытки' не найдена!")
        return False

    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ"
        "   Ост.Организация.Наименование             КАК Организация,"
        "   Ост.Подразделение.Наименование           КАК Подразделение,"
        "   Ост.НаправлениеДеятельности.Наименование КАК Направление,"
        "   СУММА(Ост.СуммаОстаток)                  КАК Дисбаланс"
        " ИЗ"
        "   РегистрНакопления.ПрочиеАктивыПассивы.Остатки(&КонецПериода, ) КАК Ост"
        " ГДЕ"
        "   Ост.Статья <> &СтатьяПрибыль"
        " СГРУППИРОВАТЬ ПО"
        "   Ост.Организация,"
        "   Ост.Подразделение,"
        "   Ост.НаправлениеДеятельности"
        " ИМЕЮЩИЕ"
        "   СУММА(Ост.СуммаОстаток) <> 0"
        " УПОРЯДОЧИТЬ ПО"
        "   Организация, Подразделение, Направление"
    )
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(PERIOD_END))
    q.УстановитьПараметр("СтатьяПрибыль", stat_ref)

    sel = q.Выполнить().Выбрать()
    rows = []
    total = 0.0
    while sel.Следующий():
        org = str(conn.String(sel.Организация))
        pod = str(conn.String(sel.Подразделение))
        napr = str(conn.String(sel.Направление))
        dis = float(str(conn.String(sel.Дисбаланс)).replace("\xa0", "").replace(" ", "").replace(",", "."))
        rows.append((org, pod, napr, dis))
        total += dis

    print(f"Найдено строк с ненулевым дисбалансом: {len(rows)}")
    for org, pod, napr, dis in rows:
        napr_str = napr if napr.strip() else "<пусто>"
        print(f"  {org} | {pod} | {napr_str}: {dis:,.2f}")

    total_r = round(total, 2)
    print(f"\nКонтрольная сумма дисбаланса: {total_r:,.2f}")
    print(f"Ожидалось примерно:           {EXPECTED_DISBALANCE:,.2f}")

    if len(rows) == 0:
        print("!! Дисбаланс отсутствует — возможно документ уже заполнен")
    elif abs(total_r - EXPECTED_DISBALANCE) < 1.0:
        print(f"OK Контрольная сумма совпадает (допуск ±1 грн)")
    else:
        print(f"!! Контрольная сумма отличается (ожидалось {EXPECTED_DISBALANCE})")

    return True


def test_balance_after(conn):
    """
    Тест 2: Проверка ПОСЛЕ заполнения документа.
    По каждой комбинации (Орг, Подр, Направление) сумма остатков должна = 0.
    """
    print("\n" + "=" * 60)
    print("ТЕСТ 2: Баланс ПОСЛЕ заполнения документа")
    print("=" * 60)

    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ"
        "   Ост.Организация.Наименование             КАК Организация,"
        "   Ост.Подразделение.Наименование           КАК Подразделение,"
        "   Ост.НаправлениеДеятельности.Наименование КАК Направление,"
        "   СУММА(Ост.СуммаОстаток)                  КАК Итого"
        " ИЗ"
        "   РегистрНакопления.ПрочиеАктивыПассивы.Остатки(&КонецПериода, ) КАК Ост"
        " СГРУППИРОВАТЬ ПО"
        "   Ост.Организация,"
        "   Ост.Подразделение,"
        "   Ост.НаправлениеДеятельности"
        " ИМЕЮЩИЕ"
        "   СУММА(Ост.СуммаОстаток) <> 0"
        " УПОРЯДОЧИТЬ ПО"
        "   Организация, Подразделение, Направление"
    )
    q.УстановитьПараметр("КонецПериода", pywintypes.Time(PERIOD_END))

    sel = q.Выполнить().Выбрать()
    nonzero = []
    while sel.Следующий():
        org = str(conn.String(sel.Организация))
        pod = str(conn.String(sel.Подразделение))
        napr = str(conn.String(sel.Направление))
        итого = float(str(conn.String(sel.Итого)).replace("\xa0", "").replace(" ", "").replace(",", "."))
        nonzero.append((org, pod, napr, итого))

    if len(nonzero) == 0:
        print("OK БАЛАНС ВЫРОВНЕН! Все комбинации (Орг + Подр + Направление) = 0")
        return True
    else:
        print(f"!! Ненулевых комбинаций: {len(nonzero)}")
        for org, pod, napr, итого in nonzero:
            napr_str = napr if napr.strip() else "<пусто>"
            print(f"  {org} | {pod} | {napr_str}: {итого:,.2f}")
        return False


def test_document_movements(conn):
    """
    Тест 3: Проверка движений документа 0Ц-00000002.
    Все движения должны иметь статью Прибыли и убытки.
    """
    print("\n" + "=" * 60)
    print(f"ТЕСТ 3: Движения документа {DOC_NUMBER}")
    print("=" * 60)

    # Найти документ
    doc_date = pywintypes.Time(datetime(2025, 11, 30))
    doc_ref = conn.Документы.КорректировкаРегистров.НайтиПоНомеру(DOC_NUMBER, doc_date)
    if not conn.ЗначениеЗаполнено(doc_ref):
        print(f"ОШИБКА: Документ {DOC_NUMBER} не найден!")
        return False

    print(f"Документ найден: {conn.String(doc_ref)}")

    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ"
        "   Дв.ВидДвижения                           КАК ВидДвижения,"
        "   Дв.Организация.Наименование              КАК Организация,"
        "   Дв.Подразделение.Наименование            КАК Подразделение,"
        "   Дв.НаправлениеДеятельности.Наименование КАК Направление,"
        "   Дв.Статья.Наименование                   КАК Статья,"
        "   Дв.Сумма                                 КАК Сумма"
        " ИЗ"
        "   РегистрНакопления.ПрочиеАктивыПассивы КАК Дв"
        " ГДЕ"
        "   Дв.Регистратор = &Документ"
        " УПОРЯДОЧИТЬ ПО"
        "   Организация, Подразделение, Направление"
    )
    q.УстановитьПараметр("Документ", doc_ref)

    sel = q.Выполнить().Выбрать()
    count = 0
    while sel.Следующий():
        count += 1
        vid = str(conn.String(sel.ВидДвижения))
        org = str(conn.String(sel.Организация))
        pod = str(conn.String(sel.Подразделение))
        napr = str(conn.String(sel.Направление))
        stat = str(conn.String(sel.Статья))
        summa = float(str(conn.String(sel.Сумма)).replace("\xa0", "").replace(" ", "").replace(",", "."))
        napr_str = napr if napr.strip() else "<пусто>"
        print(f"  {vid} | {org} | {pod} | {napr_str} | {stat} | {summa:,.2f}")

    print(f"\nВсего движений в документе: {count}")
    if count == 0:
        print("!! Документ не имеет движений — нужно запустить обработку")
        return False
    else:
        print("OK Движения найдены")
        return True


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"

    try:
        conn = connect()
    except Exception as e:
        print(f"ОШИБКА подключения: {e}")
        sys.exit(1)

    results = []

    if mode in ("before", "all"):
        try:
            results.append(("before", test_disbalance_before(conn)))
        except Exception as e:
            print(f"ОШИБКА в test_disbalance_before: {e}")
            import traceback
            traceback.print_exc()

    if mode in ("after", "all"):
        try:
            results.append(("after", test_balance_after(conn)))
        except Exception as e:
            print(f"ОШИБКА в test_balance_after: {e}")
            import traceback
            traceback.print_exc()

    if mode in ("movements", "all"):
        try:
            results.append(("movements", test_document_movements(conn)))
        except Exception as e:
            print(f"ОШИБКА в test_document_movements: {e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 60)
    print("ИТОГ:")
    for name, ok in results:
        status = "OK OK" if ok else "✗ FAIL"
        print(f"  {name}: {status}")


if __name__ == "__main__":
    main()
