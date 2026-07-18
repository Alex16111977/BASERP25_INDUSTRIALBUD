# -*- coding: utf-8 -*-
"""
Test: check data for filling OtrazhenieSalaryReflection
Checks source data in A_RaspredelenijeKazna, A_RaspredelenijeNalogi
and register VzaimoraschyotySoSotrudnikami
"""

import win32com.client
import sys
import datetime

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="\u0410\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440";Pwd="24043"'

def main():
    print("=" * 60)
    print("Test: salary reflection data check")
    print("=" * 60)

    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONN_ERP)

    # Ищем документ ОтражениеЗарплатыВФинансовомУчете за декабрь 2025
    query = conn.NewObject("Запрос")
    query.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 5
        Док.Ссылка КАК Ссылка,
        Док.Дата КАК Дата,
        Док.Номер КАК Номер,
        Док.ПериодРегистрации КАК ПериодРегистрации,
        Док.Организация КАК Организация,
        Док.А_РаспределениеКазна.Количество() КАК КолКазна,
        Док.А_РаспределениеНалоги.Количество() КАК КолНалоги,
        Док.НачисленнаяЗарплатаИВзносыПоФизлицам.Количество() КАК КолНачисленияПоФЛ,
        Док.НачисленнаяЗарплатаИВзносы.Количество() КАК КолНачисления,
        Док.НачисленныйНДФЛ.Количество() КАК КолНДФЛ
    ИЗ
        Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док
    ГДЕ
        Док.ПериодРегистрации >= &НачДата
        И Док.ПериодРегистрации <= &КонДата
        И НЕ Док.ПометкаУдаления
    УПОРЯДОЧИТЬ ПО
        Док.Дата УБЫВ
    """

    query.SetParameter("\u041d\u0430\u0447\u0414\u0430\u0442\u0430", datetime.datetime(2025, 10, 1))
    query.SetParameter("\u041a\u043e\u043d\u0414\u0430\u0442\u0430", datetime.datetime(2026, 4, 1))

    result = query.Execute()

    if result.IsEmpty():
        print("\nНет документов ОтражениеЗарплатыВФинансовомУчете в периоде!")
        return

    sel = result.Choose()

    print("\nНайденные документы:")
    print("-" * 80)
    doc_ref = None
    while sel.Next():
        period_str = str(sel.ПериодРегистрации)[:10]
        print(f"  №{sel.Номер} от {str(sel.Дата)[:10]}, Период: {period_str}")
        print(f"    Организация: {sel.Организация}")
        print(f"    А_РаспределениеКазна: {sel.КолКазна} строк")
        print(f"    А_РаспределениеНалоги: {sel.КолНалоги} строк")
        print(f"    НачисленнаяЗарплатаИВзносыПоФизлицам: {sel.КолНачисленияПоФЛ} строк")
        print(f"    НачисленнаяЗарплатаИВзносы: {sel.КолНачисления} строк")
        print(f"    НачисленныйНДФЛ: {sel.КолНДФЛ} строк")
        print()

        # Берем первый подходящий документ (с заполненными источниками)
        if doc_ref is None and sel.КолКазна > 0 and sel.КолНалоги > 0:
            doc_ref = sel.Ссылка

    if doc_ref is None:
        print("Не найден документ с заполненными А_РаспределениеКазна и А_РаспределениеНалоги!")
        return

    print("=" * 60)
    print(f"Анализ документа: {doc_ref}")
    print("=" * 60)

    doc_obj = doc_ref.GetObject()

    # Анализ А_РаспределениеКазна
    print(f"\n--- А_РаспределениеКазна ({doc_obj.А_РаспределениеКазна.Count()} строк) ---")
    employees = set()
    depts = set()
    total_sum = 0
    total_tax = 0
    for i in range(doc_obj.А_РаспределениеКазна.Count()):
        row = doc_obj.А_РаспределениеКазна.Get(i)
        employees.add(str(row.Сотрудник))
        depts.add(str(row.Подразделение))
        total_sum += float(str(row.Сумма))
        total_tax += float(str(row.СуммаНалогов))

    print(f"  Уникальных сотрудников: {len(employees)}")
    print(f"  Уникальных подразделений: {len(depts)}")
    print(f"  Итого Сумма: {total_sum:.2f}")
    print(f"  Итого СуммаНалогов: {total_tax:.2f}")

    # Первые 3 строки для примера
    print("  Примеры строк:")
    for i in range(min(3, doc_obj.А_РаспределениеКазна.Count())):
        row = doc_obj.А_РаспределениеКазна.Get(i)
        stat = str(row.СтатьяДвиженияДенежныхСредств)
        print(f"    [{i}] Сотрудник={str(row.Сотрудник)[:30]}, Подразделение={str(row.Подразделение)[:20]}, "
              f"Сумма={row.Сумма}, СтатьяДДС={stat[:30]}")

    # Анализ А_РаспределениеНалоги
    print(f"\n--- А_РаспределениеНалоги ({doc_obj.А_РаспределениеНалоги.Count()} строк) ---")
    total_sum_tax = 0
    total_sum_fot = 0
    count_ndfl = 0
    count_esv = 0
    for i in range(doc_obj.А_РаспределениеНалоги.Count()):
        row = doc_obj.А_РаспределениеНалоги.Get(i)
        s = float(str(row.Сумма))
        f = float(str(row.СуммаФОТ))
        total_sum_tax += s
        total_sum_fot += f
        if s > 0 and f == 0:
            count_ndfl += 1
        if f > 0:
            count_esv += 1

    print(f"  Строк НДФЛ/ВС (Сумма > 0, СуммаФОТ = 0): {count_ndfl}")
    print(f"  Строк ЕСВ (СуммаФОТ > 0): {count_esv}")
    print(f"  Итого Сумма: {total_sum_tax:.2f}")
    print(f"  Итого СуммаФОТ: {total_sum_fot:.2f}")

    # Проверяем регистр ВзаиморасчетыССотрудниками
    period_reg = doc_obj.ПериодРегистрации
    print(f"\n--- Регистр ВзаиморасчетыССотрудниками (Период: {str(period_reg)[:10]}) ---")

    query2 = conn.NewObject("Запрос")
    query2.Text = """
    ВЫБРАТЬ
        КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВзаиморасчетыССотрудниками.Сотрудник) КАК КолСотрудников,
        СУММА(ВзаиморасчетыССотрудниками.СуммаВзаиморасчетов) КАК ИтогоСумма
    ИЗ
        РегистрНакопления.ВзаиморасчетыССотрудниками КАК ВзаиморасчетыССотрудниками
    ГДЕ
        ВзаиморасчетыССотрудниками.Период МЕЖДУ &НачалоПериода И &КонецПериода
        И ВзаиморасчетыССотрудниками.Регистратор ССЫЛКА Документ.НачислениеЗарплаты
    """

    period_str = str(period_reg)
    year = int(period_str[:4])
    month = int(period_str[5:7])
    beg_month = conn.NewObject("Дата", year, month, 1, 0, 0, 0)
    if month == 12:
        end_month = conn.NewObject("Дата", year + 1, 1, 1, 0, 0, 0)
    else:
        end_month = conn.NewObject("Дата", year, month + 1, 1, 0, 0, 0)

    query2.SetParameter("НачалоПериода", beg_month)
    query2.SetParameter("КонецПериода", end_month)

    result2 = query2.Execute()
    sel2 = result2.Choose()
    if sel2.Next():
        print(f"  Сотрудников с начислениями: {sel2.КолСотрудников}")
        print(f"  Итого СуммаВзаиморасчетов: {sel2.ИтогоСумма}")

    # Текущие целевые таблицы
    print(f"\n--- Текущие целевые таблицы ---")
    print(f"  НачисленнаяЗарплатаИВзносыПоФизлицам: {doc_obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()} строк")
    print(f"  НачисленнаяЗарплатаИВзносы: {doc_obj.НачисленнаяЗарплатаИВзносы.Count()} строк")
    print(f"  НачисленныйНДФЛ: {doc_obj.НачисленныйНДФЛ.Count()} строк")

    print("\n" + "=" * 60)
    print("ТЕСТ ПРОЙДЕН: данные для заполнения присутствуют")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nОШИБКА: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
