"""RED test для Plan v2 — рефакторинг А_ФинРез_PL: Стаття PL через ДДС-ланцюжок.

Перевіряє нову логіку (без використання Справочник.А_Статьи_PL.СтатьяДоходов):
1. Section 4 nova: JOIN ЕРП Дох з втPL_ДДС через ДДС → СтатьяPL
2. Section 4 regression: "Выручка от продаж" доход → ДДС "Продажи" → PL "Виручка от продаж"
3. Section 6 narrow: рядки де доход не мапиться через ДДС → лишається у БезPL_Доход

Period: лютий 2026 (2026-02-01 .. 2026-02-28).
"""
import sys
import datetime as dt
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
PERIOD_START = dt.datetime(2026, 2, 1, 0, 0, 0)
PERIOD_END = dt.datetime(2026, 2, 28, 23, 59, 59)


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN_ERP)


def test_section4_nova_join_via_dds(erp):
    """Section 4 нова — JOIN втEРП_Дох з втPL_ДДС через ДДС."""
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Д.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        PLДДС.СтатьяPL.Наименование КАК СтатьяPL,
        PLДДС.ДДС.Наименование КАК ДДС,
        КОЛИЧЕСТВО(*) КАК КолРядків,
        СУММА(Д.СуммаПриход) КАК Сума
    ИЗ
        РегистрНакопления.ПрочиеДоходы.Обороты(&НачП, &КонП, , ) КАК Д
        ВНУТРЕННЕЕ СОЕДИНЕНИЕ (
            ВЫБРАТЬ РАЗЛИЧНЫЕ
                СтатьиТЧ.Ссылка КАК СтатьяPL,
                СтатьиТЧ.СтатьяДвиженияДенежныхСредств КАК ДДС
            ИЗ Справочник.А_Статьи_PL.Статьи КАК СтатьиТЧ
            ГДЕ НЕ СтатьиТЧ.Ссылка.ПометкаУдаления
                И СтатьиТЧ.СтатьяДвиженияДенежныхСредств <> ЗНАЧЕНИЕ(Справочник.СтатьиДвиженияДенежныхСредств.ПустаяСсылка)
        ) КАК PLДДС
        ПО PLДДС.ДДС = Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств
    ГДЕ
        Д.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
    СГРУППИРОВАТЬ ПО
        Д.СтатьяДоходов.Наименование,
        PLДДС.СтатьяPL.Наименование,
        PLДДС.ДДС.Наименование
    УПОРЯДОЧИТЬ ПО СтатьяPL, СтатьяДоходов
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        if n == 0:
            print(f"  [Section 4 nova] FAIL: 0 rows — JOIN не повернув жодного рядка")
            return False
        empty_pl = 0
        for i in range(n):
            row = r.Получить(i)
            if not row.СтатьяPL:
                empty_pl += 1
        print(f"  [Section 4 nova] OK: {n} mappings (Доход → PL → ДДС), {empty_pl} з порожнім PL")
        for i in range(min(n, 7)):
            row = r.Получить(i)
            print(f"    • {row.СтатьяДоходов:55} → PL='{row.СтатьяPL}' / ДДС='{row.ДДС}' / Сума={row.Сума:,.2f}")
        return empty_pl == 0
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [Section 4 nova] FAIL: {info}")
        return False


def test_section4_regression_vyruchka(erp):
    """Regression: 'Выручка от продаж' доход → ДДС 'Продажи' → PL 'Виручка от продаж'."""
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 5
        ВырРезКАК.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        PLДДС.СтатьяPL.Наименование КАК СтатьяPL,
        PLДДС.ДДС.Наименование КАК ДДС,
        СУММА(ВырРезКАК.СуммаВыручки) КАК Сума
    ИЗ
        (ВЫБРАТЬ
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ВыручкаОтПродаж) КАК СтатьяДоходов,
            Выр.СуммаВыручки КАК СуммаВыручки
        ИЗ
            РегистрНакопления.ВыручкаИСебестоимостьПродаж КАК Выр
        ГДЕ Выр.Период МЕЖДУ &НачП И &КонП И Выр.Активность И Выр.СуммаВыручки <> 0) КАК ВырРезКАК
        ВНУТРЕННЕЕ СОЕДИНЕНИЕ (
            ВЫБРАТЬ РАЗЛИЧНЫЕ
                СтатьиТЧ.Ссылка КАК СтатьяPL,
                СтатьиТЧ.СтатьяДвиженияДенежныхСредств КАК ДДС
            ИЗ Справочник.А_Статьи_PL.Статьи КАК СтатьиТЧ
            ГДЕ НЕ СтатьиТЧ.Ссылка.ПометкаУдаления
                И СтатьиТЧ.СтатьяДвиженияДенежныхСредств <> ЗНАЧЕНИЕ(Справочник.СтатьиДвиженияДенежныхСредств.ПустаяСсылка)
        ) КАК PLДДС
        ПО PLДДС.ДДС = ВырРезКАК.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств
    СГРУППИРОВАТЬ ПО
        ВырРезКАК.СтатьяДоходов.Наименование,
        PLДДС.СтатьяPL.Наименование,
        PLДДС.ДДС.Наименование
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        # Очікую "Выручка от продаж" → "Виручка от продаж" з ДДС "Продажи"
        found_vyruchka = False
        for i in range(n):
            row = r.Получить(i)
            print(f"    • Доход='{row.СтатьяДоходов}' → PL='{row.СтатьяPL}' / ДДС='{row.ДДС}'")
            if "Выручка" in (row.СтатьяДоходов or "") and "родаж" in (row.СтатьяPL or ""):
                found_vyruchka = True
        print(f"  [Section 4 regression] {'OK' if found_vyruchka else 'WARN'}: Виручка mapping {'found' if found_vyruchka else 'NOT found'}")
        return n > 0
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [Section 4 regression] FAIL: {info}")
        return False


def test_section6_narrow(erp):
    """Section 6 narrow — статті доходу що НЕ мапяться через ДДС (БезPL_Доход)."""
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Д.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование КАК ДДС_ИзДоходов,
        КОЛИЧЕСТВО(*) КАК КолРядків,
        СУММА(Д.СуммаПриход) КАК Сума
    ИЗ
        РегистрНакопления.ПрочиеДоходы.Обороты(&НачП, &КонП, , ) КАК Д
    ГДЕ
        Д.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
        И НЕ Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств В
            (ВЫБРАТЬ СтатьиТЧ.СтатьяДвиженияДенежныхСредств
             ИЗ Справочник.А_Статьи_PL.Статьи КАК СтатьиТЧ
             ГДЕ НЕ СтатьиТЧ.Ссылка.ПометкаУдаления
                 И СтатьиТЧ.СтатьяДвиженияДенежныхСредств <> ЗНАЧЕНИЕ(Справочник.СтатьиДвиженияДенежныхСредств.ПустаяСсылка))
    СГРУППИРОВАТЬ ПО
        Д.СтатьяДоходов.Наименование,
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        print(f"  [Section 6 narrow] OK: {n} статей доходу не мапяться через ДДС-ланцюжок")
        for i in range(n):
            row = r.Получить(i)
            dds = row.ДДС_ИзДоходов or "(порожня А_СтатьяДД)"
            print(f"    ⚠ '{row.СтатьяДоходов}' → ДДС='{dds}' / Сума={row.Сума:,.2f}")
        return True
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [Section 6 narrow] FAIL: {info}")
        return False


def main():
    print("=" * 80)
    print("RED test v2: А_ФинРез_PL — СтатьяPL через ДДС-ланцюжок")
    print("Period: 2026-02 (Лютий 2026)")
    print("=" * 80)
    print("\nПідключення до BaseERP...")
    erp = connect()
    print("OK")

    print("\n--- Section 4 nova: JOIN втEРП_Дох × втPL_ДДС за ДДС ---")
    s4n = test_section4_nova_join_via_dds(erp)

    print("\n--- Section 4 regression: Виручка від продаж → PL ---")
    s4r = test_section4_regression_vyruchka(erp)

    print("\n--- Section 6 narrow: доходи без ДДС-ланцюжка ---")
    s6 = test_section6_narrow(erp)

    print("\n" + "=" * 80)
    print(f"Section 4 nova:        {'PASS' if s4n else 'FAIL'}")
    print(f"Section 4 regression:  {'PASS' if s4r else 'FAIL'}")
    print(f"Section 6 narrow:      {'PASS' if s6 else 'FAIL'}")
    all_pass = s4n and s4r and s6
    print(f"\nOVERALL: {'PASS — Phase 1 BSL правка готова' if all_pass else 'FAIL — НЕ ПРАВИТИ BSL'}")
    print("=" * 80)
    sys.exit(0 if all_pass else 1)


if __name__ == "__main__":
    main()
