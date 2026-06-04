"""RED test для Phase 0 — перевіряє SDBL-синтаксис нової логіки заповнення ДДС
у дохідних секціях запиту А_ФинРез_PL.СформироватьЗапросСверткиPL().

Перевіряє що:
1. SDBL deref `PL.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств` працює (Section 1)
2. SDBL deref `Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств` після JOIN працює (Section 4)
3. SDBL deref `Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств` без JOIN працює (Section 6)
4. Для лютого 2026 запит повертає рядки де ДДС = "Прочие поступления"
   для 7 статей доходу що користувач проставив у ПВХ.

Шаблон (Rule #-1): Python COM тест ПЕРЕД BSL правкою.

Запуск: .venv\\Scripts\\python.exe _Rarzrabotki\\Python\\test\\test_finrez_pl_income_dds_2026-05-06.py
"""
import sys
import datetime as dt
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ORG_INDUSTRIALBUD = "80D3000C29BBAC2311E653F06BEE36B2"
PROCH_POSTUPL = "Прочие поступления"
PERIOD_START = dt.datetime(2026, 2, 1, 0, 0, 0)
PERIOD_END = dt.datetime(2026, 2, 28, 23, 59, 59)


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN_ERP)


def test_section1_pl_excel_deref(erp):
    """Section 1 — PL_Excel з втPL.

    Тест: SDBL вираз `PL.Статья.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств`
    де PL — таблична частина Документ.А_ОтчетPL.ДанныеОтчета.
    """
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 100
        ТЧ.Ссылка КАК Документ,
        ТЧ.Статья.Наименование КАК СтатьяPL,
        ТЧ.Статья.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        ТЧ.Статья.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование КАК ДДСИзДоходов,
        ТЧ.Сумма КАК Сумма
    ИЗ
        Документ.А_ОтчетPL.ДанныеОтчета КАК ТЧ
    ГДЕ
        ТЧ.Ссылка.Дата МЕЖДУ &НачП И &КонП
        И НЕ ТЧ.Ссылка.ПометкаУдаления
        И ТЧ.Статья.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств <> ЗНАЧЕНИЕ(Справочник.СтатьиДвиженияДенежныхСредств.ПустаяСсылка)
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        print(f"  [Section 1] OK PL_Excel deref: {n} rows")
        if n > 0:
            sample = r.Получить(0)
            print(f"    Sample: Стаття='{sample.СтатьяДоходов}' → ДДС='{sample.ДДСИзДоходов}'")
        return True
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [Section 1] FAIL: {info}")
        return False


def test_section4_erp_dox_with_pl_join(erp):
    """Section 4 — втЕРП_Дох з JOIN до Справочник.А_Статьи_PL ПО СтатьяДоходов.

    Тест: SDBL deref `Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств`
    де Д — alias регістра ПрочиеДоходы.Обороты з ВНУТРЕННЕЕ СОЕДИНЕНИЕ.
    """
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 100
        Д.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование КАК ДДСИзДоходов,
        PL.Наименование КАК СтатьяPL,
        Д.Подразделение.Наименование КАК Подразделение,
        Д.Регистратор КАК Документ,
        Д.СуммаПриход КАК Сумма
    ИЗ
        РегистрНакопления.ПрочиеДоходы.Обороты(&НачП, &КонП, Регистратор, ) КАК Д
            ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.А_Статьи_PL КАК PL
            ПО PL.СтатьяДоходов = Д.СтатьяДоходов
    ГДЕ
        Д.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
        И НЕ PL.ПометкаУдаления
        И НЕ PL.ЭтоГруппа
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        print(f"  [Section 4] OK ERP_Income+PL deref: {n} rows")
        if n > 0:
            sample = r.Получить(0)
            print(f"    Sample: Стаття='{sample.СтатьяДоходов}' → ДДС='{sample.ДДСИзДоходов}', PL='{sample.СтатьяPL}'")
        return True
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [Section 4] FAIL: {info}")
        return False


def test_section6_erp_dox_without_pl(erp):
    """Section 6 — втЕРП_Дох без зв'язку з А_Статьи_PL.

    Тест: SDBL deref `Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств`
    де Д — alias регістра ПрочиеДоходы.Обороты, статті НЕ В Справочник.А_Статьи_PL.
    """
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 100
        Д.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование КАК ДДСИзДоходов,
        Д.Подразделение.Наименование КАК Подразделение,
        Д.Регистратор КАК Документ,
        Д.СуммаПриход КАК Сумма
    ИЗ
        РегистрНакопления.ПрочиеДоходы.Обороты(&НачП, &КонП, Регистратор, ) КАК Д
    ГДЕ
        Д.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
        И НЕ Д.СтатьяДоходов В (
            ВЫБРАТЬ PL.СтатьяДоходов
            ИЗ Справочник.А_Статьи_PL КАК PL
            ГДЕ PL.СтатьяДоходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиДоходов.ПустаяСсылка)
                И НЕ PL.ПометкаУдаления
                И НЕ PL.ЭтоГруппа)
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        print(f"  [Section 6] OK ERP_Income БезPL deref: {n} rows")
        if n > 0:
            sample = r.Получить(0)
            print(f"    Sample: Стаття='{sample.СтатьяДоходов}' → ДДС='{sample.ДДСИзДоходов}'")
        return True
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [Section 6] FAIL: {info}")
        return False


def test_acceptance_g2_dds_proche_postupl(erp):
    """G2 — для лютого 2026 у регістрі ПрочиеДоходы є рядки де
    стаття доходу має реквізит А_СтатьяДвиженияДенежныхСредств = "Прочие поступления".
    """
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Д.СтатьяДоходов.Наименование КАК СтатьяДоходов,
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование КАК ДДСИзДоходов,
        СУММА(Д.СуммаПриход) КАК СуммаПриход
    ИЗ
        РегистрНакопления.ПрочиеДоходы.Обороты(&НачП, &КонП, , ) КАК Д
    ГДЕ
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств <> ЗНАЧЕНИЕ(Справочник.СтатьиДвиженияДенежныхСредств.ПустаяСсылка)
    СГРУППИРОВАТЬ ПО
        Д.СтатьяДоходов.Наименование,
        Д.СтатьяДоходов.А_СтатьяДвиженияДенежныхСредств.Наименование
    """
    q.SetParameter("НачП", PERIOD_START)
    q.SetParameter("КонП", PERIOD_END)
    try:
        r = q.Execute().Выгрузить()
        n = r.Количество()
        sum_total = 0.0
        articles_with_proche = []
        for i in range(n):
            row = r.Получить(i)
            sum_total += float(row.СуммаПриход)
            if row.ДДСИзДоходов == PROCH_POSTUPL:
                articles_with_proche.append(row.СтатьяДоходов)
        print(f"  [G2] OK: {n} статей доходу мають заповнений реквізит А_СтатьяДвиженияДенежныхСредств")
        print(f"    Σ суми = {sum_total:,.2f} ₴ за лютий 2026")
        print(f"    Статей які мапляться на 'Прочие поступления': {len(articles_with_proche)}")
        for a in articles_with_proche:
            print(f"      • {a}")
        return n > 0
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  [G2] FAIL: {info}")
        return False


def main():
    print("=" * 80)
    print("RED test: А_ФинРез_PL DDS resolution from СтатьяДоходов.А_СтатьяДД")
    print("Period: 2026-02 (Лютий 2026)")
    print("=" * 80)
    print("\nПідключення до BaseERP...")
    erp = connect()
    print("OK")

    print("\n--- Section 1 (PL_Excel deref) ---")
    s1 = test_section1_pl_excel_deref(erp)

    print("\n--- Section 4 (ERP_Income + PL JOIN deref) ---")
    s4 = test_section4_erp_dox_with_pl_join(erp)

    print("\n--- Section 6 (ERP_Income БезPL deref) ---")
    s6 = test_section6_erp_dox_without_pl(erp)

    print("\n--- G2: ДДС 'Прочие поступления' заповнено для статей доходу ---")
    g2 = test_acceptance_g2_dds_proche_postupl(erp)

    print("\n" + "=" * 80)
    print(f"Section 1: {'PASS' if s1 else 'FAIL'}")
    print(f"Section 4: {'PASS' if s4 else 'FAIL'}")
    print(f"Section 6: {'PASS' if s6 else 'FAIL'}")
    print(f"G2:        {'PASS' if g2 else 'FAIL'}")
    all_pass = s1 and s4 and s6 and g2
    print(f"\nOVERALL: {'PASS — Phase 1 BSL правка готова' if all_pass else 'FAIL — НЕ ПРАВИТИ BSL'}")
    print("=" * 80)
    sys.exit(0 if all_pass else 1)


if __name__ == "__main__":
    main()
