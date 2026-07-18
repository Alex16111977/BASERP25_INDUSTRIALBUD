# -*- coding: utf-8 -*-
"""
Тесты запросов для обработки ЗакрытиеОтрицательныхОстатков (Rule #-1: Python COM перед BSL).

4 теста:
1. test_query_negative_balances — запрос отрицательных остатков возвращает >0 строк
2. test_query_avg_prices — запрос средних цен из ПриобретениеТоваровУслуг работает
3. test_statya_dohodov_exists — ПВХ.СтатьиДоходов.НайтиПоКоду("000001007") даёт ссылку
4. test_sklad_podrazdelenie — все склады из остатков имеют непустое Подразделение
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ORG_NAME = 'ТОВ "ІНДАСТРІАЛБУД"'
STATYA_KOD = "000001007"


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN)


def make_query(erp, text, params=None):
    q = erp.NewObject("Запрос")
    q.Text = text
    for k, v in (params or {}).items():
        q.SetParameter(k, v)
    return q.Execute().Выгрузить()


def test_query_negative_balances(erp):
    text = """
    ВЫБРАТЬ
        Ост.Организация,
        Ост.АналитикаУчетаНоменклатуры,
        Ост.АналитикаУчетаНоменклатуры.Номенклатура         КАК Номенклатура,
        Ост.АналитикаУчетаНоменклатуры.Характеристика       КАК Характеристика,
        Ост.АналитикаУчетаНоменклатуры.СкладскаяТерритория  КАК Склад,
        Ост.ВидЗапасов,
        Ост.НомерГТД,
        -Ост.КоличествоОстаток                              КАК Количество
    ИЗ
        РегистрНакопления.ТоварыОрганизаций.Остатки(&НаДату, ) КАК Ост
    ГДЕ
        Ост.КоличествоОстаток < 0
    """
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ ДАТАВРЕМЯ(2026,4,30,23,59,59) КАК Д"
    end_of_april = q.Execute().Выгрузить().Получить(0).Д

    rows = make_query(erp, text, {"НаДату": end_of_april})
    n = rows.Количество()
    print(f"  test_query_negative_balances: rows={n}")
    assert n > 0, "Ожидались отрицательные остатки, но запрос вернул 0 строк"
    return rows


def test_query_avg_prices(erp, nomenklatura_refs):
    """Средняя цена из ПриобретениеТоваровУслуг.Товары за 12 мес."""
    text = """
    ВЫБРАТЬ
        Тов.Номенклатура КАК Номенклатура,
        СУММА(Тов.Количество) КАК Количество,
        СУММА(Тов.Сумма)      КАК Сумма,
        СУММА(Тов.Сумма) / ВЫБОР КОГДА СУММА(Тов.Количество) = 0 ТОГДА 1 ИНАЧЕ СУММА(Тов.Количество) КОНЕЦ КАК СредняяЦена
    ИЗ Документ.ПриобретениеТоваровУслуг.Товары КАК Тов
    ГДЕ Тов.Ссылка.Проведен = ИСТИНА
        И Тов.Ссылка.Дата >= &ДатаС
        И Тов.Ссылка.Дата <= &ДатаПо
        И Тов.Номенклатура В (&Номенклатуры)
    СГРУППИРОВАТЬ ПО Тов.Номенклатура
    """
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ ДАТАВРЕМЯ(2025,5,1) КАК Д1, ДАТАВРЕМЯ(2026,4,30,23,59,59) КАК Д2"
    r = q.Execute().Выгрузить().Получить(0)
    dat_s, dat_po = r.Д1, r.Д2

    arr = erp.NewObject("Массив")
    for ref in nomenklatura_refs:
        arr.Добавить(ref)

    rows = make_query(
        erp,
        text,
        {"ДатаС": dat_s, "ДатаПо": dat_po, "Номенклатуры": arr}
    )
    n = rows.Количество()
    print(f"  test_query_avg_prices: avg-prices rows={n} (из {len(nomenklatura_refs)} запрошенных номенклатур)")
    # Проверяем что цены > 0 для найденных
    bad = 0
    for i in range(n):
        s = rows.Получить(i)
        if float(s.СредняяЦена) <= 0:
            bad += 1
    assert bad == 0, f"{bad} строк со средней ценой <= 0"
    return rows


def test_statya_dohodov_exists(erp):
    """ПВХ.СтатьиДоходов.НайтиПоКоду('000001007') возвращает непустую ссылку."""
    ref = erp.ПланыВидовХарактеристик.СтатьиДоходов.НайтиПоКоду(STATYA_KOD)
    presentation = erp.String(ref)
    print(f"  test_statya_dohodov_exists: '{presentation}'")
    assert presentation, f"СтатьяДоходов с кодом {STATYA_KOD} не найдена"


def test_sklad_podrazdelenie(erp, rows_balances):
    """Все склады из остатков имеют непустое Подразделение."""
    sklady_uids = set()
    for i in range(rows_balances.Количество()):
        s = rows_balances.Получить(i)
        sklady_uids.add(erp.String(s.Склад.УникальныйИдентификатор()))

    bad = []
    for uid in sklady_uids:
        ref = erp.Справочники.Склады.ПолучитьСсылку(
            erp.NewObject("УникальныйИдентификатор", uid))
        obj = ref.ПолучитьОбъект()
        podr = obj.Подразделение
        if not erp.ЗначениеЗаполнено(podr):
            bad.append(erp.String(ref))
    print(f"  test_sklad_podrazdelenie: складов={len(sklady_uids)}, без Подразделения={len(bad)}")
    assert not bad, f"Склады без Подразделения: {bad}"


def main():
    print("=" * 60)
    print("TEST: запросы для ЗакрытиеОтрицательныхОстатков")
    print("=" * 60)
    erp = connect()

    rows = test_query_negative_balances(erp)

    nom_refs = []
    seen = set()
    for i in range(rows.Количество()):
        s = rows.Получить(i)
        uid = erp.String(s.Номенклатура.УникальныйИдентификатор())
        if uid not in seen:
            seen.add(uid)
            nom_refs.append(s.Номенклатура)

    test_query_avg_prices(erp, nom_refs)
    test_statya_dohodov_exists(erp)
    test_sklad_podrazdelenie(erp, rows)

    print("=" * 60)
    print("OK: 4/4 тестов прошли")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except AssertionError as e:
        print(f"FAIL: {e}")
        sys.exit(1)
    except Exception as e:
        if hasattr(e, "excepinfo") and e.excepinfo:
            print(f"FAIL: {e.excepinfo[2]}")
        else:
            print(f"FAIL: {e}")
        sys.exit(2)
