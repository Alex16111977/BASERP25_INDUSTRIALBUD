# -*- coding: utf-8 -*-
"""
RED-фаза: валидация запроса по статье "Аренда сторонней техники".

Проверяет, что запрос (тот же текст, что в ObjectModule.ПолучитьРасшифровкуПлатежей)
работает в реальной BSL-форме с параметром &СтатьяДДС типа CatalogRef.

Эталонные значения (получены через mcp execute_query):
  - 204 строки за период 01.12.2025 - 31.03.2026
  - Итого СуммаОплатыОборот = 10 033 642,31 грн
  - Топ-сумма = 1 092 595 на договоре "Рахунки_Глобино-2" / Європабуд ТОВ
"""

import sys
import datetime
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
STATIA_UUID = "3e2ba4c7-7a7e-11eb-a208-000c299fb278"   # "Аренда сторонней техники"

EXPECTED_ROWS  = 204
EXPECTED_TOTAL = 10_033_642.31
EXPECTED_TOP_AMOUNT = 1_092_595.0
EXPECTED_TOP_PARTNER = "Європабуд ТОВ"


def fail(msg):
    print(f"FAIL: {msg}")
    sys.exit(1)


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)
    print("OK: connected to BaseERP")

    uid = erp.NewObject("УникальныйИдентификатор", STATIA_UUID)
    statia_ref = erp.Справочники.СтатьиДвиженияДенежныхСредств.ПолучитьСсылку(uid)
    if statia_ref.ПолучитьОбъект() is None:
        fail("Статья 'Аренда сторонней техники' не найдена по UUID")
    print(f"OK: статья = '{erp.String(statia_ref)}'")

    nach = datetime.datetime(2025, 12, 1)
    okon = datetime.datetime(2026, 3, 31, 23, 59, 59)

    q = erp.NewObject("Запрос")
    q.SetParameter("НачалоПериода",    nach)
    q.SetParameter("ОкончаниеПериода", okon)
    q.SetParameter("СтатьяДДС",        statia_ref)
    q.Text = """
    ВЫБРАТЬ
        РегДДС.Регистратор    КАК Регистратор,
        РегДДС.Контрагент     КАК Контрагент,
        ВЫРАЗИТЬ(РегДДС.Договор КАК Справочник.ДоговорыКонтрагентов) КАК Договор,
        РегДДС.Подразделение  КАК Подразделение,
        РегДДС.СуммаОплатыОборот КАК Сумма
    ИЗ
        РегистрНакопления.ДвиженияДенежныеСредстваКонтрагент.Обороты(
            &НачалоПериода,
            &ОкончаниеПериода,
            Регистратор,
            СтатьяДвиженияДенежныхСредств = &СтатьяДДС
                И ТИПЗНАЧЕНИЯ(Договор) = ТИП(Справочник.ДоговорыКонтрагентов)
                И ВЫРАЗИТЬ(Договор КАК Справочник.ДоговорыКонтрагентов).ТипДоговора
                    = ЗНАЧЕНИЕ(Перечисление.ТипыДоговоров.СПоставщиком)
        ) КАК РегДДС
    ГДЕ
        РегДДС.СуммаОплатыОборот <> 0
    УПОРЯДОЧИТЬ ПО Сумма УБЫВ
    """

    try:
        result = q.Execute().Выгрузить()
    except Exception as e:
        if hasattr(e, "excepinfo") and e.excepinfo:
            fail(f"Запрос упал: {e.excepinfo[2]}")
        else:
            fail(f"Запрос упал: {e}")

    rows = result.Количество()
    print(f"OK: запрос вернул {rows} строк")

    if rows != EXPECTED_ROWS:
        fail(f"Ожидалось {EXPECTED_ROWS} строк, получено {rows}")

    total = 0.0
    top_amount = 0.0
    top_partner = ""
    for i in range(rows):
        row = result.Получить(i)
        total += float(row.Сумма)
        if i == 0:
            top_amount = float(row.Сумма)
            top_partner = str(erp.String(row.Контрагент))

    print(f"OK: итого = {total:,.2f}")
    print(f"OK: топ = {top_amount:,.2f} на {top_partner}")

    if abs(total - EXPECTED_TOTAL) > 0.5:
        fail(f"Сумма {total:.2f} не равна эталону {EXPECTED_TOTAL:.2f}")
    if abs(top_amount - EXPECTED_TOP_AMOUNT) > 0.5:
        fail(f"Топ-сумма {top_amount:.2f} != эталон {EXPECTED_TOP_AMOUNT:.2f}")
    if EXPECTED_TOP_PARTNER not in top_partner:
        fail(f"Топ-партнёр '{top_partner}' не содержит '{EXPECTED_TOP_PARTNER}'")

    print("\nSUCCESS: запрос валиден, эталоны совпали")


if __name__ == "__main__":
    main()
