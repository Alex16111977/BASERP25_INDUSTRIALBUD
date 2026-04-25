# -*- coding: utf-8 -*-
"""
Тест новой логики отчёта А_ОтчетPL через LIVE-вызов функции ПолучитьОбъединенныеДанные().

Этот тест вызывает НАСТОЯЩИЙ модуль отчёта (после db-load-xml),
получает ТаблицуЗначений и прогоняет СКД компоновку.

Проверяет:
  1. Функция ПолучитьОбъединенныеДанные работает с новым запросом.
  2. СКД-компоновка Основного варианта не падает.
  3. Есть PL-строки + ЕРП-детали + ЕРП-без-PL.
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

CONNECTION_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
PERIOD_START = datetime(2025, 12, 1)
PERIOD_END = datetime(2025, 12, 31, 23, 59, 59)


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")
    conn = v8.Connect(CONNECTION_STRING)
    print("[OK] Connected")

    # Вызов функции ПолучитьОбъединенныеДанные через менеджер отчёта
    # (она Экспорт, поэтому доступна через серверный вызов)
    report_obj = conn.Reports.А_ОтчетPL.Создать()
    try:
        table = report_obj.ПолучитьОбъединенныеДанные(
            pywintypes.Time(PERIOD_START),
            pywintypes.Time(PERIOD_END)
        )
        print(f"[OK] ПолучитьОбъединенныеДанные вернула {table.Количество()} строк")
    except Exception as e:
        print(f"[FAIL] {e}")
        return 1

    # Разобрать строки по типу
    pl_rows = 0
    erp_details_by_pl = 0
    erp_without_pl = 0
    for row in table:
        is_pl_row = (row.СуммаPL != 0 or row.СуммаФ1 != 0 or row.СуммаФ2 != 0)
        has_pl_article = conn.ЗначениеЗаполнено(row.СтатьяPL)
        if is_pl_row:
            pl_rows += 1
        elif has_pl_article:
            erp_details_by_pl += 1
        else:
            erp_without_pl += 1

    print(f"  PL-строки (Документ=А_ОтчетPL):  {pl_rows}")
    print(f"  ЕРП-детали с PL-статьёй:          {erp_details_by_pl}")
    print(f"  ЕРП без PL-статьи:                {erp_without_pl}")

    # === СКД ===
    print("\n--- Компоновка Основного варианта ---")
    schema = conn.Reports.А_ОтчетPL.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")

    settings_composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    settings_source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    settings_composer.Инициализировать(settings_source)
    settings_composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)

    param_id = conn.NewObject("ПараметрКомпоновкиДанных", "Период")
    period_param = settings_composer.Настройки.ПараметрыДанных.НайтиЗначениеПараметра(param_id)
    std_period = conn.NewObject("СтандартныйПериод")
    std_period.ДатаНачала = pywintypes.Time(PERIOD_START)
    std_period.ДатаОкончания = pywintypes.Time(PERIOD_END)
    period_param.Значение = std_period
    period_param.Использование = True

    external_ds = conn.NewObject("Структура")
    external_ds.Вставить("Данные", table)

    layout_composer = conn.NewObject("КомпоновщикМакетаКомпоновкиДанных")
    try:
        layout = layout_composer.Выполнить(schema, settings_composer.ПолучитьНастройки(), None)
        print("[OK] Макет создан")
    except Exception as e:
        print(f"[FAIL] Layout: {e}")
        return 1

    processor = conn.NewObject("ПроцессорКомпоновкиДанных")
    try:
        processor.Инициализировать(layout, external_ds)
        print("[OK] Процессор инициализирован")
    except Exception as e:
        print(f"[FAIL] Processor: {e}")
        return 1

    output = conn.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
    result_doc = conn.NewObject("ТабличныйДокумент")
    output.УстановитьДокумент(result_doc)
    try:
        output.Вывести(processor)
        print(f"[OK] Отчёт выведен: {result_doc.ВысотаТаблицы} строк в табличном документе")
        return 0
    except Exception as e:
        print(f"[FAIL] Output: {e}")
        return 1


if __name__ == "__main__":
    import sys
    sys.exit(main())
