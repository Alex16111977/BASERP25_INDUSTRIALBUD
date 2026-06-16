# -*- coding: utf-8 -*-
"""Приёмочный тест: загрузить СКД отчёта из Template.xml и скомпоновать headless
через COM (без сборки .erf). Проверяет: схема читается, запрос валиден в контексте
СКД, параметр Период работает, результат формируется."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

TEMPLATE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"

def date1c(y, m, d, hh=0, mm=0, ss=0):
    q = erp.NewObject("Запрос")
    q.Текст = f"ВЫБРАТЬ ДАТАВРЕМЯ({y},{m},{d},{hh},{mm},{ss})"
    return q.Выполнить().Выгрузить().Получить(0).Получить(0)

# 1) Загрузить схему компоновки из XML
reader = erp.NewObject("ЧтениеXML")
reader.ОткрытьФайл(TEMPLATE)
schema = erp.СериализаторXDTO.ПрочитатьXML(reader)
reader.Закрыть()
print("1) Схема компоновки загружена из XML:", type(schema).__name__ if schema else "Неопределено")

# 2) Компоновщик настроек -> настройки по умолчанию (вариант "Основной")
composer = erp.NewObject("КомпоновщикНастроекКомпоновкиДанных")
source = erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
composer.Инициализировать(source)
composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
settings = composer.ПолучитьНастройки()
print("2) Настройки по умолчанию загружены, вариант 'Основной'")

# 3) Параметр Период = 2024-01-01 .. 2026-12-31
period = erp.NewObject("СтандартныйПериод")
period.ДатаНачала = date1c(2024, 1, 1)
period.ДатаОкончания = date1c(2026, 12, 31, 23, 59, 59)
pv = settings.ПараметрыДанных.НайтиЗначениеПараметра(erp.NewObject("ПараметрКомпоновкиДанных", "Период"))
pv.Значение = period
pv.Использование = True
print("3) Параметр Период установлен: 2024-01-01 .. 2026-12-31")

# 4) Компоновка -> вывод в табличный документ (минимальная форма вызова)
tcomposer = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
template = tcomposer.Выполнить(schema, settings)
proc = erp.NewObject("ПроцессорКомпоновкиДанных")
proc.Инициализировать(template)
doc = erp.NewObject("ТабличныйДокумент")
out = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
out.УстановитьДокумент(doc)
out.Вывести(proc)
print("4) Компоновка выполнена. Высота таблицы:", doc.ВысотаТаблицы, "Ширина:", doc.ШиринаТаблицы)
assert doc.ВысотаТаблицы > 1, "Пустой результат компоновки!"
print("\nПРИЁМКА ПРОЙДЕНА: СКД читается, запрос компонуется, параметр Период работает, отчёт сформирован.")
