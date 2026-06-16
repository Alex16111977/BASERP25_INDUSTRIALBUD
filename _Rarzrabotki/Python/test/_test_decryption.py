# -*- coding: utf-8 -*-
"""Генерирует ли СКД данные расшифровки? Если да — проблема в интеграции с формой БСП."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость.erf"

report = erp.ВнешниеОтчеты.Создать(ERF)
schema = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
composer = erp.NewObject("КомпоновщикНастроекКомпоновкиДанных")
composer.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
settings = composer.ПолучитьНастройки()

decrypt = erp.NewObject("ДанныеРасшифровкиКомпоновкиДанных")
tcomposer = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
template = tcomposer.Выполнить(schema, settings, decrypt)
proc = erp.NewObject("ПроцессорКомпоновкиДанных")
proc.Инициализировать(template, erp.NewObject("Структура"), decrypt, True)
doc = erp.NewObject("ТабличныйДокумент")
out = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
out.УстановитьДокумент(doc)
out.Вывести(proc)

n = decrypt.Элементы.Количество()
print(f"Высота={doc.ВысотаТаблицы}; элементов расшифровки={n}")
if n > 0:
    print(">>> СКД ГЕНЕРИРУЕТ расшифровку. Значит проблема — в интеграции с формой БСП (нужны хуки модуля).")
else:
    print(">>> СКД НЕ генерирует расшифровку. Проблема в самой СКД.")
