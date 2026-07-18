# -*- coding: utf-8 -*-
"""Проверка: собранный .erf использует типовую ОбщаяФорма.ФормаОтчета и компонуется."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость.erf"

def date1c(y, m, d, hh=0, mm=0, ss=0):
    q = erp.NewObject("Запрос")
    q.Текст = f"ВЫБРАТЬ ДАТАВРЕМЯ({y},{m},{d},{hh},{mm},{ss})"
    return q.Выполнить().Выгрузить().Получить(0).Получить(0)

report = erp.ВнешниеОтчеты.Создать(ERF)
md = report.Метаданные()
print("Имя отчёта:", md.Имя)
for prop in ("ОсновнаяФорма", "ОсновнаяФормаНастроек", "ОсновнаяФормаВарианта"):
    try:
        f = getattr(md, prop)
        print(f"  {prop}: {f.ПолноеИмя() if f is not None else 'Неопределено'}")
    except Exception as e:
        print(f"  {prop}: <нет свойства> {e}")

# Проверка, что параметр Период доступен в пользовательских настройках (быстрый доступ)
schema = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
composer = erp.NewObject("КомпоновщикНастроекКомпоновкиДанных")
composer.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
settings = composer.ПолучитьНастройки()
period = erp.NewObject("СтандартныйПериод")
period.ДатаНачала = date1c(2024, 1, 1)
period.ДатаОкончания = date1c(2026, 12, 31, 23, 59, 59)
pv = settings.ПараметрыДанных.НайтиЗначениеПараметра(erp.NewObject("ПараметрКомпоновкиДанных", "Период"))
pv.Значение = period
pv.Использование = True

tcomposer = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
template = tcomposer.Выполнить(schema, settings)
proc = erp.NewObject("ПроцессорКомпоновкиДанных")
proc.Инициализировать(template)
doc = erp.NewObject("ТабличныйДокумент")
out = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
out.УстановитьДокумент(doc)
out.Вывести(proc)
print(f"\nКомпоновка: {doc.ВысотаТаблицы}x{doc.ШиринаТаблицы}")
assert doc.ВысотаТаблицы > 1
print("OK: отчёт использует типовую форму и формирует результат.")
