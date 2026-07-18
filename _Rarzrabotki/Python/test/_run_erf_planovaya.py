# -*- coding: utf-8 -*-
"""Тест СОБРАННОГО .erf: загрузить внешний отчёт через COM, скомпоновать, показать данные."""
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

# 1) Загрузить собранный внешний отчёт
report = erp.ВнешниеОтчеты.Создать(ERF)
print("1) Внешний отчёт загружен из .erf:", report.Метаданные().Имя)

# 2) Схема из макета отчёта + настройки по умолчанию
schema = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
composer = erp.NewObject("КомпоновщикНастроекКомпоновкиДанных")
composer.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
settings = composer.ПолучитьНастройки()
print("2) Схема и настройки получены из .erf")

# 3) Период
period = erp.NewObject("СтандартныйПериод")
period.ДатаНачала = date1c(2024, 1, 1)
period.ДатаОкончания = date1c(2026, 12, 31, 23, 59, 59)
pv = settings.ПараметрыДанных.НайтиЗначениеПараметра(erp.NewObject("ПараметрКомпоновкиДанных", "Период"))
pv.Значение = period
pv.Использование = True
print("3) Период: 2024-01-01 .. 2026-12-31")

# 4) Компоновка в табличный документ
tcomposer = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
template = tcomposer.Выполнить(schema, settings)
proc = erp.NewObject("ПроцессорКомпоновкиДанных")
proc.Инициализировать(template)
doc = erp.NewObject("ТабличныйДокумент")
out = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
out.УстановитьДокумент(doc)
out.Вывести(proc)
print(f"4) Отчёт сформирован: высота={doc.ВысотаТаблицы}, ширина={doc.ШиринаТаблицы}")

# 5) Показать первые строки результата (реальное содержимое)
print("\n--- Первые строки табличного документа ---")
for r in range(1, min(doc.ВысотаТаблицы, 18) + 1):
    cells = []
    for c in range(1, min(doc.ШиринаТаблицы, 10) + 1):
        t = doc.Область(r, c, r, c).Текст
        if t:
            cells.append(t.strip()[:22])
    if cells:
        print(f"  стр{r}: " + " | ".join(cells))

assert doc.ВысотаТаблицы > 1, "Пустой результат!"
print("\nТЕСТ .ERF ПРОЙДЕН: собранный внешний отчёт загружается и формирует результат.")
