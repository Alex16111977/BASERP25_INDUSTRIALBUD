# -*- coding: utf-8 -*-
"""Проверка: .erf с ДЕФОЛТНЫМИ настройками (период пуст) показывает ФАКТ."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость.erf"

report = erp.ВнешниеОтчеты.Создать(ERF)
schema = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
composer = erp.NewObject("КомпоновщикНастроекКомпоновкиДанных")
composer.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
settings = composer.ПолучитьНастройки()   # период по умолчанию (пуст) — НЕ трогаем

tcomposer = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
template = tcomposer.Выполнить(schema, settings)
proc = erp.NewObject("ПроцессорКомпоновкиДанных")
proc.Инициализировать(template)
doc = erp.NewObject("ТабличныйДокумент")
out = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
out.УстановитьДокумент(doc)
out.Вывести(proc)
print(f"Высота={doc.ВысотаТаблицы} Ширина={doc.ШиринаТаблицы} (период по умолчанию = пуст)")
print("\n--- Первые строки (колонки: Подр/КолСС/Кол/РазнКол/%/СуммаСС/Сумма/РазнСум/%) ---")
for r in range(1, min(doc.ВысотаТаблицы, 16) + 1):
    cells = []
    for c in range(1, min(doc.ШиринаТаблицы, 10) + 1):
        t = doc.Область(r, c, r, c).Текст
        cells.append((t or "").strip()[:20])
    print(f"  стр{r}: " + " | ".join(cells))
print("\nЕсли в колонках 'Количество' и 'Сумма' (3-я и 7-я) есть числа — ФАКТ виден.")
