# -*- coding: utf-8 -*-
"""Проверка: поле Документ в наборе данных (для расшифровки) + отчёт компонуется с фактом."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ПлановаяСебестоимость.erf"

report = erp.ВнешниеОтчеты.Создать(ERF)
schema = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
ds = schema.НаборыДанных.Получить(0)
поля = [ds.Поля.Получить(j).Поле for j in range(ds.Поля.Количество())]
print("Поля набора данных:", поля)
assert "Документ" in поля, "Поле Документ НЕ в наборе данных!"
print(">>> Документ ЕСТЬ в наборе данных -> появится в окне 'Выбор поля' расшифровки. OK")

# Регрессия: компоновка с дефолтом (период пуст) -> факт виден
composer = erp.NewObject("КомпоновщикНастроекКомпоновкиДанных")
composer.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
composer.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
settings = composer.ПолучитьНастройки()
tcomposer = erp.NewObject("КомпоновщикМакетаКомпоновкиДанных")
template = tcomposer.Выполнить(schema, settings)
proc = erp.NewObject("ПроцессорКомпоновкиДанных")
proc.Инициализировать(template)
doc = erp.NewObject("ТабличныйДокумент")
out = erp.NewObject("ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент")
out.УстановитьДокумент(doc)
out.Вывести(proc)
# Проверим, что в строке подразделения есть факт (колонка ~7)
факт_есть = False
for r in range(1, min(doc.ВысотаТаблицы, 10) + 1):
    txt = (doc.Область(r, 7, r, 7).Текст or "").strip()
    if txt and any(ch.isdigit() for ch in txt):
        факт_есть = True
        break
print(f"Компоновка {doc.ВысотаТаблицы}x{doc.ШиринаТаблицы}; факт в колонке Сумма виден: {факт_есть}")
assert doc.ВысотаТаблицы > 1 and факт_есть
print("OK: поле Документ доступно для расшифровки, факт на месте.")
