# -*- coding: utf-8 -*-
"""Smoke-тест внешнего отчёта «Отчет по общим названиям в СС» (BuhBud).

1) Прямой запрос (1:1 из СКД) -> эталон: число строк, СуммаСумма.
2) ВнешниеОтчеты.Создать(.erf) -> ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных").
3) Программная компоновка с настройками по умолчанию -> ТаблицаЗначений.
4) Детальные строки ТЗ (Структура заполнена, группировочные/итоговые -> Неопределено)
   сверяются с эталоном: число строк == эталон, |СуммаСумма - эталон| < 0.05.
"""
import sys
import pythoncom
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Загрузка СС\Отчет по общим названиям в СС.erf"

QUERY = """
ВЫБРАТЬ
	К.Ссылка КАК Структура,
	К.Ссылка.Активная КАК Активная,
	К.Этап КАК Этап,
	К.ОбщееНазвание КАК ОбщееНазвание,
	К.Номенклатура КАК Номенклатура,
	К.НоменклатураСС КАК НоменклатураСС,
	К.Единица КАК Единица,
	К.ЕдиницаСС КАК ЕдиницаСС,
	К.Количество КАК Количество,
	К.Сумма КАК Сумма,
	ВЫБОР КОГДА К.Количество <> 0 ТОГДА К.Сумма / К.Количество ИНАЧЕ 0 КОНЕЦ КАК Цена
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК К
"""


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

    # --- 1. Эталон: прямой запрос ---
    q = buh.NewObject("Запрос")
    q.Text = QUERY
    et = q.Execute().Выгрузить()
    et_rows = et.Количество()
    et_summa = 0.0
    for i in range(et_rows):
        s = et.Получить(i).Сумма
        if s is not None:
            et_summa += float(s)
    print(f"[Эталон] строк={et_rows}, СуммаИтого={et_summa:.2f}")

    # --- 2. Отчёт из .erf ---
    rep = buh.ВнешниеОтчеты.Создать(ERF, False)
    schema = rep.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
    settings = schema.НастройкиПоУмолчанию

    # --- 3. Программная компоновка -> ТаблицаЗначений ---
    # COM-грабли: Тип(...) недоступен через V83 -> трюк через ОписаниеТипов
    generator_type = buh.NewObject(
        "ОписаниеТипов",
        "ГенераторМакетаКомпоновкиДанныхДляКоллекцииЗначений"
    ).Типы().Получить(0)
    # COM-грабли: пропуск параметров = pythoncom.Empty (VT_EMPTY = Неопределено),
    # None маршалится как VT_NULL и даёт "Несоответствие типов"
    composer = buh.NewObject("КомпоновщикМакетаКомпоновкиДанных")
    layout = composer.Выполнить(schema, settings, pythoncom.Empty,
                                pythoncom.Empty, generator_type)
    processor = buh.NewObject("ПроцессорКомпоновкиДанных")
    processor.Инициализировать(layout)
    vt = buh.NewObject("ТаблицаЗначений")
    output = buh.NewObject(
        "ПроцессорВыводаРезультатаКомпоновкиДанныхВКоллекциюЗначений")
    output.УстановитьОбъект(vt)
    output.Вывести(processor)
    total_rows = vt.Количество()

    # --- 4. Детальные строки: Структура заполнена (в группировочных строках
    #        не-группировочные поля = Неопределено -> None в Python) ---
    det_rows = 0
    det_summa = 0.0
    for i in range(total_rows):
        row = vt.Получить(i)
        if row.Структура is None:
            continue  # группировка ОбщееНазвание/Номенклатура или общий итог
        det_rows += 1
        s = row.Сумма
        if s is not None:
            det_summa += float(s)
    print(f"[СКД] строк всего={total_rows}, детальных={det_rows}, "
          f"СуммаИтого(детальные)={det_summa:.2f}")

    # --- Ассерты ---
    assert det_rows == et_rows, (
        f"FAIL: детальных строк {det_rows} != эталон {et_rows}")
    diff = abs(det_summa - et_summa)
    assert diff < 0.05, (
        f"FAIL: |СуммаСКД - СуммаЭталон| = {diff:.4f} >= 0.05")
    print(f"OK: детальных строк {det_rows} == эталон, "
          f"расхождение суммы {diff:.4f} < 0.05")


if __name__ == "__main__":
    main()
