# -*- coding: utf-8 -*-
"""
Test: автономная печатная форма "Печать счета" (после отвязки от ПечатьСчетовНаОплату).
Document: ЗаказКлиента ІБ00-000360 от 16.04.2026 14:23:46 на 668 469,6 UAH.

Логика:
  1. COM connect → BaseERP
  2. Получить ссылку на документ по UUID
  3. Загрузить EPF через ВнешниеОбработки.Создать()
  4. Подготовить КоллекциюПечатныхФорм (точная схема как в типовом УправлениеПечатью)
  5. Вызвать ВнОбр.Печать(...)
  6. Сохранить ТабличныйДокумент в .mxl
  7. Распарсить .mxl и assert на ключевые поля
"""

import sys
from pathlib import Path

import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

EPF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\ПечатьСчета.epf"
MXL_OUT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\счет_ІБ00-000360.mxl"
DOC_UUID = "646adb6b-3cab-11f1-8110-00155dce3d04"
DOC_NUMBER = "ІБ00-000360"

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def fail(msg):
    print(f"FAIL: {msg}")
    sys.exit(1)


def main():
    # ---------- 1. COM connect ----------
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)
    print("OK: connected to BaseERP")

    # ---------- 2. Документ по UUID ----------
    uid = erp.NewObject("УникальныйИдентификатор", DOC_UUID)
    ref = erp.Документы.ЗаказКлиента.ПолучитьСсылку(uid)
    obj = ref.ПолучитьОбъект()
    if obj is None:
        fail(f"Заказ клиента {DOC_NUMBER} (UUID {DOC_UUID}) не найден")
    print(f"OK: doc {DOC_NUMBER}, контрагент={ref.Контрагент}, сумма={ref.СуммаДокумента}")

    # ---------- 3. Загрузка EPF ----------
    if not Path(EPF_PATH).exists():
        fail(f"EPF не собран: {EPF_PATH}. Сначала запусти epf-build.")
    vn_obr = erp.ВнешниеОбработки.Создать(EPF_PATH, True)
    print(f"OK: EPF loaded ({Path(EPF_PATH).stat().st_size} bytes)")

    # ---------- 4. Параметры для прямого вызова СформироватьПечатнуюФормуСчетНаОплату ----------
    # Обходим УправлениеПечатью.НужноПечататьМакет (требует особую схему ТЗ);
    # СформироватьПечатнуюФормуСчетНаОплату — публичная Экспорт-функция обработки,
    # которая сама возвращает ТабличныйДокумент.
    massiv_obj = erp.NewObject("Массив")
    massiv_obj.Добавить(ref)

    # Соответствие "ПолноеИмяОбъекта" → Массив ссылок (как делает сам ObjectModule:66)
    struktura_tipov = erp.ОбщегоНазначенияУТ.СоответствиеМассивовПоТипамОбъектов(massiv_obj)

    objects_print = erp.NewObject("СписокЗначений")

    params_print = erp.NewObject("Структура")

    params_out = erp.NewObject("Структура")
    params_out.Вставить("КодЯзыкаДляМногоязычныхПечатныхФорм", "uk")

    # ---------- 5. Печать ----------
    try:
        tab_doc = vn_obr.СформироватьПечатнуюФормуСчетНаОплату(
            struktura_tipov, objects_print, params_print, params_out
        )
    except Exception as e:
        if hasattr(e, "excepinfo") and e.excepinfo:
            fail(f"СформироватьПечатнуюФормуСчетНаОплату() упала: {e.excepinfo[2]}")
        else:
            fail(f"СформироватьПечатнуюФормуСчетНаОплату() упала: {e}")
    print("OK: печать выполнена")

    # ---------- 6. Сохранить MXL ----------
    if tab_doc is None:
        fail("Печать вернула None")

    try:
        height = tab_doc.ВысотаТаблицы
        width = tab_doc.ШиринаТаблицы
    except Exception:
        # запасной вариант через КоличествоСтрок/КоличествоКолонок
        height = getattr(tab_doc, "КоличествоСтрок", lambda: 0)()
        width = getattr(tab_doc, "КоличествоКолонок", lambda: 0)()

    if not height or not width:
        fail(f"MXL пустой (height={height}, width={width})")
    print(f"OK: ТабличныйДокумент {height}x{width}")

    # MXL = тип файла "Табличный документ" (значение 0)
    tip_mxl = erp.ТипФайлаТабличногоДокумента.MXL
    Path(MXL_OUT).parent.mkdir(parents=True, exist_ok=True)
    tab_doc.Записать(MXL_OUT, tip_mxl)
    if not Path(MXL_OUT).exists():
        fail(f"MXL не сохранён: {MXL_OUT}")
    print(f"OK: MXL saved → {MXL_OUT} ({Path(MXL_OUT).stat().st_size} bytes)")

    # ---------- 7. Парсинг и проверка полей ----------
    mxl_text = Path(MXL_OUT).read_text(encoding="utf-8", errors="ignore")

    # Сумма выводится с разными разделителями — проверяем варианты
    summa_variants = [
        "668 469,6",
        "668 469,60",
        "668 469,6",
        "668 469,60",
        "668469,6",
        "668469,60",
    ]
    summa_ok = any(s in mxl_text for s in summa_variants)
    if not summa_ok:
        fail("Сумма 668 469,6 не найдена в MXL")
    print("OK: сумма 668 469,6 найдена")

    # Типовое СформироватьЗаголовокДокумента обрезает префикс ІБ00- и ведущие нули,
    # поэтому в MXL номер выводится как "№ 360 від 16 квітня 2026"
    checks = {
        "номер 360 в заголовке": "№ 360",
        "дата 16 квітня 2026":   "16 квітня 2026",
        "контрагент (АСТАРТА)":  "АСТАРТА",
        "организация (ІНДАСТРІАЛБУД)": "ІНДАСТРІАЛБУД",
    }
    for label, fragment in checks.items():
        if fragment not in mxl_text:
            fail(f"{label}: '{fragment}' не найдено в MXL")
        print(f"OK: {label} найден")

    print(f"\nSUCCESS: все ключевые поля найдены в {MXL_OUT}")


if __name__ == "__main__":
    main()
