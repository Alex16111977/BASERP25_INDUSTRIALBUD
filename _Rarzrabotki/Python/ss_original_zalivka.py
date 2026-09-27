# -*- coding: utf-8 -*-
"""Заливка оригинала Excel в КомплектующиеПоМодульно + присоединение файлов-оснований к СС (BuhBud), 2026-09-27.

Спека: docs/superpowers/specs/2026-09-27-ss-original-koeficienty-fajly-design.md, п. 5.1.

Логика — ТОЛЬКО живой код базы:
  лист «Проект СС» читается через Excel COM так же, как кнопка формы СС (.Text для строк, .Value для чисел);
  отбор, разрешение ссылок и выравнивание общих названий по Комплектующим —
  Справочник.СтруктураСебестоимости.ЗаполнитьОригиналИзСырыхСтрок(Сырые, Модуль);
  файл — А_СтруктураСебестоимостиСервер.ПрисоединитьФайл + УстановитьФайлОснования.

Режимы на СС:
  ORIGINAL — КомплектующиеПоМодульно пусты: залить оригинал модуля + присоединить файл;
  ATTACH   — оригинал уже есть (МД ООН Админ): только присоединить файл; итоги строки ФайлыОснования —
             из существующих строк модуля; перед этим живым движком проверяется, что файл даёт те же строки/сумму.

Сухой прогон НЕ оставляет следов: каждая СС обрабатывается во внешней транзакции, которая без --write
откатывается целиком (движок загрузки фиксирует свою вложенную транзакцию с get-or-create этапов и общих
названий — внешняя отмена откатывает и её).

Инварианты (при нарушении транзакция СС откатывается):
  - Комплектующие не меняются: строки / Этап / ОбщееНазвание / НоменклатураСС / ЕдиницаСС / Количество / Сумма;
  - итоги залитого оригинала модуля = итоги файла по протоколу движка (контроль движка пройден);
  - после записи: строки ФайлыОснования есть, двоичные данные присоединённых файлов читаются.

Запуск:  python ss_original_zalivka.py            (сухой прогон)
         python ss_original_zalivka.py --write    (запись)
"""
import os
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import pythoncom
import win32com.client

CONN = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'
ROOT = r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Загрузка СС'
TARGETS = [
    # (код СС, режим, [(модуль, файл)])
    ('000000008', 'ORIGINAL', [(0, os.path.join(ROOT, 'IRC', 'Погоджений',
                                               'IRC 15м2 НОВИЙ ШАблон  СС Виробництво 16-06-2026_Коррект (1).xlsx'))]),
    ('000000009', 'ORIGINAL', [(0, os.path.join(ROOT, 'IRC', 'Погоджений',
                                               'IRC 30м2 НОВИЙ ШАблон  СС Виробництво 16-06-2026_Коррект (2).xlsx'))]),
    ('000000010', 'ORIGINAL', [(0, os.path.join(ROOT, 'МД ООН 2026',
                                               '14.07.2026 РАБОЧИЙ  ООН_Б1_НОВИЙ_ШАблон_СС_Виробництво_15_06_2026_ (002).xlsx'))]),
    ('000000011', 'ATTACH', [(1, os.path.join(ROOT, 'СС ООН АДМИН', 'уніф СВ S1 Новий Шаблон 21.08.xlsx')),
                             (3, os.path.join(ROOT, 'СС ООН АДМИН', 'уніф СВ S3 Новий Шаблон 21.08.xlsx'))]),
]
SHEET = 'Проект СС'
WRITE = '--write' in sys.argv
TOL = 0.005
KEY_FIELDS = ('Этап', 'ОбщееНазвание', 'НоменклатураСС', 'ЕдиницаСС', 'Количество', 'Сумма')


def num(v):
    if v is None or v == '':
        return 0
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).strip())
    except ValueError:
        return 0


def read_sheet(excel, path):
    """Сырые строки листа «Проект СС» — как ПрочитатьЛистExcel формы СС."""
    book = excel.Workbooks.Open(path, 0, True)
    try:
        sheet = None
        for i in range(1, book.Sheets.Count + 1):
            if str(book.Sheets(i).Name).strip().lower() == SHEET.lower():
                sheet = book.Sheets(i)
        assert sheet is not None, f'нет листа «{SHEET}» в {path}'
        last = sheet.Cells(1, 1).SpecialCells(11).Row
        rows = []
        for r in range(2, last + 1):
            def text(col):
                return str(sheet.Cells(r, col).Text).strip()
            rows.append({'Ряд': r, 'Тип': text(3), 'Этап': text(5), 'Наименование': text(7),
                         'ОбщееНазвание': text(8), 'Единица': text(9),
                         'Количество': num(sheet.Cells(r, 10).Value), 'Сумма': num(sheet.Cells(r, 15).Value)})
        return rows
    finally:
        book.Close(False)


def to_1c(b, rows):
    arr = b.NewObject('Массив')
    for row in rows:
        s = b.NewObject('Структура')
        for k, v in row.items():
            s.Вставить(k, v)
        arr.Добавить(s)
    return arr


def snapshot(b, tab):
    """Построчный снимок ТЧ Комплектующие (ссылки — строкой УИД)."""
    res = []
    for i in range(tab.Количество()):
        row = tab.Получить(i)
        vals = []
        for f in KEY_FIELDS:
            v = getattr(row, f)
            if f in ('Этап', 'ОбщееНазвание'):
                v = b.String(v.УникальныйИдентификатор())
            elif f in ('Количество', 'Сумма'):
                v = round(float(v), 3)
            vals.append(v)
        res.append(tuple(vals))
    return res


def module_totals(tab, module):
    rows, total = 0, 0.0
    for i in range(tab.Количество()):
        row = tab.Получить(i)
        if module == 0 or row.Модуль == module:
            rows += 1
            total += float(row.Сумма)
    return rows, round(total, 2)


def print_protocol(p):
    print(f'    загружено {p.Загружено}, пропущено {p.Пропущено}, подхвачено без типа {p.ПодхваченоБезТипа}, '
          f'создано этапов {p.СозданоЭтапов}, названий {p.СозданоНазваний}, '
          f'дозаполнено единиц {p.ДозаполненоЕдиниц}, контроль {p.КонтрольПройден}')
    for coll in ('Контроль', 'Предупреждения', 'Инфо'):
        items = getattr(p, coll)
        for i in range(items.Количество()):
            print('     ', coll[0], items.Получить(i))


def process(b, srv, ref, obj, code, mode, raws, before, fails):
    """Заливка/проверка внутри открытой внешней транзакции. Истина — можно фиксировать."""
    n_fails = len(fails)
    attach = []  # (модуль, путь, строк, сумма)
    for module, path, raw in raws:
        print(f'  файл модуля {module}: {os.path.basename(path)}')
        if mode == 'ORIGINAL':
            p = obj.ЗаполнитьОригиналИзСырыхСтрок(raw, module)
            print_protocol(p)
            if p.Ошибка or not p.КонтрольПройден:
                fails.append(f'{code}: движок: {p.ТекстОшибки}')
                continue
            rows, total = module_totals(obj.КомплектующиеПоМодульно, module)
            if rows != p.Итоги.Строк or abs(total - float(p.Итоги.Сумма)) > TOL:
                fails.append(f'{code}: итоги ТЧ {rows}/{total} != файл {p.Итоги.Строк}/{p.Итоги.Сумма}')
            attach.append((module, path, rows, total))
        else:
            # Файл прогоняется живым движком на непишущейся копии: те же строки и сумма, что уже лежат в оригинале.
            probe = ref.ПолучитьОбъект()
            p = probe.ЗаполнитьОригиналИзСырыхСтрок(raw, module)
            rows_file, total_file = module_totals(probe.КомплектующиеПоМодульно, module)
            rows_db, total_db = module_totals(obj.КомплектующиеПоМодульно, module)
            print(f'    файл {rows_file}/{total_file}, в оригинале {rows_db}/{total_db}')
            if p.Ошибка or rows_file != rows_db or abs(total_file - total_db) > 0.02:
                fails.append(f'{code}: файл модуля {module} не совпадает с оригиналом в базе')
            attach.append((module, path, rows_db, total_db))

    srv.ОтметитьДобавленныеВручную(obj)
    if snapshot(b, obj.Комплектующие) != before:
        fails.append(f'{code}: Комплектующие изменились в памяти')
    orig = obj.КомплектующиеПоМодульно
    mism = sum(1 for i in range(orig.Количество()) if orig.Получить(i).НеСовпадает)
    manual = [obj.Комплектующие.Получить(i).НоменклатураСС for i in range(obj.Комплектующие.Количество())
              if obj.Комплектующие.Получить(i).ДобавленоВручную]
    print(f'  оригинал: {orig.Количество()} строк, НеСовпадает {mism}; «Додано вручну» в Комплектующих: {len(manual)}')
    for m in manual:
        print('     вручну:', m)
    if len(fails) > n_fails:
        return False
    if not WRITE:
        return True
    for module, path, rows, total in attach:
        data = b.NewObject('ДвоичныеДанные', path)
        file_ref = srv.ПрисоединитьФайл(ref, data, os.path.basename(path))
        srv.УстановитьФайлОснования(obj, module, file_ref, rows, total)
    obj.Записать()
    return True


def verify(b, ref, code, before, fails):
    fresh = ref.ПолучитьОбъект()
    if snapshot(b, fresh.Комплектующие) != before:
        fails.append(f'{code}: Комплектующие изменились после записи')
    for i in range(fresh.ФайлыОснования.Количество()):
        row = fresh.ФайлыОснования.Получить(i)
        size = b.РаботаСФайлами.ДвоичныеДанныеФайла(row.Файл).Размер()
        print(f'  ЗАПИСАНО: модуль {row.Модуль}, файл «{row.Файл}» {size} байт, '
              f'строк {row.КоличествоСтрок}, Σ {row.СуммаИтого}')
        if size <= 0:
            fails.append(f'{code}: пустой файл модуля {row.Модуль}')


def main():
    pythoncom.CoInitialize()
    b = win32com.client.Dispatch('V83.COMConnector').Connect(CONN)
    srv = b.А_СтруктураСебестоимостиСервер
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    fails = []
    try:
        for code, mode, files in TARGETS:
            ref = b.Справочники.СтруктураСебестоимости.НайтиПоКоду(code)
            assert not ref.Пустая(), code
            obj = ref.ПолучитьОбъект()
            print(f'\n=== {code} {obj.Наименование} [{mode}]')
            before = snapshot(b, obj.Комплектующие)
            print(f'  Комплектующие: {len(before)} строк, Σ Сумма {round(sum(r[5] for r in before), 2)}')
            if mode == 'ORIGINAL' and obj.КомплектующиеПоМодульно.Количество() > 0:
                print('  ПРОПУСК: КомплектующиеПоМодульно уже заполнены')
                continue

            raws = []
            for module, path in files:
                assert os.path.exists(path), path
                raws.append((module, path, to_1c(b, read_sheet(excel, path))))

            b.НачатьТранзакцию()
            try:
                ok = process(b, srv, ref, obj, code, mode, raws, before, fails)
            except Exception:
                b.ОтменитьТранзакцию()
                raise
            if WRITE and ok:
                b.ЗафиксироватьТранзакцию()
                verify(b, ref, code, before, fails)
            else:
                b.ОтменитьТранзакцию()
    finally:
        excel.Quit()

    if fails:
        print('\nПРОВАЛЫ:\n  ' + '\n  '.join(fails))
    else:
        print('\nЗАПИСАНО, инварианты соблюдены' if WRITE else '\nСУХОЙ ПРОГОН ОК, база не изменена (запись: --write)')
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())
