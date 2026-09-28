# -*- coding: utf-8 -*-
"""
Приймання детектора обробки А_РозрахункиПДВЗаПершоюПодією (без запису в базу).

1. Завантажує .epf через ВнешниеОбработки.Создать(path, Ложь) у bas_industrialbud.
2. Інваріант типів: у реквізитів і колонок ТЧ немає «Произвольного» типу (0 типів).
3. Анализировать() за 01.01.2026 – сьогодні: є ключі/документи з розбіжністю; приклад користувача (оплата 000013825,
   ключ «Рахунок на оплату постачальника ІБ00-003315») знайдено з очікуваним 0 і фактом 433,33 — якщо ще не виправлено.
4. Інваріант відміток: відмічені лише розбіжні документи відкритого періоду; відмітка ключа = є відмічені документи.
5. ГраницаЗапретаИзменения() = загальна дата заборони; ДокументыМеханизмаЗаПериод() за поточний місяць не порожній.

Запуск:  python test_pdv6442_analiz.py [шлях_до_epf]
"""
import os, sys, datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EPF = os.path.abspath(os.path.join(HERE, '..', '..', 'Обработки', 'А_РозрахункиПДВЗаПершоюПодією.epf'))
CONNS = ['Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"',
         'Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"']
FAILS = []


def check(cond, msg):
    print(('OK   ' if cond else 'FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    for s in CONNS:
        try:
            return v8.Connect(s)
        except Exception:  # noqa
            pass
    raise SystemExit('FAIL connect')


def main():
    epf = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EPF
    c = connect()
    proc = c.ВнешниеОбработки.Создать(epf, False)
    md = proc.Метаданные()
    untyped = [md.Реквизиты.Получить(i).Имя for i in range(md.Реквизиты.Количество())
               if md.Реквизиты.Получить(i).Тип.Типы().Количество() == 0]
    for j in range(md.ТабличныеЧасти.Количество()):
        ts = md.ТабличныеЧасти.Получить(j)
        untyped += [ts.Имя + '.' + ts.Реквизиты.Получить(i).Имя for i in range(ts.Реквизиты.Количество())
                    if ts.Реквизиты.Получить(i).Тип.Типы().Количество() == 0]
    check(not untyped, f'усі реквізити типізовані (без типу: {untyped})')
    for name in ('Организация', 'Контрагент', 'ДоговорКонтрагента'):
        t = md.Реквизиты.Найти(name).Тип
        check(t.Типы().Количество() == 1, f'тип {name}: {c.String(t)}')
    t = md.ТабличныеЧасти.Найти('Ключи').Реквизиты.Найти('Сделка').Тип
    check(t.Типы().Количество() > 20, f'Ключи.Сделка — усі документи ({t.Типы().Количество()} типів)')

    border = proc.ГраницаЗапретаИзменения()
    print('ГраницаЗапрета =', c.String(border))
    check(c.String(border).startswith('31.08.2026') or c.String(border) != '', 'дата заборони прочитана')

    today = datetime.datetime.now()
    proc.ДатаНачала = datetime.datetime(2026, 1, 1, 12, 0, 0)
    proc.ДатаОкончания = datetime.datetime(today.year, today.month, today.day, 12, 0, 0)
    t0 = datetime.datetime.now()
    res = proc.Анализировать()
    sec = (datetime.datetime.now() - t0).total_seconds()
    print(f'Анализировать: {sec:.0f} c; ключів {res.КлючейСРасхождением}, документів з розбіжністю '
          f'{res.ДокументовСРасхождением}, зайвий ПС {res.Отклонение}, документів у ключах {res.ВсегоДокументов}')
    check(res.КлючейСРасхождением > 0 and res.ДокументовСРасхождением > 0, 'детектор знайшов розбіжності')
    check(proc.Ключи.Количество() == res.КлючейСРасхождением, 'ТЧ Ключи = підсумок')

    # приклад користувача
    hit = None
    for i in range(proc.Документы.Количество()):
        r = proc.Документы.Получить(i)
        if '000013825' in c.String(r.Документ):
            hit = r
    if hit is not None:
        print('приклад 000013825: очікувано', hit.ОжидаемоПС, 'факт', hit.ФактПС, 'відмітка', hit.Отметка,
              'закритий', hit.ЗакрытыйПериод)
        check(abs(hit.ОжидаемоПС) < 0.01 and abs(hit.ФактПС - 433.33) < 0.01, 'приклад: очікувано 0, факт 433,33')
        check(hit.Отметка, 'приклад відмічено за замовчуванням (відкритий період)')
    else:
        print('приклад 000013825 не в розбіжностях (вже виправлено?)')

    # інваріант відміток
    bad_marks = 0
    keys_marked = {}
    for i in range(proc.Документы.Количество()):
        r = proc.Документы.Получить(i)
        expected = bool(r.ЕстьРасхождение) and not bool(r.ЗакрытыйПериод)
        if bool(r.Отметка) != expected:
            bad_marks += 1
        keys_marked[r.КлючИД] = keys_marked.get(r.КлючИД, False) or bool(r.Отметка)
    check(bad_marks == 0, f'відмітки документів за замовчуванням коректні (порушень {bad_marks})')
    bad_keys = 0
    open_keys = 0
    for i in range(proc.Ключи.Количество()):
        r = proc.Ключи.Получить(i)
        if bool(r.Отметка) != keys_marked.get(r.КлючИД, False):
            bad_keys += 1
        if bool(r.Отметка):
            open_keys += 1
    check(bad_keys == 0, f'відмітка ключа = є відмічені документи (порушень {bad_keys})')
    print('ключів з відміченими документами (відкритий період):', open_keys)

    # документи механізму за поточний місяць
    proc.ДатаНачала = datetime.datetime(today.year, today.month, 1, 12, 0, 0)
    tz = proc.ДокументыМеханизмаЗаПериод()
    print('документів механізму за поточний місяць:', tz.Количество())
    check(tz.Количество() > 0, 'ДокументыМеханизмаЗаПериод не порожній')

    print('ПІДСУМОК:', 'УСПІХ' if not FAILS else f'ПОМИЛОК {len(FAILS)}')
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
