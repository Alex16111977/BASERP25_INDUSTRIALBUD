# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

# Список ЕДРПОУ организаций (как org-фильтр). Используем КодПоЕДРПОУ В список.
edrpou_list = ["41597184","45117388","45143198","45067600","43698485",
               "40645273","44590697","44628382","3244112838"]

def test_date(d):
    # КонецДня = начало след.дня (виртуальная таблица Остатки на момент)
    moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)
    q = buh.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
        Ост.Организация.Наименование КАК Орг,
        Ост.Счет.Код КАК Счет,
        Ост.Счет.Наименование КАК СчетНаим,
        Ост.СуммаОстатокДт КАК СальдоДт,
        Ост.СуммаОстатокКт КАК СальдоКт,
        Ост.СуммаОстаток КАК Сальдо
    ИЗ
        РегистрБухгалтерии.Хозрасчетный.Остатки(&Момент, , , ) КАК Ост
    ГДЕ
        Ост.СуммаОстаток <> 0
        И (Ост.Счет.Код ПОДОБНО "641%" ИЛИ Ост.Счет.Код ПОДОБНО "642%" ИЛИ Ост.Счет.Код ПОДОБНО "651%")
        И Ост.Организация.КодПоЕДРПОУ В (&Список)
    УПОРЯДОЧИТЬ ПО ЕДРПОУ, Счет
    """
    масс = buh.NewObject("Массив")
    for e in edrpou_list:
        масс.Добавить(e)
    q.SetParameter("Момент", moment)
    q.SetParameter("Список", масс)
    try:
        r = q.Execute().Выгрузить()
        print(f"\n=== BuhBud Хозрасчетный Остатки на {d.strftime('%d.%m.%Y')} 23:59:59 — рядкiв={r.Количество()} ===")
        итог = 0
        for s in r:
            знак = "Кт" if s.Сальдо < 0 else "Дт"
            print(f"  ЕДРПОУ={str(s.ЕДРПОУ):11} | сч{s.Счет:6} | Дт={s.СальдоДт:>16,.2f} | Кт={s.СальдоКт:>16,.2f} | Сальдо={s.Сальдо:>16,.2f} ({знак}) | {s.Орг[:25]}")
            итог += s.Сальдо
        print(f"  ИТОГО Сальдо(Дт-Кт)={итог:,.2f}")
        return r.Количество()
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo:
            print(f"FAIL {d}: {e.excepinfo[2]}")
        else:
            print(f"FAIL {d}: {e}")
        return -1

for d in [datetime.date(2025,11,30), datetime.date(2025,6,30), datetime.date(2025,12,31)]:
    test_date(d)
