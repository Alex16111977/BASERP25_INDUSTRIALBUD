# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# ЕДРПОУ организаций с А_ВБалансе (ERP-side)
qe = erp.NewObject("Запрос")
qe.Text = """
ВЫБРАТЬ Орг.КодПоЕДРПОУ КАК ЕДРПОУ
ИЗ Справочник.Организации КАК Орг
ГДЕ Орг.А_ВБалансе И Орг.КодПоЕДРПОУ <> ""
"""
edrpou = [s.ЕДРПОУ for s in qe.Execute().Выгрузить()]
print(f"ЕДРПОУ с А_ВБалансе (ERP): {edrpou}")

масс_e = erp.NewObject("Массив")
for e in edrpou:
    масс_e.Добавить(e)

def test_erp_buh(d):
    moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
        Ост.Организация.Наименование КАК Орг,
        Ост.Счет.Код КАК Счет,
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
    q.SetParameter("Момент", moment)
    q.SetParameter("Список", масс_e)
    try:
        r = q.Execute().Выгрузить()
        print(f"\n=== ERP Хозрасчетный Остатки на {d.strftime('%d.%m.%Y')} — рядкiв={r.Количество()} ===")
        итог = 0
        for s in r:
            знак = "Кт" if s.Сальдо < 0 else "Дт"
            print(f"  ЕДРПОУ={str(s.ЕДРПОУ):11} | сч{s.Счет:6} | Дт={s.СальдоДт:>15,.2f} | Кт={s.СальдоКт:>15,.2f} | Сальдо={s.Сальдо:>15,.2f} ({знак}) | {s.Орг[:22]}")
            итог += s.Сальдо
        print(f"  ИТОГО Сальдо(Дт-Кт)={итог:,.2f}")
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo:
            print(f"FAIL {d}: {e.excepinfo[2]}")
        else:
            print(f"FAIL {d}: {e}")

for d in [datetime.date(2025,11,30), datetime.date(2025,12,31)]:
    test_erp_buh(d)
