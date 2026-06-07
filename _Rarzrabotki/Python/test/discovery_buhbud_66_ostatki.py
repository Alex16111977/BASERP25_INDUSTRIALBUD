"""Discovery Q2 (исправлено): .Остатки 66 на дату, Дт/Кт per счёт (группировка по Ост.Счет)."""
import sys
from datetime import datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
S = buh.String

for dt in [datetime(2025, 12, 31, 12, 0, 0), datetime(2026, 4, 30, 12, 0, 0), datetime(2026, 5, 31, 12, 0, 0)]:
    q = buh.NewObject("Запрос")
    q.УстановитьПараметр("Дата", dt)
    q.Text = """
    ВЫБРАТЬ
        Ост.Счет.Код КАК Код,
        КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица)) КАК ЧислоФЛ,
        СУММА(Ост.СуммаОстатокКт) КАК Кт,
        СУММА(Ост.СуммаОстатокДт) КАК Дт
    ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Дата, Счет В ИЕРАРХИИ(ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.РасчетыПоОплатеТруда)), , ) КАК Ост
    ГДЕ Ост.Субконто1 ССЫЛКА Справочник.ФизическиеЛица
    СГРУППИРОВАТЬ ПО Ост.Счет
    УПОРЯДОЧИТЬ ПО Код
    """
    try:
        r = q.Выполнить().Выгрузить()
        tot_kt = sum(float(x.Кт) for x in r)
        tot_dt = sum(float(x.Дт) for x in r)
        print(f"\n.Остатки на {dt:%Y-%m-%d}: счетов={r.Количество()}, ΣКт={tot_kt:,.2f}, ΣДт={tot_dt:,.2f}, Кт-Дт={tot_kt-tot_dt:,.2f}")
        for row in r:
            print(f"   {S(row.Код):6} | ФЛ={int(row.ЧислоФЛ):4} | Кт={float(row.Кт):>15,.2f} | Дт={float(row.Дт):>12,.2f} | Кт-Дт={float(row.Кт)-float(row.Дт):>15,.2f}")
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        print(f".Остатки {dt:%Y-%m-%d} FAIL: {msg}")

print("\nDONE")
