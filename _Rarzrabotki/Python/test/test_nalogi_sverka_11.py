# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

d = datetime.date(2025,12,31)
moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)

# BuhBud: остатки по счетам с субконто "Податки" — разрез Счёт x Вид налога (субконто)
# для собственной орг 40645273
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Ост.Счет.Код КАК Счет,
    ПРЕДСТАВЛЕНИЕ(Ост.Субконто1) КАК Субконто1,
    Ост.СуммаОстаток КАК Сальдо
ИЗ
    РегистрБухгалтерии.Хозрасчетный.Остатки(
        &Момент,
        ,
        ,
        Организация.КодПоЕДРПОУ = &ЕДРПОУ) КАК Ост
ГДЕ
    Ост.СуммаОстаток <> 0
    И (Ост.Счет.Код ПОДОБНО "641%" ИЛИ Ост.Счет.Код ПОДОБНО "642%" ИЛИ Ост.Счет.Код ПОДОБНО "651%")
УПОРЯДОЧИТЬ ПО Счет
"""
q.SetParameter("Момент", moment)
q.SetParameter("ЕДРПОУ", "40645273")
try:
    r = q.Execute().Выгрузить()
    print(f"=== BuhBud ІНДАСТРІАЛБУД остатки Счёт x Субконто1 (Податки) на 31.12.2025 — рядкiв={r.Количество()} ===")
    for s in r:
        суб = str(s.Субконто1) if s.Субконто1 else "(без субконто)"
        print(f"  сч{s.Счет:6} | субконто='{суб:30}' | Сальдо={s.Сальдо:>15,.2f}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"  FAIL: {e.excepinfo[2]}")
    else:
        print(f"  FAIL: {e}")

# Виды налогов (субконто Податки) в BuhBud — справочник
print("\n=== BuhBud: справочник видов налогов (субконто Податки) ===")
q2 = buh.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ РАЗЛИЧНЫЕ
    ПРЕДСТАВЛЕНИЕ(Дв.Субконто1) КАК Налог,
    Дв.Счет.Код КАК Счет
ИЗ РегистрБухгалтерии.Хозрасчетный КАК Дв
ГДЕ (Дв.Счет.Код ПОДОБНО "641%" ИЛИ Дв.Счет.Код ПОДОБНО "642%" ИЛИ Дв.Счет.Код ПОДОБНО "651%")
    И Дв.Субконто1 <> НЕОПРЕДЕЛЕНО
"""
try:
    r2 = q2.Execute().Выгрузить()
    seen=set()
    for s in r2:
        key=(s.Счет, str(s.Налог))
        if key in seen: continue
        seen.add(key)
        print(f"  сч{s.Счет:6} -> '{s.Налог}'")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"  FAIL: {e.excepinfo[2]}")
    else:
        print(f"  FAIL: {e}")
