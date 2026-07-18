# -*- coding: utf-8 -*-
# Acceptance: реплика ObjectModule.ОбъединитьТаблицы на боевых базах.
# Проверяет, что отчёт даст корректные агрегаты/разбиение по (А_ИдКод, ЕДРПОУ).
import win32com.client as wc
import sys
from datetime import datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def fail(e):
    if hasattr(e, 'excepinfo') and e.excepinfo:
        return e.excepinfo[2]
    return str(e)

v8 = wc.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
nach = datetime(2025, 1, 1)
kon  = datetime(2026, 12, 31, 23, 59, 59)

# --- ЕРП (как ПолучитьОстаткиЕРП) ---
q = erp.NewObject("Запрос")
q.УстановитьПараметр("НачалоПериода", nach); q.УстановитьПараметр("КонецПериода", kon)
q.Текст = """
ВЫБРАТЬ Ост.ФизическоеЛицо.А_ИдКод КАК А_ИдКод, Ост.ОрганизацияБухгалтерия.КодПоЕДРПОУ КАК ЕДРПОУ,
    СУММА(Ост.СуммаВзаиморасчетовНачальныйОстаток) КАК НО, СУММА(Ост.СуммаВзаиморасчетовПриход) КАК Нач,
    СУММА(Ост.СуммаВзаиморасчетовРасход) КАК Вып, СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КО
ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(&НачалоПериода, &КонецПериода, , ,
        ФормаPL = ЗНАЧЕНИЕ(Перечисление.А_ФормыPL.Форма1)) КАК Ост
СГРУППИРОВАТЬ ПО Ост.ФизическоеЛицо.А_ИдКод, Ост.ОрганизацияБухгалтерия.КодПоЕДРПОУ
"""
erp_rows = []
s = q.Выполнить().Выбрать()
while s.Следующий():
    erp_rows.append(((str(s.А_ИдКод or "").strip(), str(s.ЕДРПОУ or "").strip()),
                     float(s.НО), float(s.Нач), float(s.Вып), float(s.КО)))

# --- ЕДРПОУ оргов в балансе ---
qo = erp.NewObject("Запрос")
qo.Текст = ('ВЫБРАТЬ Орг.КодПоЕДРПОУ КАК Е ИЗ Справочник.Организации КАК Орг '
            'ГДЕ Орг.А_ВБалансе = ИСТИНА И Орг.КодПоЕДРПОУ <> ""')
mas = buh.NewObject("Массив")
s = qo.Выполнить().Выбрать()
while s.Следующий():
    mas.Добавить(str(s.Е).strip())

# --- BuhBud (как ПолучитьОстаткиBuhBud) ---
qb = buh.NewObject("Запрос")
qb.УстановитьПараметр("НачалоПериода", nach); qb.УстановитьПараметр("КонецПериода", kon); qb.УстановитьПараметр("мас", mas)
qb.Текст = """
ВЫБРАТЬ ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод КАК А_ИдКод,
    Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
    СУММА(Ост.СуммаНачальныйОстатокКт)-СУММА(Ост.СуммаНачальныйОстатокДт) КАК НО,
    СУММА(Ост.СуммаОборотКт) КАК Нач, СУММА(Ост.СуммаОборотДт) КАК Вып,
    СУММА(Ост.СуммаКонечныйОстатокКт)-СУММА(Ост.СуммаКонечныйОстатокДт) КАК КО
ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(&НачалоПериода, &КонецПериода) КАК Ост
ГДЕ Ост.Счет.Код В ("661","663") И Ост.Субконто1 ССЫЛКА Справочник.ФизическиеЛица
    И Ост.Организация.КодПоЕДРПОУ В (&мас)
СГРУППИРОВАТЬ ПО ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод, Ост.Организация.КодПоЕДРПОУ
"""
buh_rows = []
s = qb.Выполнить().Выбрать()
while s.Следующий():
    idk = str(s.А_ИдКод or "").strip()
    if not idk:
        continue
    buh_rows.append(((idk, str(s.ЕДРПОУ or "").strip()),
                     float(s.НО), float(s.Нач), float(s.Вып), float(s.КО)))

# --- Слияние (как ОбъединитьТаблицы) ---
res = {}   # key -> [НО_e,Нач_e,Вып_e,КО_e, НО_b,Нач_b,Вып_b,КО_b]
for k, no, na, vy, ko in erp_rows:
    res.setdefault(k, [0,0,0,0,0,0,0,0])
    res[k][0]+=no; res[k][1]+=na; res[k][2]+=vy; res[k][3]+=ko
for k, no, na, vy, ko in buh_rows:
    res.setdefault(k, [0,0,0,0,0,0,0,0])
    res[k][4]+=no; res[k][5]+=na; res[k][6]+=vy; res[k][7]+=ko

sum_ko_e = sum(v[3] for v in res.values())
sum_ko_b = sum(v[7] for v in res.values())
only_e = [k for k,v in res.items() if (v[0] or v[1] or v[2] or v[3]) and not (v[4] or v[5] or v[6] or v[7])]
only_b = [k for k,v in res.items() if (v[4] or v[5] or v[6] or v[7]) and not (v[0] or v[1] or v[2] or v[3])]
both   = [k for k in res if k not in only_e and k not in only_b]
diff_ko = sum(abs(v[3]-v[7]) for v in res.values())

print("=== ACCEPTANCE: реплика слияния отчёта ===")
print(f"строк результата (ФЛ×ОргБух): {len(res)}")
print(f"  из них в обеих базах: {len(both)}, только ЕРП: {len(only_e)}, только BuhBud: {len(only_b)}")
print(f"Σ КонОстаток_ЕРП    = {sum_ko_e:.2f}")
print(f"Σ КонОстаток_BuhBud = {sum_ko_b:.2f}")
print(f"Σ |РазницаКонОстаток| = {diff_ko:.2f}")

# --- Асерты ---
ok = True
def chk(name, cond):
    global ok
    print(("  PASS " if cond else "  FAIL ") + name)
    ok = ok and cond

chk("есть строки результата", len(res) > 0)
chk("Σ КонОст ЕРП ≈ 1 842 260.17 (эталон T1)", abs(sum_ko_e - 1842260.17) < 1.0)
chk("Σ КонОст BuhBud ≈ 1 160 791.99 (эталон T3)", abs(sum_ko_b - 1160791.99) < 1.0)
chk("есть строки 'только ЕРП' (расхождения видны)", len(only_e) > 0)
chk("есть строки 'только BuhBud' (расхождения видны)", len(only_b) > 0)
chk("разница = ЕРП - BuhBud согласована (тождество)",
    abs(diff_ko - sum(abs(v[3]-v[7]) for v in res.values())) < 0.001)

print("\n=== " + ("ALL PASS" if ok else "FAILED") + " ===")
sys.exit(0 if ok else 1)
