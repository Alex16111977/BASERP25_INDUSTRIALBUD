# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

d = datetime.date(2025,12,31)
moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)

# ЕДРПОУ А_ВБалансе из ERP
qe = erp.NewObject("Запрос")
qe.Text = """ВЫБРАТЬ Орг.КодПоЕДРПОУ КАК ЕДРПОУ ИЗ Справочник.Организации КАК Орг
ГДЕ Орг.А_ВБалансе И Орг.КодПоЕДРПОУ <> "" """
edrpou = [s.ЕДРПОУ for s in qe.Execute().Выгрузить()]
print(f"ЕДРПОУ А_ВБалансе (ERP-side фильтр): {edrpou}\n")

def saldo_buh_hozras(conn, edrpou_list, moment, lk):
    масс = conn.NewObject("Массив")
    for e in edrpou_list: масс.Добавить(e)
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ, Ост.Счет.Код КАК Счет, Ост.СуммаОстаток КАК Сальдо
    ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Момент, , , ) КАК Ост
    ГДЕ Ост.СуммаОстаток <> 0
      И (Ост.Счет.Код ПОДОБНО "641%" ИЛИ Ост.Счет.Код ПОДОБНО "642%" ИЛИ Ост.Счет.Код ПОДОБНО "651%")
      И Ост.Организация.КодПоЕДРПОУ В (&Список)
    """
    q.SetParameter("Момент", moment); q.SetParameter("Список", масс)
    res = {}
    for s in q.Execute().Выгрузить():
        # нормализуем код счёта до базового (6411, 6412... 642, 651)
        к = s.Счет
        res[(s.ЕДРПОУ, к)] = res.get((s.ЕДРПОУ, к),0) + s.Сальдо
    return res

buh_h = saldo_buh_hozras(buh, edrpou, moment, "BUH")
erp_h = saldo_buh_hozras(erp, edrpou, moment, "ERP")

print("=== СВЕРКА Орг(ЕДРПОУ) x Счёт — Хозрасчетный (regl) обе стороны на 31.12.2025 ===")
print(f"{'ЕДРПОУ':12} {'Счёт':6} {'Сальдо_ЕРП':>16} {'Сальдо_Бух':>16} {'Разница':>16}")
ключи = sorted(set(list(buh_h.keys()) + list(erp_h.keys())))
итЕ=итБ=0
for k in ключи:
    e = erp_h.get(k,0); b = buh_h.get(k,0); diff = e-b
    итЕ+=e; итБ+=b
    флаг = "" if abs(diff)<0.01 else "  <-- РАЗНИЦА"
    print(f"{k[0]:12} {k[1]:6} {e:>16,.2f} {b:>16,.2f} {diff:>16,.2f}{флаг}")
print(f"{'ИТОГО':12} {'':6} {итЕ:>16,.2f} {итБ:>16,.2f} {итЕ-итБ:>16,.2f}")
