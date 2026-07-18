# -*- coding: utf-8 -*-
# Финальная сверочная таблица Орг(ЕДРПОУ) x Счёт, знак Кт-Дт (пассив-плюс, как в обработке переноса)
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

d = datetime.date(2025,12,31)
moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)

qe = erp.NewObject("Запрос")
qe.Text = """ВЫБРАТЬ Орг.КодПоЕДРПОУ КАК ЕДРПОУ, Орг.Наименование КАК Наим
ИЗ Справочник.Организации КАК Орг ГДЕ Орг.А_ВБалансе И Орг.КодПоЕДРПОУ <> "" """
orgmap = {}
edrpou=[]
for s in qe.Execute().Выгрузить():
    edrpou.append(s.ЕДРПОУ); orgmap[s.ЕДРПОУ]=s.Наим

def saldo(conn, moment):
    масс = conn.NewObject("Массив")
    for e in edrpou: масс.Добавить(e)
    q = conn.NewObject("Запрос")
    # знак Кт-Дт => пассив(обязательство бюджету)=плюс
    q.Text = """
    ВЫБРАТЬ Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ, Ост.Счет.Код КАК Счет,
        -Ост.СуммаОстаток КАК СальдоКтДт
    ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Момент, , , ) КАК Ост
    ГДЕ Ост.СуммаОстаток <> 0
      И (Ост.Счет.Код ПОДОБНО "641%" ИЛИ Ост.Счет.Код ПОДОБНО "642%" ИЛИ Ост.Счет.Код ПОДОБНО "651%")
      И Ост.Организация.КодПоЕДРПОУ В (&Список)
    """
    q.SetParameter("Момент", moment); q.SetParameter("Список", масс)
    res={}
    for s in q.Execute().Выгрузить():
        res[(s.ЕДРПОУ, s.Счет)] = res.get((s.ЕДРПОУ,s.Счет),0)+s.СальдоКтДт
    return res

bh = saldo(buh, moment); eh = saldo(erp, moment)
МАП={"6411":"НДФЛ","642":"ВоенныйСбор","651":"НачисленныйЕСВ","6412":"НДС","6413":"НалогНаПрибыль","6414":"ЕдиныйНалог","6415":"ДругиеНалоги","6417":"ДругиеНалоги"}

print(f"=== СВЕРКА Орг x Счёт на 31.12.2025, знак Кт-Дт (пассив=плюс) ===")
print(f"{'ЕДРПОУ':11} {'Счёт':5} {'ТипНалогаERP':16} {'Сальдо_ЕРП':>15} {'Сальдо_Бух':>15} {'Разница':>14}")
ключи=sorted(set(list(bh)+list(eh)))
tE=tB=0
for k in ключи:
    e=eh.get(k,0); b=bh.get(k,0); diff=e-b; tE+=e; tB+=b
    tn=МАП.get(k[1], "ДругиеНалоги")
    fl="" if abs(diff)<0.01 else " <-"
    print(f"{k[0]:11} {k[1]:5} {tn:16} {e:>15,.2f} {b:>15,.2f} {diff:>14,.2f}{fl}")
print(f"{'ИТОГО':11} {'':5} {'':16} {tE:>15,.2f} {tB:>15,.2f} {tE-tB:>14,.2f}")
print(f"\nОрг А_ВБалансе: " + "; ".join(f"{k}={v[:20]}" for k,v in orgmap.items()))
