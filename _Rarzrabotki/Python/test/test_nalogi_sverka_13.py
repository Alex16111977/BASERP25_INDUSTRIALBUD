# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

d = datetime.date(2025,12,31)
moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)

# Развёрнутое сальдо по субконто: для каждой (Орг,Счёт) сравним
# свёрнутое (Дт-Кт по счёту) vs сумму модулей по субконто
масс = buh.NewObject("Массив")
for e in ['40645273','41597184','44590697']: масс.Добавить(e)

q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
    Ост.Счет.Код КАК Счет,
    ПРЕДСТАВЛЕНИЕ(Ост.Субконто1) КАК Суб,
    Ост.СуммаОстаток КАК Сальдо
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Момент, , , Организация.КодПоЕДРПОУ В (&Список)) КАК Ост
ГДЕ (Ост.Счет.Код ПОДОБНО "641%" ИЛИ Ост.Счет.Код ПОДОБНО "642%" ИЛИ Ост.Счет.Код ПОДОБНО "651%")
    И Ост.СуммаОстаток <> 0
"""
q.SetParameter("Момент", moment)
q.SetParameter("Список", масс)
r = q.Execute().Выгрузить()

from collections import defaultdict
сводка = defaultdict(lambda: {"sum":0.0,"absum":0.0,"pos":0,"neg":0,"subs":[]})
for s in r:
    k=(str(s.ЕДРПОУ), s.Счет)
    сводка[k]["sum"]+=s.Сальдо
    сводка[k]["absum"]+=abs(s.Сальдо)
    if s.Сальдо>0: сводка[k]["pos"]+=1
    else: сводка[k]["neg"]+=1
    сводка[k]["subs"].append((str(s.Суб), s.Сальдо))

print("=== Развёрнутость сальдо по субконто (3 А_ВБалансе орг, 31.12) ===")
print("Если pos>0 И neg>0 — счёт активно-пассивный с развёрнутым сальдо по субконто\n")
for k in sorted(сводка):
    v=сводка[k]
    разв = "РАЗВЁРНУТО" if (v["pos"]>0 and v["neg"]>0) else "свёрнуто  "
    print(f"ЕДРПОУ={k[0]:11} сч{k[1]:6} | {разв} | Свёрнуто(Дт-Кт)={v['sum']:>14,.2f} | Σ|субконто|={v['absum']:>14,.2f}")
    if v["pos"]>0 and v["neg"]>0:
        for суб,сум in v["subs"]:
            print(f"        субконто '{суб:35}' = {сум:>14,.2f}")
