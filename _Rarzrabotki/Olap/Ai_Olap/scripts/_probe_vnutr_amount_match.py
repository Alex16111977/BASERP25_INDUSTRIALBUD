# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
from collections import Counter, defaultdict

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Рег.Статья)          КАК Статья,
    ПРЕДСТАВЛЕНИЕ(Рег.ДокументДвижения) КАК Док,
    ПРЕДСТАВЛЕНИЕ(Рег.Контрагент)      КАК Контр,
    Рег.СуммаКонечныйОстаток           КАК КонОст
ИЗ
    РегистрСведений.А_ОтчетБаланс_Свод КАК Рег
ГДЕ
    НАЧАЛОПЕРИОДА(Рег.Регистратор.Дата, МЕСЯЦ) = ДАТАВРЕМЯ(2025, 12, 1)
    И Рег.Договор.А_ВидКонтрагента.Наименование = "Внутренние подразделения"
    И (Рег.Source = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам)
       ИЛИ Рег.Source = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам))
    И (Рег.Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов)
       ИЛИ Рег.Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками))
"""
sel = q.Выполнить().Выбрать()
ZK = "Задолженность клиентов"
recv = []  # client debt rows  (amount>0)
pay  = []  # supplier debt rows (amount<0)
while sel.Следующий():
    row = (round(float(sel.КонОст), 2), sel.Док, sel.Контр)
    if sel.Статья == ZK:
        recv.append(row)
    else:
        pay.append(row)

sr = sum(r[0] for r in recv); sp = sum(r[0] for r in pay)
print(f"ЗадКл  rows={len(recv)}  Σ={sr:,.2f}")
print(f"ЗадПеред rows={len(pay)}  Σ={sp:,.2f}")
print(f"NET = {sr+sp:,.2f}\n")

# multiset match by |amount|
rc = Counter(round(r[0], 2) for r in recv)
pc = Counter(round(-p[0], 2) for p in pay)   # flip sign of payables to positive
unmatched_recv = rc - pc       # client amounts with no supplier mirror
unmatched_pay  = pc - rc       # supplier amounts with no client mirror

def show(title, cnt, rows, signflip=False):
    print(f"=== {title} (Σ={sum(a*n for a,n in cnt.items()):,.2f}) ===")
    idx = defaultdict(list)
    for amt, dok, contr in rows:
        idx[round(-amt,2) if signflip else round(amt,2)].append((dok, contr))
    for amt, n in sorted(cnt.items(), key=lambda x:-abs(x[0]*x[1])):
        ex = idx.get(amt, [])[:n]
        for dok, contr in ex:
            print(f"  {amt:>14,.2f}  Контр={contr[:26]:<26}  Док={dok[:60]}")

show("НЕсматченные ЗадКл (есть долг клиента, нет зеркала у поставщика)", unmatched_recv, recv)
print()
show("НЕсматченные ЗадПеред (есть долг поставщику, нет зеркала у клиента)", unmatched_pay, pay, signflip=True)

print(f"\nΣ нематч ЗадКл  = {sum(a*n for a,n in unmatched_recv.items()):,.2f}")
print(f"Σ нематч ЗадПеред = {-sum(a*n for a,n in unmatched_pay.items()):,.2f}")
print(f"Разница (= NET разлёт) = {sum(a*n for a,n in unmatched_recv.items()) - sum(a*n for a,n in unmatched_pay.items()):,.2f}")
