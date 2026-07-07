# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Рег.Подразделение) КАК Подр,
    ПРЕДСТАВЛЕНИЕ(Рег.Контрагент)    КАК Контр,
    ПРЕДСТАВЛЕНИЕ(Рег.Статья)        КАК Статья,
    СУММА(Рег.СуммаКонечныйОстаток)  КАК КонОст,
    КОЛИЧЕСТВО(*)                    КАК Кол
ИЗ
    РегистрСведений.А_ОтчетБаланс_Свод КАК Рег
ГДЕ
    НАЧАЛОПЕРИОДА(Рег.Регистратор.Дата, МЕСЯЦ) = ДАТАВРЕМЯ(2025, 12, 1)
    И Рег.Договор.А_ВидКонтрагента.Наименование = "Внутренние подразделения"
    И (Рег.Source = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам)
       ИЛИ Рег.Source = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам))
СГРУППИРОВАТЬ ПО
    Рег.Подразделение, Рег.Контрагент, Рег.Статья
"""
sel = q.Выполнить().Выбрать()
ZK = "Задолженность клиентов"
ZP = "Задолженность перед поставщиками"
recv = {}   # (A,B) -> ЗадКл  (A's claim on B)
pay  = {}   # (A,B) -> ЗадПеред (A owes B)  stored as signed (negative)
other = {}
tot = 0.0
while sel.Следующий():
    a = sel.Подр; b = sel.Контр; st = sel.Статья; v = float(sel.КонОст)
    tot += v
    if st == ZK:
        recv[(a, b)] = recv.get((a, b), 0.0) + v
    elif st == ZP:
        pay[(a, b)] = pay.get((a, b), 0.0) + v
    else:
        other[st] = other.get(st, 0.0) + v

print(f"Сумма КонОст (2 статьи долга) = {tot:,.2f}")
print(f"Σ ЗадКл  = {sum(recv.values()):,.2f}  (пар={len(recv)})")
print(f"Σ ЗадПеред = {sum(pay.values()):,.2f}  (пар={len(pay)})")
if other:
    print("ПРОЧИЕ статьи (не должно быть):", other)

# Mirror check: recv[(A,B)] должен = -pay[(B,A)]
print("\n=== РАЗЛЁТ по зеркальным парам (A claim on B)  vs  (B owes A) ===")
pairs = set(recv) | {(b, a) for (a, b) in pay}
rows = []
for (a, b) in pairs:
    r = recv.get((a, b), 0.0)        # A's receivable from B
    p = pay.get((b, a), 0.0)         # B's payable to A (signed negative)
    resid = r + p                    # должно быть 0
    if abs(resid) > 0.005:
        rows.append((resid, a, b, r, p))
rows.sort(key=lambda x: -abs(x[0]))
print(f"{'Резид':>14} | {'Подр(A)':<22} | {'Контр(B)':<22} | {'ЗадКл A→B':>14} | {'ЗадПеред B→A':>14}")
s = 0.0
for resid, a, b, r, p in rows:
    s += resid
    print(f"{resid:>14,.2f} | {a[:22]:<22} | {b[:22]:<22} | {r:>14,.2f} | {p:>14,.2f}")
print(f"\nΣ residual = {s:,.2f}  (всего незеркальных пар: {len(rows)})")
