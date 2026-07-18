# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def run(reg, dogovor, contr):
    q = erp.NewObject("Запрос")
    q.Текст = f"""
    ВЫБРАТЬ
        ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор) КАК Док,
        Р.ДокументРегистратор.Дата          КАК Дата,
        Р.ДокументРегистратор                КАК Реф,
        СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                    ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ) КАК Долг
    ИЗ
        РегистрНакопления.{reg} КАК Р
        ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК А
            ПО Р.АналитикаУчетаПоПартнерам = А.КлючАналитики
    ГДЕ
        Р.Период <= КОНЕЦПЕРИОДА(ДАТАВРЕМЯ(2025, 12, 1), МЕСЯЦ)
        И Р.ОбъектРасчетов.Договор.Наименование = "{dogovor}"
        И А.Контрагент.Наименование = "{contr}"
    СГРУППИРОВАТЬ ПО
        Р.ДокументРегистратор, Р.ДокументРегистратор.Дата
    УПОРЯДОЧИТЬ ПО
        Дата
    """
    sel = q.Выполнить().Выбрать()
    rows = []
    while sel.Следующий():
        tp = ""
        try:
            tp = sel.Реф.Метаданные().Имя
        except Exception:
            tp = "?"
        rows.append((sel.Дата, sel.Док, tp, round(float(sel.Долг), 2)))
    return rows

def show(title, rows):
    print(f"\n===== {title} =====")
    tot = 0.0
    for d, dok, tp, v in rows:
        tot += v
        ds = d.strftime("%d.%m.%Y") if hasattr(d, "strftime") else str(d)
        print(f"  {ds}  {tp:<34}  {v:>15,.2f}   {dok[:60]}")
    print(f"  {'ИТОГО':<48}  {tot:>15,.2f}   (строк {len(rows)})")
    return tot

A = run("РасчетыСКлиентамиПоСрокам", "Астарта. Тищенки внутр.(Реализация)", "Строительство")
B = run("РасчетыСПоставщикамиПоСрокам", "Строительство внутр.(Закупка)", "Астарта. Тищенки")

tA = show("РЕАЛИЗАЦИЯ  (РСКПС: Астарта.Тищенки внутр.(Реализация), клиент Строительство)", A)
tB = show("ЗАКУПКА     (РСППС: Строительство внутр.(Закупка), поставщик Астарта.Тищенки)", B)

print(f"\n>>> РЕАЛИЗАЦИЯ={tA:,.2f}  ЗАКУПКА={tB:,.2f}  РАЗНИЦА={tA - tB:,.2f}")

# align by amount to find the unmatched doc(s)
from collections import Counter
ca = Counter(v for *_, v in A)
cb = Counter(v for *_, v in B)
only_a = ca - cb
only_b = cb - ca
print("\n--- суммы, что есть в РЕАЛИЗАЦИИ, но нет (или иной) в ЗАКУПКЕ ---")
for v, n in sorted(only_a.items(), key=lambda x: -abs(x[0])):
    docs = [f"{d.strftime('%d.%m.%Y')} {dok}" for d, dok, tp, vv in A if vv == v][:n]
    for dd in docs:
        print(f"  {v:>15,.2f}   {dd}")
print("\n--- суммы, что есть в ЗАКУПКЕ, но нет (или иной) в РЕАЛИЗАЦИИ ---")
for v, n in sorted(only_b.items(), key=lambda x: -abs(x[0])):
    docs = [f"{d.strftime('%d.%m.%Y')} {dok}" for d, dok, tp, vv in B if vv == v][:n]
    for dd in docs:
        print(f"  {v:>15,.2f}   {dd}")
