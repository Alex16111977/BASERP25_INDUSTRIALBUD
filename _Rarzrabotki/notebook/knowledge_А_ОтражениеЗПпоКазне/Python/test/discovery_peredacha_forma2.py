# -*- coding: utf-8 -*-
"""Discovery (Rule #-1) для PROMPT_ПередачаНачисленийФорма2_Корень.
Эталон №000000005: Σ Ф2 переносимых начислений, карта подразделение->корень, no-op."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка ИЗ Документ.А_ОтражениеЗПпоКазне ГДЕ Номер = &Н"
q.УстановитьПараметр("Н", "000000005")
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("Документ №000000005 не найден"); sys.exit(1)
doc = sel.Ссылка.ПолучитьОбъект()
print("Док:", doc.Номер, "от", doc.Дата, "Проведен =", doc.Проведен)

def корень(п):
    тек = п
    while erp.ЗначениеЗаполнено(тек) and erp.ЗначениеЗаполнено(тек.Родитель):
        тек = тек.Родитель
    return тек

тч = doc.НачисленнаяЗарплатаИВзносыПоФизлицам
n_rows = n_noop = 0
sum_f2 = sum_noop = sum_transfer = 0.0
roots = {}
noop_podr = set()
no_podr = 0
for i in range(тч.Количество()):
    row = тч.Получить(i)
    if erp.XMLСтрока(row.ФормаPL) != "Форма2":
        continue
    if float(row.ВзносыВсего) != 0:
        continue
    sp = row.СпособОтраженияЗарплатыВБухучете
    if erp.ЗначениеЗаполнено(sp) and sp.А_ЭтоУдержание:
        continue
    s = float(row.Сумма)
    if s == 0:
        continue
    n_rows += 1; sum_f2 += s
    подр = row.ПодразделениеПредприятия
    if not erp.ЗначениеЗаполнено(подр):
        no_podr += 1; continue
    к = корень(подр)
    pn = подр.Наименование
    roots[pn] = (к.Наименование if erp.ЗначениеЗаполнено(к) else "(пусто)")
    if erp.ЗначениеЗаполнено(к) and erp.XMLСтрока(к) == erp.XMLСтрока(подр):
        n_noop += 1; sum_noop += s; noop_podr.add(pn)
    else:
        sum_transfer += s

print(f"\nФ2 строк-кандидатов: {n_rows} | Sum Ф2 = {round(sum_f2,2)}")
print(f"no-op (корень=исходному): строк {n_noop}, Sum {round(sum_noop,2)}")
print(f"без ПодразделениеПредприятия: {no_podr} строк")
print(f"Sum Ф2_перенос (корень<>исходному) = {round(sum_transfer,2)}")
print(f"\nКарта подразделение -> КОРЕНЬ ({len(roots)} уник.):")
for k in sorted(roots):
    flag = "  [NO-OP: уже корень]" if k in noop_podr else ""
    print(f"  {k!r:42} -> {roots[k]!r}{flag}")
