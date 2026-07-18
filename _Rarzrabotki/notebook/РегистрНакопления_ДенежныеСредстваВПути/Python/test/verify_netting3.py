# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

out = []
out.append("="*80)
out.append("ПРОВЕРКА НЕТТИНГА ПО ПОДРАЗДЕЛЕНИЮ")
out.append("Пара: РКО N0000052985 (Приход) + ПКО N0000023550 (Расход)")
out.append("="*80)
out.append("")

docs = {"РКО": "N0000052985", "ПКО": "N0000023550"}
doc_refs = {}

for dtype, num in docs.items():
    if dtype == "РКО":
        q = erp.NewObject("Запрос")
        q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = "' + num + '"'
    else:
        q = erp.NewObject("Запрос")
        q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПриходныйКассовыйОрдер ГДЕ Номер = "' + num + '"'
    
    sel = q.Выполнить().Выбрать()
    if sel.Следующий():
        doc_refs[dtype] = sel.Ссылка
        out.append("[" + dtype + "] " + str(S(doc_refs[dtype])))
    else:
        out.append("[" + dtype + "] НЕ НАЙДЕН")

out.append("")
out.append("-"*80)
out.append("РегистрНакопления.ПрочиеАктивыПассивы")
out.append("Статья: Денежные средства в пути")
out.append("-"*80)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("РКО", doc_refs.get("РКО"))
q.УстановитьПараметр("ПКО", doc_refs.get("ПКО"))
q.Текст = "ВЫБРАТЬ Р.Организация, Р.Подразделение, Р.Статья, Р.ВидДвижения, СУММА(Р.Сумма) КАК Сумма ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Р ГДЕ (Р.Регистратор = &РКО ИЛИ Р.Регистратор = &ПКО) И Р.Статья.Наименование LIKE '%Денежные средства в пути%' СГРУППИРОВАТЬ ПО Р.Организация, Р.Подразделение, Р.Статья, Р.ВидДвижения"

try:
    rr = q.Выполнить().Выгрузить()
    rows_pap = []
    saldo_map = {}
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        org = str(S(rec.Организация)) if erp.ЗначениеЗаполнено(rec.Организация) else "?"
        podr = str(S(rec.Подразделение)) if erp.ЗначениеЗаполнено(rec.Подразделение) else "?"
        vd = int(rec.ВидДвижения)
        summa = float(rec.Сумма or 0)
        
        key = (org, podr)
        if key not in saldo_map:
            saldo_map[key] = {"prikhod": 0.0, "raskhod": 0.0}
        
        if vd == 0:
            saldo_map[key]["prikhod"] += summa
        else:
            saldo_map[key]["raskhod"] += summa
    
    for (org, podr), amounts in sorted(saldo_map.items()):
        prikhod = amounts["prikhod"]
        raskhod = amounts["raskhod"]
        saldo = prikhod - raskhod
        rows_pap.append({"org": org, "podr": podr, "prikhod": prikhod, "raskhod": raskhod, "saldo": saldo})
        out.append("Org=%s | Podr=%s | Приход=%12.2f | Расход=%12.2f | Сальдо=%12.2f" % (org.ljust(14), podr.ljust(20), prikhod, raskhod, saldo))
    
    if not rows_pap:
        out.append("(нет строк)")
except Exception as e:
    out.append("[ERROR] " + str(e))
    import traceback
    out.append(traceback.format_exc())
    rows_pap = []

out.append("")
out.append("-"*80)
out.append("РегистрНакопления.ДенежныеСредстваВПути")
out.append("-"*80)

q2 = erp.NewObject("Запрос")
q2.УстановитьПараметр("РКО", doc_refs.get("РКО"))
q2.УстановитьПараметр("ПКО", doc_refs.get("ПКО"))
q2.Текст = "ВЫБРАТЬ Р.Организация, Р.Подразделение, Р.ВидДвижения, СУММА(Р.СуммаУпр) КАК Сумма ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Р ГДЕ (Р.Регистратор = &РКО ИЛИ Р.Регистратор = &ПКО) СГРУППИРОВАТЬ ПО Р.Организация, Р.Подразделение, Р.ВидДвижения"

try:
    rr2 = q2.Выполнить().Выгрузить()
    rows_dsv = []
    saldo_map2 = {}
    for i in range(rr2.Количество()):
        rec = rr2.Получить(i)
        org = str(S(rec.Организация)) if erp.ЗначениеЗаполнено(rec.Организация) else "?"
        podr = str(S(rec.Подразделение)) if erp.ЗначениеЗаполнено(rec.Подразделение) else "?"
        vd = int(rec.ВидДвижения)
        summa = float(rec.Сумма or 0)
        
        key = (org, podr)
        if key not in saldo_map2:
            saldo_map2[key] = {"prikhod": 0.0, "raskhod": 0.0}
        
        if vd == 0:
            saldo_map2[key]["prikhod"] += summa
        else:
            saldo_map2[key]["raskhod"] += summa
    
    for (org, podr), amounts in sorted(saldo_map2.items()):
        prikhod = amounts["prikhod"]
        raskhod = amounts["raskhod"]
        saldo = prikhod - raskhod
        rows_dsv.append({"org": org, "podr": podr, "prikhod": prikhod, "raskhod": raskhod, "saldo": saldo})
        out.append("Org=%s | Podr=%s | Приход=%12.2f | Расход=%12.2f | Сальдо=%12.2f" % (org.ljust(14), podr.ljust(20), prikhod, raskhod, saldo))
    
    if not rows_dsv:
        out.append("(нет строк)")
except Exception as e:
    out.append("[ERROR] " + str(e))
    import traceback
    out.append(traceback.format_exc())
    rows_dsv = []

out.append("")
out.append("="*80)
out.append("АНАЛИЗ НЕТТИНГА")
out.append("="*80)
out.append("")

has_fail = False

out.append("ПрочиеАктивыПассивы по подразделению:")
for row in rows_pap:
    if abs(row["saldo"]) > 0.01:
        out.append("  X %s: Сальдо = %.2f (НЕ НОЛЬ) — FAIL" % (row['podr'], row['saldo']))
        has_fail = True
    else:
        out.append("  + %s: Сальдо = %.2f" % (row['podr'], row['saldo']))

out.append("")
out.append("ДенежныеСредстваВПути по подразделению:")
for row in rows_dsv:
    if abs(row["saldo"]) > 0.01:
        out.append("  X %s: Сальдо = %.2f (НЕ НОЛЬ) — FAIL" % (row['podr'], row['saldo']))
        has_fail = True
    else:
        out.append("  + %s: Сальдо = %.2f" % (row['podr'], row['saldo']))

out.append("")
out.append("="*80)
if not has_fail:
    out.append("РЕЗУЛЬТАТ: PASS — весь транзит сворачивается по подразделению")
else:
    out.append("РЕЗУЛЬТАТ: FAIL — есть ненулевые остатки")
out.append("="*80)

with open("C:\\temp\\netting_result.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))

for line in out:
    print(line)

sys.exit(1 if has_fail else 0)
