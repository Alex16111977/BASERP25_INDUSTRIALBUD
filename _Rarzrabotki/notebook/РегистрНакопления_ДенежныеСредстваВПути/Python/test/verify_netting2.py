# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
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
q.Текст = """ВЫБРАТЬ
    Организация,
    Подразделение,
    Статья,
    СУММА(СЛУЧАЙ КОГДА ВидДвижения = 0 THEN Сумма ELSE 0 END) КАК Приход,
    СУММА(СЛУЧАЙ КОГДА ВидДвижения = 1 THEN Сумма ELSE 0 END) КАК Расход,
    СУММА(СЛУЧАЙ КОГДА ВидДвижения = 0 THEN Сумма ИНАЧЕ -Сумма END) КАК Сальдо
FROM РегистрНакопления.ПрочиеАктивыПассивы
WHERE (Регистратор = &РКО OR Регистратор = &ПКО)
  AND Статья.Наименование LIKE '%Денежные средства в пути%'
GROUP BY Организация, Подразделение, Статья"""

try:
    rr = q.Выполнить().Выгрузить()
    rows_pap = []
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        org = str(S(rec.Организация)) if erp.ЗначениеЗаполнено(rec.Организация) else "?"
        podr = str(S(rec.Подразделение)) if erp.ЗначениеЗаполнено(rec.Подразделение) else "?"
        prikhod = float(rec.Приход or 0)
        raskhod = float(rec.Расход or 0)
        saldo = float(rec.Сальдо or 0)
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
q2.Текст = """ВЫБРАТЬ
    Организация,
    Подразделение,
    СУММА(СЛУЧАЙ КОГДА ВидДвижения = 0 THEN СуммаУпр ELSE 0 END) КАК Приход,
    СУММА(СЛУЧАЙ КОГДА ВидДвижения = 1 THEN СуммаУпр ELSE 0 END) КАК Расход,
    СУММА(СЛУЧАЙ КОГДА ВидДвижения = 0 THEN СуммаУпр ИНАЧЕ -СуммаУпр END) КАК Сальдо
FROM РегистрНакопления.ДенежныеСредстваВПути
WHERE Регистратор = &РКО OR Регистратор = &ПКО
GROUP BY Организация, Подразделение"""

try:
    rr2 = q2.Выполнить().Выгрузить()
    rows_dsv = []
    for i in range(rr2.Количество()):
        rec = rr2.Получить(i)
        org = str(S(rec.Организация)) if erp.ЗначениеЗаполнено(rec.Организация) else "?"
        podr = str(S(rec.Подразделение)) if erp.ЗначениеЗаполнено(rec.Подразделение) else "?"
        prikhod = float(rec.Приход or 0)
        raskhod = float(rec.Расход or 0)
        saldo = float(rec.Сальдо or 0)
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
