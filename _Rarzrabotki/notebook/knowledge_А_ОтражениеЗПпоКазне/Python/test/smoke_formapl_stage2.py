# -*- coding: utf-8 -*-
# Smoke Stage 2: разметка ФормаPL в НачисленнаяЗарплата/НачисленныйНДФЛ. В память (без записи).
# Инварианты: Σ(Ф1+Ф2)=текущее отражение (не меняется); Σ Ф1≈bukh; ЕСВ+НДФЛ все Ф1.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected ERP")
q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка, Номер КАК Номер ИЗ Документ.А_ОтражениеЗПпоКазне "
           "ГДЕ Дата МЕЖДУ ДАТАВРЕМЯ(2026,4,1,0,0,0) И ДАТАВРЕМЯ(2026,4,30,23,59,59) УПОРЯДОЧИТЬ ПО Дата")
sel = q.Выполнить().Выбрать(); sel.Следующий()
print(f"Документ: №{sel.Номер}")
obj = sel.Ссылка.ПолучитьОбъект()

print("Полная последовательность заполнения...")
obj.ЗаполнитьНачисленияУпризЗуп2()
obj.ЗагрузитьНалогиИУдержанияИзЗарплатыБух()
obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
obj.ЗаполнитьДляОтражения_ДокументОтражениеЗарплатыВФинансовомУчете()

def en(v):
    return erp.XMLСтрока(v) if erp.ЗначениеЗаполнено(v) else "<пусто>"

# bukh ожидаемый (из discovery)
BUKH_EXP = 2655557.75
TOTAL_EXP = 7038544.73

nz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам
sum_net_f1 = sum_net_f2 = sum_net_none = 0.0
sum_esv_f1 = sum_esv_other = 0.0
for i in range(nz.Количество()):
    r = nz.Получить(i)
    фп = en(r.ФормаPL)
    s = float(r.Сумма); vz = float(r.ВзносыВсего)
    if vz != 0:
        if фп == "Форма1": sum_esv_f1 += vz
        else: sum_esv_other += vz
    if s != 0:
        if фп == "Форма1": sum_net_f1 += s
        elif фп == "Форма2": sum_net_f2 += s
        else: sum_net_none += s

ndfl = obj.НачисленныйНДФЛ
ndfl_f1 = ndfl_other = 0.0
for i in range(ndfl.Количество()):
    r = ndfl.Получить(i)
    if en(r.ФормаPL) == "Форма1": ndfl_f1 += float(r.Сумма)
    else: ndfl_other += float(r.Сумма)

print(f"\n=== НачисленнаяЗарплата.Сумма (NET) ===")
print(f"Форма1            = {sum_net_f1:,.2f}  (ожид bukh ≈ {BUKH_EXP:,.2f})")
print(f"Форма2            = {sum_net_f2:,.2f}")
print(f"без ФормаPL       = {sum_net_none:,.2f}  (ожид 0)")
print(f"Σ(Ф1+Ф2)          = {sum_net_f1+sum_net_f2:,.2f}  (ожид total ≈ {TOTAL_EXP:,.2f})")
print(f"\n=== ВзносыВсего (ЕСВ) ===")
print(f"Форма1            = {sum_esv_f1:,.2f}  (ожид 759 187.33)")
print(f"НЕ Форма1         = {sum_esv_other:,.2f}  (ожид 0)")
print(f"\n=== НачисленныйНДФЛ ===")
print(f"Форма1            = {ndfl_f1:,.2f}  (ожид 812 314.37)")
print(f"НЕ Форма1         = {ndfl_other:,.2f}  (ожид 0)")

# вердикты
ok_total = abs((sum_net_f1+sum_net_f2) - TOTAL_EXP) < 1.0
ok_bukh  = abs(sum_net_f1 - BUKH_EXP) < 50.0
ok_none  = abs(sum_net_none) < 0.01
ok_esv   = abs(sum_esv_other) < 0.01
ok_ndfl  = abs(ndfl_other) < 0.01
print(f"\n=== ВЕРДИКТ ===")
print(f"total неизменна : {'OK' if ok_total else 'FAIL'}")
print(f"Σ Ф1 = bukh     : {'OK' if ok_bukh else 'FAIL'} (Δ={sum_net_f1-BUKH_EXP:,.2f})")
print(f"нет без ФормаPL  : {'OK' if ok_none else 'FAIL'}")
print(f"ЕСВ все Ф1       : {'OK' if ok_esv else 'FAIL'}")
print(f"НДФЛ все Ф1      : {'OK' if ok_ndfl else 'FAIL'}")
print("ИТОГ: " + ("ВСЕ OK" if all([ok_total,ok_bukh,ok_none,ok_esv,ok_ndfl]) else "ЕСТЬ ПРОБЛЕМЫ"))

erp = None
print("\nDone (не записан).")
