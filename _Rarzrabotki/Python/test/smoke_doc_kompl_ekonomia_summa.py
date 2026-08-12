# -*- coding: utf-8 -*-
"""Smoke формулы ЕкономіяСума (= Економія × НормаСум/Норма) после пересчёта.

Scratch-док с маркером SMOKE_EKON_SUM_v1 (get-or-create, НЕ удаляется), спецификация
и склад дома №1. Инварианты (живые данные дрейфуют — абсолюты не проверяем):
 A. ЕкономіяСума > 0 <=> Экономия > 0 (по каждому эталону);
 B. ЕкономіяСума == Окр(Экономия × НормаСумма/Норма, 2) ± 0.02;
 C. построчно ВНорме + ПонадНорму == Остаток;
 D. есть >= 1 эталон с Экономия > 0 и ЕкономіяСума > 0.
"""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

MARKER = "SMOKE_EKON_SUM_v1"

v8 = win32com.client.Dispatch("V83.COMConnector")
try:
    erp = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
except Exception:
    erp = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = erp.String

FAILS = []
def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f" | {detail}" if detail else ""))
    if not cond:
        FAILS.append(name)

# эталонный док №3 — источник спецификации и склада
q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д
ГДЕ Д.Номер = "00000000003" И НЕ Д.ПометкаУдаления"""
doc3 = q.Execute().Выгрузить().Получить(0).Ссылка.ПолучитьОбъект()

# get-or-create scratch-дока по маркеру (Комментарий — неограниченная строка: ВЫРАЗИТЬ)
q2 = erp.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д
ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления"""
q2.SetParameter("М", MARKER)
r2 = q2.Execute().Выгрузить()
if r2.Количество() > 0:
    doc = r2.Получить(0).Ссылка.ПолучитьОбъект()
    print(f"scratch-док найден: №{S(doc.Номер)}")
else:
    import datetime
    doc = erp.Документы.РасчетКомплектаций.СоздатьДокумент()
    # глобальные ТекущаяДата/CurrentDate не резолвятся COM-прокси — дата из Python
    doc.Дата = datetime.datetime.now().replace(microsecond=0)
    doc.Заполнить(None)
    doc.Комментарий = MARKER
    print("scratch-док создаётся")

doc.Спецификация = doc3.Спецификация
doc.Период = doc3.Период
doc.Организация = doc3.Организация
doc.СкладыОстатков.Очистить()
for i in range(doc3.СкладыОстатков.Количество()):
    doc.СкладыОстатков.Добавить().Склад = doc3.СкладыОстатков.Получить(i).Склад
doc.СчетаОстатков.Очистить()
for i in range(doc3.СчетаОстатков.Количество()):
    doc.СчетаОстатков.Добавить().Счет = doc3.СчетаОстатков.Получить(i).Счет
doc.СчетаМалоценки.Очистить()
for i in range(doc3.СчетаМалоценки.Количество()):
    doc.СчетаМалоценки.Добавить().Счет = doc3.СчетаМалоценки.Получить(i).Счет

doc.РассчитатьАнализ()
doc.Записать()

tch = doc.ТабличнаяЧастьОстатков
print(f"строк после пересчёта: {tch.Количество()}")
if tch.Количество() == 0:
    print("Остатков на дату нет (после списаний) — сдвигаю период на день раньше")
    per = doc.Период
    import datetime
    doc.Период = per - datetime.timedelta(days=1)
    doc.РассчитатьАнализ()
    doc.Записать()
    tch = doc.ТабличнаяЧастьОстатков
    print(f"строк после пересчёта (период -1 день): {tch.Количество()}")

et = {}
bad_c = 0
for i in range(tch.Количество()):
    s = tch.Получить(i)
    d = et.setdefault(S(s.ОбщееНазвание),
                      {"norma": 0.0, "nsum": 0.0, "ek": 0.0, "eksum": 0.0})
    d["norma"] += s.Норма; d["nsum"] += s.НормаСумма
    d["ek"] += s.Экономия; d["eksum"] += s.ЕкономіяСума
    if abs(s.ВНорме + s.ПонадНорму - s.Остаток) > 0.001:
        bad_c += 1
check("C: построчный баланс", bad_c == 0, f"нарушений {bad_c}")

bad_a, bad_b, pos = 0, 0, 0
for name, d in et.items():
    if (d["ek"] > 0.0005) != (d["eksum"] > 0.005):
        bad_a += 1
        print(f"  A-нарушение {name!r}: ек={d['ek']} сум={d['eksum']}")
    if d["norma"] > 0:
        expect = round(d["ek"] * d["nsum"] / d["norma"], 2)
        if abs(d["eksum"] - expect) > 0.02:
            bad_b += 1
            print(f"  B-нарушение {name!r}: сум={d['eksum']} ожид={expect}")
    if d["ek"] > 0 and d["eksum"] > 0:
        pos += 1
check("A: ЕкономіяСума>0 <=> Економія>0", bad_a == 0, f"нарушений {bad_a}")
check("B: сумма = кол × нормовая цена", bad_b == 0, f"нарушений {bad_b}")
check("D: есть эталоны с экономией", pos >= 1, f"pos={pos}")
tot = sum(d["eksum"] for d in et.values())
print(f"Σ ЕкономіяСума по scratch-доку: {tot:.2f} грн (справочно; по ТЧ дока №3 старая была 109072.97)")

print(f"\n{'='*50}\nИТОГ: {'ALL PASS' if not FAILS else 'FAILS: ' + ', '.join(FAILS)}")
sys.exit(0 if not FAILS else 1)
