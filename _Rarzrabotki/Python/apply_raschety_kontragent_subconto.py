# -*- coding: utf-8 -*-
"""Шаг 2 — перенос верифицированного запроса в BSL: добавляет субконто
Контрагент/Партнер/Договор в функцию Свод_РасчетыСПартнерами
(Documents/А_ФинРез_Баланс/Ext/ObjectModule.bsl). ТОЛЬКО тело функции.

Дефенсивно: работает в пределах [Функция Свод_РасчетыСПартнерами .. КонецФункции];
вставки определяются по ТОЧНОМУ содержимому строки (после '|' и пробелов);
проверяет точные счётчики (a=1,b=3,c=4,d=1,e=1) — иначе abort без записи.

КРИТИЧНО для (c) СГРУППИРОВАТЬ ПО: исходный последний элемент
'Р.ОбъектРасчетов' БЕЗ запятой → нужно добавить запятую к нему, а ПОСЛЕДНИЙ
вставляемый элемент группировки — БЕЗ запятой (он становится последним).
SELECT-ветки (a,b), плуг (d), финал-детали (e): якорь уже с запятой,
вставки с запятыми — без изменений."""
import io, os, sys, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

WT = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\naughty-neumann-691338\Documents\А_ФинРез_Баланс\Ext\ObjectModule.bsl"
MAIN = r"C:\Configuration_downloads\BASERP25\Documents\А_ФинРез_Баланс\Ext\ObjectModule.bsl"

raw = open(WT, "rb").read()
bom = raw.startswith(b"\xef\xbb\xbf")
text = raw.decode("utf-8-sig")
nl = "\r\n" if "\r\n" in text else "\n"
lines = text.split(nl)


def prefix(line):
    return line[:len(line) - len(line.lstrip(" \t|"))]


def content(line):
    return line.lstrip(" \t|").rstrip()


fs = next(i for i, l in enumerate(lines)
          if "Функция Свод_РасчетыСПартнерами(" in l)
fe = next(i for i in range(fs + 1, len(lines))
          if lines[i].strip() == "КонецФункции")
print(f"Функция Свод_РасчетыСПартнерами: строки {fs+1}..{fe+1}")

К = "ЕСТЬNULL(ВЫРАЗИТЬ(АП.Контрагент КАК Справочник.Контрагенты), ЗНАЧЕНИЕ(Справочник.Контрагенты.ПустаяСсылка))"
Д = "ЕСТЬNULL(ВЫРАЗИТЬ(АП.Договор КАК Справочник.ДоговорыКонтрагентов), ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка))"
ВЕТКА_АЛИАС = ["АП.Партнер КАК Партнер,", К + " КАК Контрагент,", Д + " КАК Договор,"]
ВЕТКА_ПОЗ = ["АП.Партнер,", К + ",", Д + ","]                  # SELECT 2-4 (есть СУММА после)
ВЕТКА_ГРУПП = ["АП.Партнер,", К + ",", Д]                       # GROUP BY: последний БЕЗ запятой
ПЛУГ = ["ЗНАЧЕНИЕ(Справочник.Партнеры.ПустаяСсылка),",
        "ЗНАЧЕНИЕ(Справочник.Контрагенты.ПустаяСсылка),",
        "ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка),"]
ДЕТАЛЬ = ["втРасч.Партнер КАК Партнер,", "втРасч.Контрагент КАК Контрагент,",
          "втРасч.Договор КАК Договор,"]

out = lines[:fs + 1]
ca = cb = cc = cd = ce = 0
i = fs + 1
while i <= fe:
    ln = lines[i]
    c = content(ln)
    pf = prefix(ln)
    if c == "Р.ОбъектРасчетов КАК ОбъектРасчетов,":              # (a) ветка1 SELECT
        out.append(ln); out += [pf + x for x in ВЕТКА_АЛИАС]; ca += 1
    elif c == "Р.ОбъектРасчетов," and i + 1 <= fe and "СУММА(" in lines[i + 1]:
        out.append(ln); out += [pf + x for x in ВЕТКА_ПОЗ]; cb += 1  # (b) SELECT 2-4
    elif c == "Р.ОбъектРасчетов":                                # (c) GROUP BY x4
        out.append(pf + "Р.ОбъектРасчетов,")
        out += [pf + x for x in ВЕТКА_ГРУПП]; cc += 1
    elif c == "ЗНАЧЕНИЕ(Справочник.ОбъектыРасчетов.ПустаяСсылка),":  # (d) плуг
        out.append(ln); out += [pf + x for x in ПЛУГ]; cd += 1
    elif c == "втРасч.ОбъектРасчетов КАК ОбъектРасчетов,":       # (e) финал детали
        out.append(ln); out += [pf + x for x in ДЕТАЛЬ]; ce += 1
    else:
        out.append(ln)
    i += 1
out += lines[fe + 1:]

print(f"вставки: a(ветка1 SELECT)={ca} b(ветки2-4 SELECT)={cb} "
      f"c(GROUP BY +запятая, посл.без запятой)={cc} d(плуг)={cd} e(финал детали)={ce}")
assert (ca, cb, cc, cd, ce) == (1, 3, 4, 1, 1), \
    f"FAIL: неожиданные счётчики {(ca,cb,cc,cd,ce)} — запись отменена"

new = nl.join(out)
data = ("﻿" + new if bom else new).encode("utf-8")
shutil.copyfile(WT, WT + ".bak")
open(WT, "wb").write(data)
os.makedirs(os.path.dirname(MAIN), exist_ok=True)
open(MAIN, "wb").write(data)
print(f"OK: записано worktree (+{len(out)-len(lines)} строк), .bak создан, "
      f"скопировано в основную конфигурацию")
