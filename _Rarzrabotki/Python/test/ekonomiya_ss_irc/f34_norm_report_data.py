# -*- coding: utf-8 -*-
"""Ф1/Ф2. Детальные строки отчёта «Списання за нормами і понад норму (бухоблік)» (BuhBud, встроенный
Отчеты.А_ОтчетПоСписаниюПоНормам_Бухгалтерский) через экспортную ПолучитьДанные — ровно его числа.
Параметры как в «Списання IRC.xlsx»: период 01.08–31.10.2026, ГП=Ложь, факт з ПДВ=Истина, склады — автоподбор,
все остатки закупки = Истина (чтобы видеть и неразложенное). Сверка итогов по домикам с xlsx. Только чтение."""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(__file__))
from common import *
import pywintypes

def dt(y, m, d, h=0, mi=0, s=0):
    return pywintypes.Time(datetime.datetime(y, m, d, h, mi, s) + datetime.timedelta(hours=3))   # COM −3 ч

Bh = buh()
rep = Bh.Отчеты.А_ОтчетПоСписаниюПоНормам_Бухгалтерский.Создать()
per = Bh.NewObject("СтандартныйПериод")
per.ДатаНачала = dt(2026, 8, 1)
per.ДатаОкончания = dt(2026, 10, 31, 12)
ALL = (sys.argv[1] if len(sys.argv) > 1 else "1") == "1"
tz = rep.ПолучитьДанные(per, False, True, None, ALL)
cols = [tz.Колонки.Получить(i).Имя for i in range(tz.Колонки.Количество())]
print("колонки:", cols)
R = rows(Bh, tz, cols)
dump_json({"cols": cols, "rows": R}, "norm_report_rows.json" if ALL else "norm_report_rows_default.json")
print("строк:", len(R))
