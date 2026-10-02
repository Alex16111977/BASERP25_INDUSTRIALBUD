# -*- coding: utf-8 -*-
"""Ф0. Версия СС: построчная сверка ERP↔BuhBud и история версий карточек (ВерсииОбъектов) в обеих базах."""
import os, sys, json
sys.path.insert(0, os.path.dirname(__file__))
from common import *

se = json.load(open(os.path.join(OUT, "ss_erp.json"), encoding="utf-8"))
sb = json.load(open(os.path.join(OUT, "ss_buh.json"), encoding="utf-8"))
diff = []
for ss in sorted({r["СС"] for r in se["kom"]}):
    a = [r for r in se["kom"] if r["СС"] == ss]
    b = [r for r in sb["kom"] if r["СС"] == ss]
    for x, y in zip(a, b):
        if (norm_name(x["ОН"]) != norm_name(y["ОН"]) or abs(D(x["Кол"]) - D(y["Кол"])) > D("0.006")
                or D(x["Сумма"]) != D(y["Сумма"]) or norm_name(x["НомСС"]) != norm_name(y["НомСС"])):
            diff.append((ss, x["НС"], x["ОН"], y["ОН"], x["Кол"], y["Кол"], x["Сумма"], y["Сумма"]))
print("Построчных расхождений ERP↔BuhBud:", len(diff))
for d in diff[:30]: print("  ", d)

def versions(conn, cat):
    try:
        vt = q(conn, f"""
ВЫБРАТЬ С.Наименование КАК СС, Вер.НомерВерсии КАК N, Вер.ДатаВерсии КАК Дата, Вер.АвторВерсии КАК Автор, Вер.Комментарий КАК Комм
ИЗ РегистрСведений.ВерсииОбъектов КАК Вер
  ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.{cat} КАК С ПО Вер.Объект = С.Ссылка
ГДЕ С.Наименование ПОДОБНО "%IR%"
УПОРЯДОЧИТЬ ПО СС, N
""")
        return rows(conn, vt, ["СС", "N", "Дата", "Автор", "Комм"])
    except RuntimeError as e:
        return [str(e)]

for nm, conn, cat in (("ERP", erp(), "А_СтруктураСебестоимости"), ("BUH", buh(), "СтруктураСебестоимости")):
    v = versions(conn, cat)
    print(nm, "версий:", len(v))
    for r in v: print("  ", r)
