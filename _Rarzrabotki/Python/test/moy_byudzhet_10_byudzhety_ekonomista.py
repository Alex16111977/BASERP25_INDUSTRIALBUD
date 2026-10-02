# -*- coding: utf-8 -*-
"""Аудит п.10: ориентиры для калибровки — бюджеты экономиста в РН.А_ПредварительныйБюджетНаМесяц.

00007 «Бюджет экономиста ДДС» (сен–дек, версия 18.09) и 00008 «… (октябрь-декабрь 2026)» (принятый бюджет 01.10).
Суммы регистра положительные, знак — А_ТипПоказателя статьи. Вывод: data/byudzhety_ekonomista.json + сводка.
"""
import collections
from moy_byudzhet_common import *

pm, root = projects()
art = articles()
erp = conn()
t = query("""ВЫБРАТЬ
    Р.Сценарий.Код КАК КодСц, Р.Сценарий.Наименование КАК ИмяСц, Р.Подразделение КАК Подр,
    Р.СтатьяДвиженияДенежныхСредств КАК Статья, Р.ВидОплаты КАК Вид, Р.Месяц КАК Месяц, СУММА(Р.Сумма) КАК Сумма
ИЗ РегистрНакопления.А_ПредварительныйБюджетНаМесяц КАК Р
ГДЕ Р.Сценарий.Код В (&Коды)
СГРУППИРОВАТЬ ПО Р.Сценарий.Код, Р.Сценарий.Наименование, Р.Подразделение, Р.СтатьяДвиженияДенежныхСредств, Р.ВидОплаты, Р.Месяц""",
          Коды=["00007", "00008"])
out = []
for i in range(t.Количество()):
    r = t.Получить(i)
    pu = uid(r.Подр)
    a = art[uid(r.Статья)]
    p = pm.get(pu, {"uid": pu, "name": S(r.Подр), "dept": S(r.Подр)})
    sign = 1 if a["type"] == "Приход" else -1
    out.append({"sc": S(r.КодСц).strip(), "proj": p["uid"], "projname": p["name"], "dept": p["dept"],
                "art": a["code"], "artname": a["name"], "type": S(r.Вид), "month": ym(r.Месяц),
                "sum": float(r.Сумма), "signed": sign * float(r.Сумма), "dir": a["type"]})
save("byudzhety_ekonomista.json", {"stamp": snapshot_stamp(), "rows": out})

for sc in ("00007", "00008"):
    agg = collections.defaultdict(lambda: [0.0, 0.0])
    for r in out:
        if r["sc"] != sc:
            continue
        agg[r["month"]][0 if r["signed"] > 0 else 1] += r["signed"]
    print(f"\nСценарий {sc}: месяц | поступления | выплаты | чистый")
    for m in sorted(agg):
        print(f"  {m} | {fmt(agg[m][0]):>16} | {fmt(-agg[m][1]):>16} | {fmt(agg[m][0] + agg[m][1]):>16}")
    pr = collections.defaultdict(lambda: [0.0, 0.0])
    for r in out:
        if r["sc"] == sc and r["month"] >= "2026-10":
            pr[r["projname"]][0 if r["signed"] > 0 else 1] += r["signed"]
    print(f"  по проектам окт–дек:")
    for p in sorted(pr):
        print(f"    {p[:32]:<32} {fmt(pr[p][0]):>16} {fmt(-pr[p][1]):>16}")
