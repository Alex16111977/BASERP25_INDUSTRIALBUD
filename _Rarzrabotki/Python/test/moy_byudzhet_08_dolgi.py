# -*- coding: utf-8 -*-
"""Аудит п.8: долги и авансы на 01.10.2026 — сверочные ограничители бюджета.

ЕРП: РН.РасчетыСКлиентамиПоСрокам / РасчетыСПоставщикамиПоСрокам, остатки на конец 30.09.2026 по проекту
(ОбъектРасчетов.Подразделение -> проект) и партнёру: Долг (нам должны / мы должны) и Предоплата (аванс).
Книга экономиста «Долги Коррект.xlsx» (лист «Долги»): колонка I «ДОЛГ» — долг поставщику на 01.10,
J/K/L — его план оплат по месяцам (включая новые заказы). Вывод: data/dolgi.json.
"""
import collections, openpyxl
from moy_byudzhet_common import *

pm, root = projects()
erp = conn()
гр = dt(2026, 10, 1)  # остатки на начало 01.10 = конец 30.09
out = {"stamp": snapshot_stamp(), "erp": [], "kniga": []}
for вид, рег in (("клиенты", "РасчетыСКлиентамиПоСрокам"), ("поставщики", "РасчетыСПоставщикамиПоСрокам")):
    t = query(f"""ВЫБРАТЬ Р.ОбъектРасчетов.Подразделение КАК Подр, Р.АналитикаУчетаПоПартнерам.Партнер.Наименование КАК Партнер,
        СУММА(Р.ДолгОстаток) КАК Долг, СУММА(Р.ПредоплатаОстаток) КАК Аванс
    ИЗ РегистрНакопления.{рег}.Остатки(&Гр, ) КАК Р
    СГРУППИРОВАТЬ ПО Р.ОбъектРасчетов.Подразделение, Р.АналитикаУчетаПоПартнерам.Партнер.Наименование""", Гр=гр)
    for i in range(t.Количество()):
        r = t.Получить(i)
        pu = uid(r.Подр) if erp.ЗначениеЗаполнено(r.Подр) else ""
        if pu not in pm:
            continue  # не производство
        out["erp"].append({"kind": вид, "proj": pm[pu]["uid"], "projname": pm[pu]["name"], "partner": S(r.Партнер).strip(),
                           "debt": float(r.Долг), "advance": float(r.Аванс)})

wb = openpyxl.load_workbook(os.path.join(BASE_DIR, "Долги Коррект.xlsx"), data_only=True)
ws = wb["Долги"]
for r in range(12, ws.max_row + 1):
    цфо = ws.cell(r, 1).value
    if not цфо:
        continue
    def num(c):
        v = ws.cell(r, c).value
        return float(v) if isinstance(v, (int, float)) else 0.0
    out["kniga"].append({"cfo": str(цфо).strip(), "article": str(ws.cell(r, 5).value or "").replace("\xa0", " ").strip(),
                         "kontr": str(ws.cell(r, 14).value or "").replace("\xa0", " ").strip(),
                         "name": str(ws.cell(r, 15).value or "").replace("\xa0", " ").strip()[:120],
                         "debt": num(9), "plan": [num(10), num(11), num(12)], "row": r})
save("dolgi.json", out)

agg = collections.defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
for x in out["erp"]:
    a = agg[x["projname"]]
    if x["kind"] == "клиенты":
        a[0] += x["debt"]; a[1] += x["advance"]
    else:
        a[2] += x["debt"]; a[3] += x["advance"]
print("ЕРП на 01.10.2026: проект | нам должны | аванс получен | мы должны | аванс выдан")
for p, a in sorted(agg.items()):
    if any(abs(v) > 1 for v in a):
        print(f"  {p[:30]:<30} {fmt(a[0]):>15} {fmt(a[1]):>15} {fmt(a[2]):>15} {fmt(a[3]):>15}")
kn = collections.defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
for x in out["kniga"]:
    k = kn[x["cfo"]]
    k[0] += x["debt"]
    for i in range(3):
        k[i + 1] += x["plan"][i]
print("\nКнига «Долги Коррект»: ЦФО | долг на 01.10 | план окт | ноя | дек")
for c, k in sorted(kn.items()):
    print(f"  {c[:30]:<30} {fmt(k[0]):>15} {fmt(k[1]):>15} {fmt(k[2]):>12} {fmt(k[3]):>12}")
