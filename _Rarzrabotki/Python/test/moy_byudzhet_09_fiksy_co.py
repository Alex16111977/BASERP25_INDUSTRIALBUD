# -*- coding: utf-8 -*-
"""Аудит п.9: постоянные затраты ЦО «Производство» — фиксы и факт по статьям.

Фиксы: РН.А_БюджетыНаМесяц, ВидПериода = Фиксы, подразделение = корень «Производство» (норма направления).
Факт: снимок fakt_deneg.json (А_ДвиженияДенегИзКазны) по корню, 07–09.2026, без транзита и финансирования.
Вывод: data/fiksy_co.json + таблица «статья | фикс | факт 07 | 08 | 09 | среднее».
"""
import collections
from moy_byudzhet_common import *

pm, root = projects()
art = articles()
erp = conn()
t = query("""ВЫБРАТЬ Б.Подразделение КАК Подр, Б.СтатьяДвиженияДенежныхСредств КАК Статья, Б.ВидОплаты КАК Вид,
    НАЧАЛОПЕРИОДА(Б.Месяц, МЕСЯЦ) КАК Месяц, СУММА(Б.Сумма) КАК Сумма
ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Б
ГДЕ Б.ВидПериода = ЗНАЧЕНИЕ(Перечисление.А_ВидыПериодовБюджетирования.Фиксы)
  И Б.Подразделение В ИЕРАРХИИ (ВЫБРАТЬ П.Ссылка ИЗ Справочник.СтруктураПредприятия КАК П ГДЕ П.Код = &Корень)
СГРУППИРОВАТЬ ПО Б.Подразделение, Б.СтатьяДвиженияДенежныхСредств, Б.ВидОплаты, НАЧАЛОПЕРИОДА(Б.Месяц, МЕСЯЦ)""", Корень=ROOT_CODE)
fix = []
for i in range(t.Количество()):
    r = t.Получить(i)
    a = art.get(uid(r.Статья), {"code": "", "name": "?", "type": ""})
    fix.append({"proj": pm[uid(r.Подр)]["uid"], "projname": pm[uid(r.Подр)]["name"], "dept": pm[uid(r.Подр)]["dept"],
                "art": a["code"], "artname": a["name"], "dir": a["type"], "type": S(r.Вид), "month": ym(r.Месяц),
                "sum": float(r.Сумма)})
save("fiksy_co.json", {"stamp": snapshot_stamp(), "rows": fix})

months_fix = sorted({x["month"] for x in fix})
print("Фиксы: месяцы", months_fix, "| подразделения", sorted({x["dept"] for x in fix}))
fakt = load("fakt_deneg.json")["rows"]
исключить = ТРАНЗИТ | ФИНАНСИРОВАНИЕ | {"УТ-001181", "УТ-001195"}
F = collections.defaultdict(lambda: collections.defaultdict(float))
for r in fakt:
    if r["proj"] != root or r["art"] in исключить or r["month"] not in ("2026-07", "2026-08", "2026-09"):
        continue
    F[(r["art"], r["artname"], r["type"][:4])][r["month"]] += r["sum"]
FX = collections.defaultdict(float)
last = months_fix[-1] if months_fix else ""
for x in fix:
    if x["month"] == last and x["proj"] == root:
        FX[(x["art"], x["artname"], x["type"][:4])] += x["sum"] * (1 if x["dir"] == "Приход" else -1)
print(f"\nстатья ЦО | фикс ({last}) | факт 07 | 08 | 09 | среднее 07–09")
for k in sorted(set(F) | set(FX)):
    v = F[k]
    ср = sum(v.values()) / 3
    print(f"  {k[0]:<10} {k[1][:34]:<34} {k[2]:<4} {FX.get(k, 0):>12,.0f} | {v.get('2026-07', 0):>11,.0f} {v.get('2026-08', 0):>11,.0f} {v.get('2026-09', 0):>11,.0f} | {ср:>11,.0f}")
