import win32com.client, sys, datetime
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
b = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = b.String
per = b.NewObject("СтандартныйПериод")
per.ДатаНачала = datetime.datetime(2026, 8, 1, 12, 0, 0)
per.ДатаОкончания = datetime.datetime(2026, 10, 31, 12, 0, 0)
rep = b.Отчеты.А_ОтчетПоСписаниюПоНормам_Бухгалтерский.Создать()
tz = rep.ПолучитьДанные(per, False, True, None, True)
print("ТЗ строк", tz.Количество())
cols = [tz.Колонки.Получить(i).Имя for i in range(tz.Колонки.Количество())]
print("колонки:", cols)

# детали по септику — сначала, до свёртки
print("\n=== септик (IRC) детали")
for i in range(tz.Количество()):
    r = tz.Получить(i)
    on = S(r.ОбщееНазвание)
    if ("септ" in on.lower() or "комлп" in on.lower()):
        dr = S(r.Подразделение)
        if "IR" not in dr.upper().replace("С", "C"):
            continue
        print(" ", dr, "|", S(r.Домик), "|", on, "|", S(r.Аналитика), "|", S(r.ВидДокумента), "| план", r.ПланКол,
              "факт", r.ФактКол, "лежить", r.ОстатокОбъектаКол, "закуп", r.ОстатокЗакупкиКол, "| док", S(r.Документ))

# свёртка по объект+ОН
t2 = tz.Скопировать()
t2.Свернуть("Подразделение, ОбщееНазвание",
            "ПланКол, ПланГрн, ФактКол, ПотребностьПозицииКол, ПотребностьПозицииГрн, ПереместитьИзЗакупкиКол, ДокупитьКол, НеПеремещеноКол, ПерекинутьМеждуДомикамиКол")
agg = []
for i in range(t2.Количество()):
    r = t2.Получить(i)
    agg.append(dict(dr=S(r.Подразделение), on=S(r.ОбщееНазвание), plan=r.ПланКол, plang=r.ПланГрн, fact=r.ФактКол,
                    need=r.ПотребностьПозицииКол, needg=r.ПотребностьПозицииГрн, move=r.ПереместитьИзЗакупкиКол,
                    buy=r.ДокупитьКол, notmoved=r.НеПеремещеноКол, swap=r.ПерекинутьМеждуДомикамиКол))
by = defaultdict(list)
for a in agg:
    by[a["dr"]].append(a)
print("\n=== Итог по объектам (период 01.08–31.10.2026, все остатки закупки)")
for dr, rows in by.items():
    buy = [r for r in rows if r["buy"] > 0]
    buy_grn = sum(r["buy"] * (r["plang"] / r["plan"]) for r in buy if r["plan"] > 0)
    buy_noprice = [r for r in buy if not r["plan"] > 0]
    print(f"{dr}: ОН {len(rows)}; докупити ОН {len(buy)} (≈{buy_grn:,.2f} грн за план. ціною з ПДВ, без ціни {len(buy_noprice)}); "
          f"перемістити ОН {sum(1 for r in rows if r['move']>0)}; не переміщено ОН {sum(1 for r in rows if r['notmoved']>0)} / кіл {sum(r['notmoved'] for r in rows):,.3f}; "
          f"перекинути ОН {sum(1 for r in rows if r['swap']>0)}; без ОН рядків {sum(1 for r in rows if not r['on'])}")
for dr, rows in by.items():
    if "30 м" not in dr and "15 м" not in dr:
        continue
    print(f"\n--- {dr}: ТОП докупити")
    for r in sorted([r for r in rows if r["buy"] > 0], key=lambda r: -(r["buy"] * (r["plang"] / r["plan"]) if r["plan"] > 0 else 0))[:25]:
        pr = r["plang"] / r["plan"] if r["plan"] > 0 else 0
        print(f"   {r['on']!r}: план {r['plan']:.3f} факт {r['fact']:.3f} треба {r['need']:.3f} перемістити {r['move']:.3f} ДОКУПИТИ {r['buy']:.3f} (≈{r['buy']*pr:,.2f})")
    nm = [r for r in rows if r["notmoved"] > 0]
    print(f"--- {dr}: не переміщено (топ)", [(r["on"], round(r["notmoved"], 3)) for r in sorted(nm, key=lambda r: -r["notmoved"])[:15]])

# позиции СС внутри ОН: сколько строк плана (Аналитика) на ОН, где докупити > 0
print("\n=== ОН з докупити і кількома позиціями плану (Аналітика) — IRC")
pos = defaultdict(set)
for i in range(tz.Количество()):
    r = tz.Получить(i)
    if S(r.ВидДокумента).startswith("0."):
        pos[(S(r.Подразделение), S(r.ОбщееНазвание))].add(S(r.Аналитика))
for dr, rows in by.items():
    if "IR" not in dr.upper().replace("С", "C"):
        continue
    for r in rows:
        p = pos.get((dr, r["on"]), set())
        if r["buy"] > 0 and len(p) > 1:
            print(f"  {dr} | {r['on']} | докупити {r['buy']:.3f} | позиції: {sorted(p)}")
