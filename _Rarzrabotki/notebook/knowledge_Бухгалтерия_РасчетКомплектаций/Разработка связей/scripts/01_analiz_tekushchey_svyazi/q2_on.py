# -*- coding: utf-8 -*-
# ТОЛЬКО ЧТЕНИЕ. BuhBud: согласованность ОН между картами СС, дубли ОН, карточки на ОН.
import win32com.client, sys, collections, re, json
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
B = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = B.String
def run(t):
    q = B.NewObject("Запрос"); q.Text = t
    try:
        return q.Execute().Выгрузить()
    except Exception as e:
        print("FAIL:", e.excepinfo[2] if getattr(e, 'excepinfo', None) else e); raise
# 1. одна и та же НоменклатураСС в 008 и 009 с разным ОН
t = run("""ВЫБРАТЬ СС.Код КАК Код, ВЫРАЗИТЬ(К.НоменклатураСС КАК СТРОКА(300)) КАК НомСС, К.ОбщееНазвание.Наименование КАК ОН, К.Сумма КАК Сум
ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.СтруктураСебестоимости КАК СС ПО СС.Ссылка = К.Ссылка
ГДЕ СС.Код В ("000000008", "000000009")""")
pos = collections.defaultdict(dict)
for i in range(t.Количество()):
    r = t.Получить(i)
    pos[" ".join(S(r.НомСС).split()).casefold()][S(r.Код).strip()] = (S(r.ОН), float(r.Сум))
both = {k: v for k, v in pos.items() if len(v) == 2}
diff = {k: v for k, v in both.items() if v["000000008"][0] != v["000000009"][0]}
print("позиций в обеих картах:", len(both), "; с разным ОН:", len(diff))
for k, v in diff.items():
    print(f"   «{k}»: 15 м «{v['000000008'][0]}» / 30 м «{v['000000009'][0]}» ({v['000000008'][1]:,.2f} / {v['000000009'][1]:,.2f})")
# 2. дубли ОН по нормализованному имени + «септик»
t = run("""ВЫБРАТЬ О.Ссылка КАК Ссылка, О.Наименование КАК Имя, О.Код КАК Код, О.ПометкаУдаления КАК ПУ,
    О.Единица.Наименование КАК Ед, ВЫРАЗИТЬ(О.НазваниеДляОбъединения КАК СТРОКА(300)) КАК ДляОбъед, О.Активное КАК Акт,
    О.НеПроверятьВЗакупкахпоСС КАК НеПров
ИЗ Справочник.ОбщиеНазванияНоменклатуры КАК О""")
print("\nОН всего:", t.Количество())
def norm(s):
    s = str(s).replace("\xa0", " ").replace("’", "'").replace("ʼ", "'").replace("`", "'").casefold()
    s = s.replace("i", "і").replace("c", "с").replace("e", "е").replace("o", "о").replace("a", "а").replace("p", "р")
    return " ".join(re.sub(r"[\"'.,()\-]", " ", s).split())
grp = collections.defaultdict(list); allon = []
for i in range(t.Количество()):
    r = t.Получить(i)
    d = dict(name=S(r.Имя), code=S(r.Код).strip(), del_=bool(r.ПУ), unit=S(r.Ед), merge=S(r.ДляОбъед), act=r.Акт, noch=bool(r.НеПров))
    allon.append(d); grp[norm(d["name"])].append(d)
dups = {k: v for k, v in grp.items() if len([x for x in v if not x["del_"]]) > 1}
print("групп-дублей (после нормализации, без помеченных):", len(dups))
for k, v in list(dups.items())[:40]:
    print("   ", [(x["code"], x["name"], x["unit"]) for x in v if not x["del_"]])
print("с заполненным НазваниеДляОбъединения:", sum(1 for d in allon if d["merge"].strip()), [ (d['name'], d['merge']) for d in allon if d['merge'].strip()][:15])
print("НеПроверятьВЗакупкахпоСС:", sum(1 for d in allon if d["noch"]))
# опечатки: расстояние Левенштейна 1-2 (перестановка) среди непомеченных
def lev(a, b):
    if abs(len(a) - len(b)) > 2: return 9
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]
live = [d for d in allon if not d["del_"]]
keys = sorted({norm(d["name"]) for d in live})
near = []
for i, a in enumerate(keys):
    for b in keys[i + 1:]:
        if len(a) >= 6 and a[:2] == b[:2] and lev(a, b) <= 2 and not re.search(r"\d", a + b):
            near.append((a, b))
print("пары ОН с расстоянием ≤2 (без цифр):", len(near)); print("   ", near[:60])
print("септик:", [(d['code'], d['name'], d['del_'], d['unit']) for d in allon if 'септик' in d['name'].casefold()])
