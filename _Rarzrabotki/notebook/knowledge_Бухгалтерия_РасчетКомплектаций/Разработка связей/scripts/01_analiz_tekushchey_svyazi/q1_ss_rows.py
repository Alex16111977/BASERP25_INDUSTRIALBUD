# -*- coding: utf-8 -*-
# ТОЛЬКО ЧТЕНИЕ. BuhBud: строки СС IRC (000000008/000000009) и их связь с карточками.
import win32com.client, sys, collections
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
t = run("""ВЫБРАТЬ
    СС.Код КАК КодСС, СС.Наименование КАК ИмяСС, К.НомерСтроки КАК НС,
    К.Этап КАК Этап, ВЫРАЗИТЬ(К.НоменклатураСС КАК СТРОКА(300)) КАК НомСС,
    К.ОбщееНазвание КАК ОН, К.ОбщееНазвание.Наименование КАК ОНИмя,
    К.Номенклатура КАК Ном, К.Номенклатура.Наименование КАК НомИмя,
    К.Номенклатура.ОбщееНазваниеНоменклатуры КАК ОННом,
    К.ОбщееНазвание.ОсновнаяНоменклатура КАК ОснНом,
    К.Единица КАК Ед, К.ЕдиницаСС КАК ЕдСС, К.ЕдиницаСС_Оригинал КАК ЕдОриг, К.НеСовпадает КАК НеСовп,
    К.ДобавленоВручную КАК Вручную, К.Количество КАК Кол, К.Сумма КАК Сум
ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.СтруктураСебестоимости КАК СС ПО СС.Ссылка = К.Ссылка
ГДЕ СС.Код В ("000000008", "000000009")""")
print("строк", t.Количество())
st = collections.Counter(); money = collections.defaultdict(float)
rows = []
for i in range(t.Количество()):
    r = t.Получить(i)
    d = dict(code=S(r.КодСС).strip(), ss=S(r.ИмяСС), n=int(r.НС), etap=S(r.Этап), nomss=S(r.НомСС), on=S(r.ОНИмя),
             nom=S(r.НомИмя), on_nom=S(r.ОННом), has_nom=not r.Ном.Пустая() if r.Ном is not None else False,
             nom_is_main=(S(r.Ном) == S(r.ОснНом)) and S(r.Ном) != "", on_mismatch=(S(r.ОННом) != S(r.ОН)) and S(r.Ном) != "",
             ed=S(r.Ед), edss=S(r.ЕдСС), edorig=S(r.ЕдОриг), nesov=bool(r.НеСовп), man=bool(r.Вручную), q=float(r.Кол), s=float(r.Сум))
    rows.append(d)
for code in ("000000008", "000000009"):
    R = [d for d in rows if d["code"] == code]
    print(f"\n== {code} {R[0]['ss']}: строк {len(R)}, сумма {sum(d['s'] for d in R):,.2f}")
    print("  с Номенклатурой:", sum(d['has_nom'] for d in R), "; = ОсновнаяНоменклатура ОН:", sum(d['nom_is_main'] for d in R),
          "; ОН строки ≠ ОН карточки:", sum(d['on_mismatch'] for d in R), f"({sum(d['s'] for d in R if d['on_mismatch']):,.2f} грн)")
    print("  НеСовпадает (ед.):", sum(d['nesov'] for d in R), f"({sum(d['s'] for d in R if d['nesov']):,.2f} грн); ДобавленоВручную:", sum(d['man'] for d in R))
    ons = collections.Counter(d['on'] for d in R)
    print("  уникальных ОН:", len(ons), "; ОН с ≥2 строками СС:", sum(1 for v in ons.values() if v > 1),
          [(k, v) for k, v in ons.most_common(12) if v > 1])
    noms = collections.Counter(d['nomss'].strip().lower() for d in R)
    print("  уникальных НоменклатураСС:", len(noms), "; повтор НомСС:", [(k, v) for k, v in noms.items() if v > 1][:10])
    print("  ОН строки ≠ ОН карточки — примеры:")
    for d in [d for d in R if d['on_mismatch']][:25]:
        print(f"    #{d['n']} СС «{d['nomss'][:45]}» ОН «{d['on']}» → карточка «{d['nom'][:40]}» ОН карточки «{d['on_nom']}» {d['s']:,.2f}")
    print("  нет Номенклатуры:", [(d['n'], d['nomss'][:40], d['on'], d['s']) for d in R if not d['has_nom']][:15])
    print("  НеСовпадает:", [(d['n'], d['nomss'][:35], d['on'], d['edorig'], d['edss'], d['ed'], d['s']) for d in R if d['nesov']][:40])
