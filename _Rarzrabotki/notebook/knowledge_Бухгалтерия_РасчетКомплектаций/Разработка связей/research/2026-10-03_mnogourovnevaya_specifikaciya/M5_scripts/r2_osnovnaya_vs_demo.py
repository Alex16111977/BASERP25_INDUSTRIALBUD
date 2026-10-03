# r2: насколько ОН.ОсновнаяНоменклатура годится как кандидат строки спецификации (СС 009, 30 м).
# Сверка с прототипом связей DEMO (mapping[].links, роль «Основна») — только как эталон для оценки, не как источник.
# Только чтение BuhBud.
import sys, json, re
from h import connect, q
c = connect('buh')
DEMO = r"C:/Configuration_downloads/BASERP25/.claude/worktrees/budget-calc-oct-dec-2026-20c998/_Rarzrabotki/notebook/knowledge_Бухгалтерия_РасчетКомплектаций/Разработка связей/data/demo_svyaz_ss_30m.json"
d = json.load(open(DEMO, encoding='utf-8'))
demo = {m['n']: m for m in d['mapping']}
rows = q(c, """
ВЫБРАТЬ
	К.НомерСтроки КАК Н,
	К.НоменклатураСС КАК Поз,
	К.ОбщееНазвание КАК ОНСсылка,
	К.ОбщееНазвание.ОсновнаяНоменклатура.Код КАК ОснКод,
	К.Сумма КАК Сум
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК К
ГДЕ
	К.Ссылка.Код = &Код
УПОРЯДОЧИТЬ ПО
	Н
""", Код='000000009')
cards = q(c, """
ВЫБРАТЬ
	Н.ОбщееНазваниеНоменклатуры КАК ОНСсылка,
	Н.Код КАК Код,
	Н.Наименование КАК Имя
ИЗ
	Справочник.Номенклатура КАК Н
ГДЕ
	НЕ Н.ПометкаУдаления
	И НЕ Н.ЭтоГруппа
	И Н.ОбщееНазваниеНоменклатуры В
		(ВЫБРАТЬ РАЗЛИЧНЫЕ К.ОбщееНазвание ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К
		 ГДЕ К.Ссылка.Код = &Код)
""", Код='000000009')
by_on = {}
for r in cards:
    by_on.setdefault(r['ОНСсылка'], []).append(r)
def nums(t):
    return set(x.replace(',', '.') for x in re.findall(r'\d+(?:[.,]\d+)?', t or '') if len(x) >= 2 or x not in ('1','2','3','4','5','6','7','8','9','0'))
stat = {}
def add(k, s):
    stat.setdefault(k, [0, 0.0]); stat[k][0] += 1; stat[k][1] += s
for r in rows:
    n = r['Н']; s = r['Сум'] or 0
    lst = by_on.get(r['ОНСсылка'], [])
    dm = demo.get(n, {})
    dcode = None
    for l in dm.get('links', []):
        if l.get('role') == 'Основна':
            dcode = l['code']; break
    osn = (r['ОснКод'] or '').strip()
    ncards = len(lst)
    grp = '1 карточка' if ncards == 1 else ('2+ карточек' if ncards >= 2 else '0 карточек')
    # правило «размер в названии»: числа из текста позиции все есть в названии ровно одной карточки ОН
    pn = nums(r['Поз'])
    size_hits = [x for x in lst if pn and pn <= nums(x['Имя'])] if ncards >= 2 else []
    if ncards == 1:
        prop = lst[0]['Код'].strip(); how = 'B одна карточка'
    elif len(size_hits) == 1:
        prop = size_hits[0]['Код'].strip(); how = 'S розмір у назві (одна)'
    elif ncards >= 2:
        prop = None; how = 'C вибір бухгалтера'
    else:
        prop = None; how = 'D немає'
    add('ОН: ' + grp + (' + Основна' if osn else ''), s)
    if osn and ncards >= 2:
        add('Основна при 2+ = DEMO' if dcode and osn == dcode else 'Основна при 2+ != DEMO', s)
    add('Новые правила: ' + how, s)
    if prop:
        add('  ' + how + (' = DEMO' if prop == dcode else ' != DEMO'), s)
for k in sorted(stat):
    print(f"{k}: позиций {stat[k][0]}, сумма {stat[k][1]:,.2f}")
