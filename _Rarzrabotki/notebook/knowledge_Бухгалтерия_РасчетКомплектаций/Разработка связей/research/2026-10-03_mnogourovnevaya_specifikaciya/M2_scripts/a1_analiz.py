# M2: анализ q1_out.json (без подключения к базе): единицы строк СС против единиц карточек
import json, sys, collections, re, os
sys.stdout.reconfigure(encoding='utf-8')
H = os.path.dirname(os.path.abspath(__file__))
r = json.load(open(os.path.join(H, 'q1_out.json'), encoding='utf-8'))
SYN = {'шт.': 'шт', 'м.п.': 'м', 'п.м': 'м', 'пог.м': 'м', 'м.п': 'м', 'мп': 'м', 'м.кв.': 'м2', 'м кв': 'м2', 'м.кв': 'м2',
       'кв.м': 'м2', 'м.куб': 'м3', 'куб': 'м3', 'компл.': 'компл', 'к-т': 'компл', 'кг.': 'кг', 'т.': 'т', 'пач': 'пак'}
def k(u):
    u = re.sub(r'\s+', ' ', (u or '').strip().lower()); return SYN.get(u, u)
B = collections.defaultdict(collections.Counter)
for x in r['B']:
    B[x['КодОН']][k(x['Ед'])] += x['Карточек']
for ss in ['000000008', '000000009']:
    rows = [x for x in r['A'] if x['КодСС'] == ss]
    on = set(x['КодОН'] for x in rows if x['КодОН'])
    multi = {o for o in on if len(B.get(o, {})) > 1}
    s = sum(x['Сумма'] for x in rows if x['КодОН'] in multi); tot = sum(x['Сумма'] for x in rows)
    diff = [x for x in rows if x['КодОН'] in B and k(x['ЕдиницаСС']) != B[x['КодОН']].most_common(1)[0][0]]
    print(ss, 'строк', len(rows), 'ОН', len(on), '| ОН с карточками в 2+ базовых единицах', len(multi),
          'строк', sum(1 for x in rows if x['КодОН'] in multi), 'норма', round(s, 2), f'({100*s/tot:.1f} %)',
          '| ЕдиницаСС != модальной единице карточек ОН:', len(diff), 'строк', round(sum(x['Сумма'] for x in diff), 2),
          '| ЕдиницаСС_Оригинал заполнена:', sum(1 for x in rows if x['ЕдиницаСС_Оригинал']))
