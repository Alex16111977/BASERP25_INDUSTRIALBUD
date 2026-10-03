import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
d=json.load(open(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026\Мой расчет\data\prichiny_ponad_normu_irc.json",encoding='utf-8'))
S=[x for x in d if x['тч']=='СписаниеСверхНормы']
# rules by card name (номенклатура) first, then by reason key
CARD={
 'Держак труби Profil метал. L 160 90 графіт':'A','Захист від перенапруги D63':'A','Плінтус на стільницю 3м':'A',
 'Кліпса ППР 20':'A','Гнучкий шланг 1\2 ВВ довжина 30 см':'A',
}
REASON=[('Септик','C'),('Плитку тротуарну','D'),('пересорт профілів','E'),('Обидві складові поручня','D'),
 ('Плінтуса перевищено','E'),('плінтуса перевищено','E'),('Гідро- та пароізоляція','B'),('Піна та клей','E'),('кабелю понад норму','F'),
 ('Мембрани враховані','E'),('Т-подібний','D'),('ПЕ-труби','G'),('Дахову мембрану','D'),('Змішувач не','D'),('Коліно ринви','D'),
 ('групі «Гайка»','F'),('Бетонну суміш','D'),('розмір вікна','E'),('Муфта','D'),('Американку','D'),('Жалюзі','D'),
 ('з’єднувача','B'),('Заклепки','D'),('Заглушку плінтуса','D'),('групі «Болт»','F'),('Профіль для підлоги','D'),
 ('Внутрішні кутики','D'),('Гнучка труба','D'),('Шланг 0,4','G'),('клапани','G'),('трійників','G'),('Сантехнічний куточок','D'),
 ('розетки','G'),('Каналізаційну трубу','D'),('групі «Шайба»','F')]
def typ15(x):
    if x['номенклатура'] in CARD: return CARD[x['номенклатура']]
    r=x['причина'] or ''
    if not r.strip(): return 'G-МШП' if x['загальна_назва']=='мшп' else '?'
    for k,t in REASON:
        if k in r: return t
    return '?'
for x in S:
    if x['тип_дома']=='15 м': x['_t']=typ15(x)
# 30m analogy: by card, then by загальна_назва (majority type in 15m)
bycard={}; byon=defaultdict(lambda:defaultdict(float))
for x in S:
    if x['тип_дома']=='15 м':
        bycard[x['номенклатура']]=x['_t']; byon[x['загальна_назва']][x['_t']]+=x['сума']
for x in S:
    if x['тип_дома']=='30 м':
        t=bycard.get(x['номенклатура'])
        if t is None and x['загальна_назва'] in byon:
            t=max(byon[x['загальна_назва']].items(), key=lambda kv:kv[1])[0]
        x['_t']=(t or 'нет аналога')
print('unclassified 15m:',[ (x['номенклатура'],x['причина'][:60]) for x in S if x['тип_дома']=='15 м' and x['_t']=='?'])
T=defaultdict(lambda:defaultdict(lambda:[0,0.0]))
for x in S:
    key=('15' if x['тип_дома']=='15 м' else '30')+'-'+str(x['номер_дома'])
    T[x['_t']][key][0]+=1; T[x['_t']][key][1]+=x['сума']
houses=['15-1','15-2','15-3','15-4','15-5','15-6','30-1','30-2','30-3']
print('type | '+' | '.join(houses)+' | 15м итог | 30м итог | всего')
tot=defaultdict(float)
for t in sorted(T):
    v=T[t]; s15=sum(v[h][1] for h in houses[:6]); n15=sum(v[h][0] for h in houses[:6]); s30=sum(v[h][1] for h in houses[6:]); n30=sum(v[h][0] for h in houses[6:])
    print(t,'|',' | '.join('%d/%.2f'%(v[h][0],v[h][1]) if h in v else '—' for h in houses),'| %d/%.2f | %d/%.2f | %d/%.2f'%(n15,s15,n30,s30,n15+n30,s15+s30))
    for h in houses: tot[h]+=v[h][1] if h in v else 0
print('TOTAL',{h:round(tot[h],2) for h in houses}, round(sum(tot.values()),2), len(S))
# 30m detail
print('--- 30m detail')
g=defaultdict(lambda:[0,0.0,set()])
for x in S:
    if x['тип_дома']=='30 м':
        k=(x['_t'],x['загальна_назва']); g[k][0]+=1; g[k][1]+=x['сума']; g[k][2].add(x['номер_дома'])
for k,v in sorted(g.items()): print(k, v[0], round(v[1],2), sorted(v[2]))
json.dump([dict(дом=x['дом'],номенклатура=x['номенклатура'],он=x['загальна_назва'],сума=x['сума'],тип=x['_t']) for x in S], open('cls.json','w',encoding='utf-8'),ensure_ascii=False)
# per-house pересорт-word count
w=defaultdict(float)
for x in S:
    if 'пересорт' in (x['причина'] or ''): w[x['номер_дома']]+=x['сума']
print('word пересорт by house', dict(w), sum(w.values()))
