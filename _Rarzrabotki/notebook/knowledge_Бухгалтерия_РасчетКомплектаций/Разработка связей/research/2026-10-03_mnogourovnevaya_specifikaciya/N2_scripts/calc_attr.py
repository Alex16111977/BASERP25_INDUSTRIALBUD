import json,sys
sys.stdout.reconfigure(encoding='utf-8')
rows=json.load(open('kn_rows.json',encoding='utf-8'))
D=json.load(open(r'C:/Configuration_downloads/BASERP25/.claude/worktrees/budget-calc-oct-dec-2026-20c998/_Rarzrabotki/notebook/knowledge_Бухгалтерия_РасчетКомплектаций/Разработка связей/data/demo_svyaz_ss_30m.json',encoding='utf-8'))
grp=[r for r in D['mapping'] if 33<=r['n']<=41]
codes={l['code'].strip() for r in grp for l in r['links']}
# ERP money (с ПДВ) per card, 30m №1 (q9)
erp={'Саморіз 3,5х35':626.45,'Саморіз 4,2*19':772.82,'Саморіз 4,8*19':586.34,'Саморіз 4,8х35':1524.04+609.61,
 'Універсальний шуруп 5,0х90 потай':856.91,'Шуруп 4,5х50':96.87,'Шуруп 4,5х70':532.57,'Шуруп 5,5х75':1780.72}
fact={}
for r in rows:
    if '30 м №1' in r['Дом'] and r['Код'].strip() in codes:
        k=r['Карт'].strip(); f=fact.setdefault(k,[0,0]); f[0]+=r['Кол']; f[1]+=r['Сум']
Q=sum(v[0] for v in fact.values()); M=sum(v[1] for v in fact.values()); ME=sum(erp.values())
N=sum(r['norm_q'] for r in grp); NU=sum(r['norm_uah'] for r in grp)
print('fact qty',Q,'buh net',round(M,2),'erp gross',round(ME,2),'norm qty',N,'norm uah',NU)
print('--- Rule1 share')
for r in grp:
    s=r['norm_q']/N
    print(r['n'],r['text'][:32],r['norm_q'],round(Q*s,1),round(ME*s,2), 'норма грн',r['norm_uah'], 'откл кол',round(Q*s-r['norm_q'],1))
print('--- Rule2 order (cards ordered by doc line order)')
# card order by first appearance in docs (doc date, line)
order=[]
for r in rows:
    if '30 м №1' in r['Дом'] and r['Код'].strip() in codes:
        k=r['Карт'].strip()
        if k not in order: order.append(k)
pool=[[k,fact[k][0],erp[k]/fact[k][0]] for k in order]
for r in grp:
    need=r['norm_q']; got=0; money=0
    while need>1e-9 and pool:
        k,q,p=pool[0]; t=min(q,need); got+=t; money+=t*p; need-=t; pool[0][1]-=t
        if pool[0][1]<=1e-9: pool.pop(0)
    print(r['n'],r['text'][:32],r['norm_q'],round(got,1),round(money,2),'откл кол',round(got-r['norm_q'],1))
print('остаток',pool)
# Rule2b: cards ordered by price desc
pool=sorted([[k,fact[k][0],erp[k]/fact[k][0]] for k in order], key=lambda x:-x[2])
tot=[]
for r in grp:
    need=r['norm_q']; got=0; money=0
    while need>1e-9 and pool:
        k,q,p=pool[0]; t=min(q,need); got+=t; money+=t*p; need-=t; pool[0][1]-=t
        if pool[0][1]<=1e-9: pool.pop(0)
    tot.append((r['n'],round(got,1),round(money,2)))
print('Rule2 price-desc',tot)
print('card order',order)
