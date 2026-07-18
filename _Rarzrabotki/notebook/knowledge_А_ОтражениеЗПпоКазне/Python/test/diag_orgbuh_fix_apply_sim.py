# -*- coding: utf-8 -*-
# §5 полная симуляция АЛГОРИТМА фикса (Подход B + guard) на реальных НЗ-строках:
# для каждого multi-org ФЛ: распределить per-ФЛ Форма1-итог по юрлицам пропорц. НачисленияБух GROSS;
# guard: если текущее per-org уже совпадает с GROSS-долей (в копейку) — НЕ трогаем (балансные).
# Доказывает: 6 «сломанных» приводятся к GROSS-доле, 5 балансных не меняются. Σ по ФЛ неизменна.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
def run(t):
    q = erp.NewObject("Запрос"); q.Текст = t; return q.Выполнить().Выгрузить()
P = lambda a: f"ГОД({a})=2026 И МЕСЯЦ({a})=2"

# НачисленияБух GROSS per (ФЛ,Орг) — веса
buh = {}
tz = run(f"""ВЫБРАТЬ НБ.ФизЛицо.Наименование КАК ФИО, НБ.Организация.Наименование КАК Орг, СУММА(НБ.Сумма) КАК С
ИЗ Документ.А_ОтражениеЗПпоКазне.НачисленияБух КАК НБ ГДЕ {P('НБ.Ссылка.Дата')}
СГРУППИРОВАТЬ ПО НБ.ФизЛицо.Наименование, НБ.Организация.Наименование""")
for i in range(tz.Количество()):
    r=tz.Получить(i); buh.setdefault(r.ФИО,{})[r.Орг]=round(float(r.С),2)

# Текущие НЗ Форма1 per (ФЛ,ОргБух): Сумма+ВзносыВсего (то, что разносим)
cur = {}
tz = run(f"""ВЫБРАТЬ НЗ.ФизическоеЛицо.Наименование КАК ФИО, НЗ.ОрганизацияБухгалтерия.Наименование КАК Орг,
  СУММА(НЗ.Сумма + НЗ.ВзносыВсего) КАК С
ИЗ Документ.А_ОтражениеЗПпоКазне.НачисленнаяЗарплатаИВзносыПоФизлицам КАК НЗ
ГДЕ {P('НЗ.Ссылка.Дата')} И НЗ.ФормаPL=ЗНАЧЕНИЕ(Перечисление.А_ФормыPL.Форма1)
СГРУППИРОВАТЬ ПО НЗ.ФизическоеЛицо.Наименование, НЗ.ОрганизацияБухгалтерия.Наименование""")
for i in range(tz.Количество()):
    r=tz.Получить(i); cur.setdefault(r.ФИО,{})[r.Орг or "(пусто)"]=round(float(r.С),2)

def split(total, weights):  # как _ОтрЗП_РазбитьСтрокуПоОрг: остаток в последнее юрлицо
    items=sorted(weights.items()); tw=round(sum(weights.values()),2); out={}; rem=total
    for k,(o,w) in enumerate(items):
        if k==len(items)-1: out[o]=round(rem,2)
        else: part=round(total*w/tw,2); out[o]=part; rem=round(rem-part,2)
    return out

multi=sorted([f for f in buh if len(buh[f])>1])
print(f"=== Симуляция применения фикса (guard) по {len(multi)} ФЛ ===")
changed=0; untouched=0; bad=0
for f in multi:
    tot=round(sum(cur.get(f,{}).values()),2)        # per-ФЛ Форма1 итог (инвариант)
    tgt=split(tot, buh[f])                            # целевое разнесение по GROSS-доле
    # guard: текущее уже == target?
    orgs=sorted(set(buh[f])|set(cur.get(f,{})))
    match=all(abs(cur.get(f,{}).get(o,0)-tgt.get(o,0))<=0.02 for o in orgs)
    sumtgt=round(sum(tgt.values()),2)
    inv = "Σ OK" if abs(sumtgt-tot)<=0.02 else "Σ СЛОМ!"
    if abs(sumtgt-tot)>0.02: bad+=1
    if match:
        untouched+=1; verdict="БАЛАНС → guard пропускает (без изменений)"
    else:
        changed+=1; verdict="РЕ-СПЛИТ"
    print(f"\n{f}  (итог Форма1={tot:,.2f}, {inv})  → {verdict}")
    for o in orgs:
        print(f"    {o[:26]:27} GROSS={buh[f].get(o,0):11,.2f}  тек={cur.get(f,{}).get(o,0):11,.2f}  →  target={tgt.get(o,0):11,.2f}")
print(f"\nИТОГ: ре-сплит={changed}  без изменений(guard)={untouched}  Σ-инвариант сломан={bad}")
print("Ожидаемо: ре-сплит=6 (ДЖИ ТРИ), без изменений=5 (ІНДЕПТ), Σ сломан=0")
erp=None
