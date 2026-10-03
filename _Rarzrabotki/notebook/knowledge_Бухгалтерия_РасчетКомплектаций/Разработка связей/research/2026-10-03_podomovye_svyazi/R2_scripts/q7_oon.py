from h import *
import json, re
from collections import defaultdict
B = connect('buh')
S = json.load(open('specs.json', encoding='utf-8'))
specs = {s['Код']: s for s in S['specs']}
byspec = defaultdict(dict)
for r in S['rows']:
    byspec[r['КодСпец']][r['КодНом']] = (r['Ном'], r['ОН'], r['Кол'])

docs = q(B, """
ВЫБРАТЬ Д.Ссылка КАК Ссылка, Д.Номер КАК Номер, Д.Дата КАК Дата, Д.Номенклатура.Наименование КАК Комплект,
	Д.Склад.Наименование КАК Склад, Д.Проведен КАК Пров, Д.Количество КАК КолКомпл, Д.СуммаДокумента КАК Сумма, Д.Комментарий КАК Комм
ИЗ Документ.КомплектацияНоменклатуры КАК Д
ГДЕ НЕ Д.ПометкаУдаления И Д.Дата >= ДАТАВРЕМЯ(2025,1,1)
	И (Д.Номенклатура.Наименование ПОДОБНО "МД ООН%" ИЛИ Д.Номенклатура.Наименование ПОДОБНО "%ООН%")
УПОРЯДОЧИТЬ ПО Д.Дата""")
rows = q(B, """
ВЫБРАТЬ К.Ссылка.Номер КАК Номер, К.Номенклатура.Код КАК КодНом, К.Номенклатура.Наименование КАК Ном, СУММА(К.Количество) КАК Кол
ИЗ Документ.КомплектацияНоменклатуры.Комплектующие КАК К
ГДЕ НЕ К.Ссылка.ПометкаУдаления И К.Ссылка.Дата >= ДАТАВРЕМЯ(2025,1,1)
	И (К.Ссылка.Номенклатура.Наименование ПОДОБНО "МД ООН%" ИЛИ К.Ссылка.Номенклатура.Наименование ПОДОБНО "%ООН%")
СГРУППИРОВАТЬ ПО К.Ссылка.Номер, К.Номенклатура.Код, К.Номенклатура.Наименование""")
bydoc = defaultdict(dict)
for r in rows:
    bydoc[r['Номер']][r['КодНом']] = (r['Ном'], r['Кол'])
print("=== Комплектации ООН 2025-2026: документ -> лучшая спецификация по совпадению состава")
def house_no(name):
    m = re.search(r'№\s*(\d+)', name)
    return int(m.group(1)) if m else None
for d in docs:
    comp = bydoc.get(d['Номер'], {})
    best = None
    for code, sp in byspec.items():
        if not sp: continue
        inter = set(comp) & set(sp)
        if not inter: continue
        jac = len(inter) / len(set(comp) | set(sp))
        qty_eq = sum(1 for k in inter if abs((comp[k][1] or 0) - (sp[k][2] or 0)) < 1e-6)
        cand = (jac, qty_eq, code)
        if best is None or cand > best: best = cand
    bs = specs.get(best[2]) if best else None
    print(d['Номер'], d['Дата'][:10], "|", d['Комплект'].strip()[:40], "| склад", d['Склад'].strip()[:35], "| строк", len(comp),
          "| спец:", (bs['Спец'].strip()[:45] + " [" + best[2] + "] строк " + str(len(byspec[best[2]])) + f" jac={best[0]:.2f} кол_совп={best[1]}") if bs else "-")
json.dump({"docs": docs, "rows": rows}, open('oon_kompl.json', 'w', encoding='utf-8'), ensure_ascii=False)

print("\n=== попарное сравнение спецификаций соседних домов ООН (49..60, вариант 'Виготовлення ... NN')")
lst = []
for code, s in specs.items():
    nm = s['Спец'].strip()
    m = re.match(r'Виготовлення\s+модульних будинків, типу -А1 \(ООН\)\s*(\d+)$', nm)
    if m and byspec.get(code):
        lst.append((int(m.group(1)), code))
lst.sort()
prev = None
for n, code in lst:
    sp = byspec[code]
    if prev:
        pn, pc = prev; ps = byspec[pc]
        only_this = set(sp) - set(ps); only_prev = set(ps) - set(sp)
        common = set(sp) & set(ps)
        qd = sum(1 for k in common if abs((sp[k][2] or 0) - (ps[k][2] or 0)) > 1e-6)
        print(f"дом {pn:>2} -> {n:>2}: строк {len(ps)}->{len(sp)}; общих {len(common)}, кол-во отличается {qd}; только в {n}: {len(only_this)}, только в {pn}: {len(only_prev)}")
    prev = (n, code)
