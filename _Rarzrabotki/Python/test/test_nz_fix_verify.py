# -*- coding: utf-8 -*-
# TDD verify: Документ.НачислениеЗарплаты 000Ц-000005 — больничные «за счёт фондов».
# READ-ONLY (читає збережений документ). Перший прогон зберігає baseline-снимок.
#
# Assertions:
#   A1: жодного від'ємного рядка Начисления з видом <> "Штраф (упр)".
#   A2: для кожного ФЛ з ФСС>0:  Σ Начисления (усі види) == bukhGROSS = ΣНачисленияБух + ΣФСС (±0.01).
#   A3 (тільки коли baseline вже існував — пост-фікс): множина ФЛ зі ЗМІНЕНИМ записом
#       (nz_oklad/nz_shtraf/nz_total) == множина ФЛ з ФСС>0. Інші ФЛ byte-identical.
import win32com.client, sys, os, json

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_nz_baseline_000005.json")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String; X = erp.XMLСтрока; Z = erp.ЗначениеЗаполнено


def num(x):
    return float(S(x).replace('\xa0', '').replace(' ', '').replace(',', '.') or "0")


def kk(ref):
    return X(ref) if Z(ref) else None


def run(text, **p):
    q = erp.NewObject("Запрос"); q.Text = text
    for k, vv in p.items():
        q.SetParameter(k, vv)
    return q.Execute().Выгрузить()


parent = run("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Р ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д ГДЕ Д.Номер=&Н", Н="000000002").Получить(0).Р
nz = run("ВЫБРАТЬ ПЕРВЫЕ 1 НЗ.Ссылка КАК Р ИЗ Документ.НачислениеЗарплаты КАК НЗ ГДЕ НЗ.А_ДокОтражениеЗПпоКазне=&О И НЕ НЗ.ПометкаУдаления", О=parent).Получить(0).Р

D = {}


def slot(k, nm):
    if k not in D:
        D[k] = dict(name=nm, bukh=0.0, fss=0.0, nz_oklad=0.0, nz_shtraf=0.0, nz_total=0.0)
    return D[k]


# bukh / fss
for tbl, fld in [("НачисленияБух", "bukh"), ("НачисленияБухЗасчетФондов", "fss")]:
    t = run(f"ВЫБРАТЬ Т.ФизЛицо КАК ФЛ, СУММА(Т.Сумма) КАК С ИЗ Документ.А_ОтражениеЗПпоКазне.{tbl} КАК Т "
            "ГДЕ Т.Ссылка=&Р СГРУППИРОВАТЬ ПО Т.ФизЛицо", Р=parent)
    for i in range(t.Количество()):
        r = t.Получить(i); k = kk(r.ФЛ)
        if k: slot(k, S(r.ФЛ))[fld] += num(r.С)

# НЗ Начисления по (ФЛ, вид)
t = run("ВЫБРАТЬ Т.Сотрудник.ФизическоеЛицо КАК ФЛ, Т.Начисление КАК Вид, СУММА(Т.Результат) КАК С "
        "ИЗ Документ.НачислениеЗарплаты.Начисления КАК Т ГДЕ Т.Ссылка=&Р "
        "СГРУППИРОВАТЬ ПО Т.Сотрудник.ФизическоеЛицо, Т.Начисление", Р=nz)
for i in range(t.Количество()):
    r = t.Получить(i); k = kk(r.ФЛ)
    if not k: continue
    s = slot(k, S(r.ФЛ)); v = num(r.С)
    s['nz_total'] += v
    if "Штраф" in S(r.Вид): s['nz_shtraf'] += v
    else: s['nz_oklad'] += v

# A1: від'ємні не-штраф рядки
neg = run("ВЫБРАТЬ Т.Сотрудник.ФизическоеЛицо.Наименование КАК ФЛ, Т.Начисление КАК Вид, Т.Результат КАК Рез "
          "ИЗ Документ.НачислениеЗарплаты.Начисления КАК Т ГДЕ Т.Ссылка=&Р И Т.Результат<0", Р=nz)
neg_bad = []
for i in range(neg.Количество()):
    r = neg.Получить(i)
    if "Штраф" not in S(r.Вид):
        neg_bad.append((S(r.ФЛ), S(r.Вид), num(r.Рез)))

fss_set = {k for k, s in D.items() if s['fss'] > 0.005}

# baseline
first_run = not os.path.exists(BASE)
baseline = {} if first_run else json.load(open(BASE, encoding='utf-8'))
if first_run:
    json.dump({k: {kk2: v for kk2, v in s.items()} for k, s in D.items()}, open(BASE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f"BASELINE SAVED -> {BASE}  (RED phase, {len(D)} ФЛ)")

print("\n=== A1: від'ємні не-штраф рядки Начисления ===")
print(f"  знайдено: {len(neg_bad)}")
for fl, v, rez in neg_bad:
    print(f"    {fl[:34]:34s} | {v[:20]:20s} | {rez:.2f}")
a1 = (len(neg_bad) == 0)

print("\n=== A2: больничні (ФСС>0) Начислено == bukhGROSS ===")
a2 = True
for k in sorted(fss_set, key=lambda x: D[x]['name']):
    s = D[k]; gross = s['bukh'] + s['fss']; ok = abs(s['nz_total'] - gross) <= 0.01
    a2 = a2 and ok
    print(f"  [{'OK' if ok else 'FAIL'}] {s['name'][:32]:32s} nz_total={s['nz_total']:>10.2f}  GROSS={gross:>10.2f}  "
          f"(oklad={s['nz_oklad']:.2f} shtraf={s['nz_shtraf']:.2f})")

a3 = True
if not first_run:
    print("\n=== A3: регресія — змінилися РІВНО больничні (ФСС>0) ===")
    changed = set()
    for k, s in D.items():
        b = baseline.get(k)
        if b is None:
            changed.add(k); continue
        if (abs(s['nz_oklad'] - b['nz_oklad']) > 0.01 or abs(s['nz_shtraf'] - b['nz_shtraf']) > 0.01
                or abs(s['nz_total'] - b['nz_total']) > 0.01):
            changed.add(k)
    extra = changed - fss_set
    missing = fss_set - changed
    a3 = (len(extra) == 0)
    print(f"  змінено ФЛ: {len(changed)}; больничних(ФСС>0): {len(fss_set)}")
    for k in changed:
        tag = "больн" if k in fss_set else "!!!ЛИШНІЙ!!!"
        print(f"    {D[k]['name'][:34]:34s} [{tag}]  total {baseline.get(k,{}).get('nz_total','?')} -> {D[k]['nz_total']:.2f}")
    if extra:
        print(f"  ПОРУШЕННЯ РЕГРЕСІЇ: зачеплено {len(extra)} НЕ-больничних ФЛ!")
    if missing:
        print(f"  УВАГА: больничні без змін (composition?): {[D[k]['name'] for k in missing]}")

print("\n" + "=" * 50)
allok = a1 and a2 and a3
print(f"A1(no-neg)={a1}  A2(GROSS)={a2}  A3(regress)={a3}  =>  {'GREEN PASS' if allok else 'RED FAIL'}")
sys.exit(0 if allok else 1)
