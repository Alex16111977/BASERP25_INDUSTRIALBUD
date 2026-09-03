# -*- coding: utf-8 -*-
"""Аудит связей «статья денег (ДДС) ↔ статья затрат (СтатьяРасходов)» в BaseERP. Только чтение.

ОБЛАСТЬ ПРОВЕРКИ (по указанию финансиста): только ДДС и статьи затрат, выбранные в справочник
А_Статьи_PL (ТЧ «Статьи»: ДДС выбрана вручную, статья затрат подставлена из ДДС.А_СтатьяРасходов).

Связи между ними в справочниках:
  ДДС.А_СтатьяРасходов                          (деньги → затраты; так заполняется ТЧ PL)
  СтатьиРасходов.А_СтатьяДвиженияДенежныхСредств (затраты → деньги; по ней ФАКТ ЕРП попадает в PL)

Правило финансиста: 1 статья затрат ↔ 1 статья денег. Печатает нарушения:
  1. статья затрат связана с 2+ статьями денег;  2. статья денег связана с 2+ статьями затрат;
  3. обратная ссылка ведёт на другую ДДС (деньги и затраты уходят в разные PL);
  4. своя ДДС статьи затрат не выбрана ни в одной PL (факт по этой статье выпадает из PL).
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def query(text):
    q = erp.NewObject("Запрос")
    q.Text = text
    return q.Execute().Выгрузить()


def uid(ref):
    return S(ref.УникальныйИдентификатор())


# --- область: ТЧ PL
r = query("""
ВЫБРАТЬ ТЧ.Ссылка.Код КАК Код, ТЧ.Ссылка КАК PL, ТЧ.СтатьяДвиженияДенежныхСредств КАК ДДС, ТЧ.СтатьяРасходов КАК СР,
    ТЧ.СтатьяДвиженияДенежныхСредств.А_СтатьяРасходов КАК СРуДДС,
    ТЧ.СтатьяРасходов.А_СтатьяДвиженияДенежныхСредств КАК ДДСуСР
ИЗ Справочник.А_Статьи_PL.Статьи КАК ТЧ
ГДЕ НЕ ТЧ.Ссылка.ПометкаУдаления
""")
dds_name, sr_name, dds_pl = {}, {}, {}
dds_sr = {}      # ДДС(в PL) -> СР по реквизиту ДДС.А_СтатьяРасходов
sr_dds = {}      # СР(в PL) -> ДДС по реквизиту СР.А_СтатьяДвиженияДенежныхСредств
dds_in_pl, sr_in_pl = set(), set()
for i in range(r.Количество()):
    x = r.Получить(i)
    d = uid(x.ДДС)
    dds_name[d] = S(x.ДДС)
    dds_pl[d] = f"{S(x.Код)} {S(x.PL)}"
    dds_in_pl.add(d)
    dds_sr[d] = None if x.СРуДДС.Пустая() else uid(x.СРуДДС)
    if not x.СРуДДС.Пустая():
        sr_name[uid(x.СРуДДС)] = S(x.СРуДДС)
    if not x.СР.Пустая():
        s = uid(x.СР)
        sr_in_pl.add(s)
        sr_name[s] = S(x.СР)
        sr_dds[s] = None if x.ДДСуСР.Пустая() else uid(x.ДДСуСР)
        if not x.ДДСуСР.Пустая():
            dds_name.setdefault(uid(x.ДДСуСР), S(x.ДДСуСР))

# --- связи только внутри области (ДДС в PL, СР в PL)
sr_to_dds, dds_to_sr = {}, {}
for d in dds_in_pl:
    s = dds_sr.get(d)
    if s and s in sr_in_pl:
        sr_to_dds.setdefault(s, set()).add(d)
        dds_to_sr.setdefault(d, set()).add(s)
for s in sr_in_pl:
    d = sr_dds.get(s)
    if d and d in dds_in_pl:
        sr_to_dds.setdefault(s, set()).add(d)
        dds_to_sr.setdefault(d, set()).add(s)

print(f"Область: ДДС в PL = {len(dds_in_pl)}, статей затрат в PL = {len(sr_in_pl)}")

print("\n=== 1. Статья затрат ↔ 2+ статей денег (внутри PL) ===")
multi_sr = sorted(((s, v) for s, v in sr_to_dds.items() if len(v) >= 2), key=lambda kv: -len(kv[1]))
print(f"всего: {len(multi_sr)}")
for s, v in multi_sr:
    own = sr_dds.get(s)
    pls = sorted({dds_pl[d] for d in v})
    print(f"  {sr_name.get(s, s)} : {len(v)} ДДС в {len(pls)} PL; своя ДДС затрат: {dds_name.get(own, '—') if own else '—'}"
          f"{' [НЕ в PL]' if own and own not in dds_in_pl else ''}")
    for d in sorted(v, key=lambda x: dds_name.get(x, x)):
        print(f"      ← {dds_name.get(d, d)}  [{dds_pl.get(d, '—')}]{'  (обратная совпала)' if own == d else ''}")

print("\n=== 2. Статья денег ↔ 2+ статей затрат (внутри PL) ===")
multi_dds = sorted(((d, v) for d, v in dds_to_sr.items() if len(v) >= 2), key=lambda kv: -len(kv[1]))
print(f"всего: {len(multi_dds)}")
for d, v in multi_dds:
    print(f"  {dds_name.get(d, d)} [{dds_pl.get(d, '—')}] : " + "; ".join(sr_name.get(s, s) for s in v))

print("\n=== 3. Обратная ссылка ведёт на другую ДДС → деньги и затраты в разных PL ===")
cnt = 0
for d in sorted(dds_in_pl, key=lambda x: dds_pl.get(x, '')):
    s = dds_sr.get(d)
    if not s:
        continue
    d2 = sr_dds.get(s)
    if d2 and d2 != d:
        cnt += 1
        pl2 = dds_pl.get(d2, "НЕ в PL")
        mark = "РАЗНЫЕ PL" if pl2 != dds_pl.get(d) else "та же PL"
        print(f"  деньги «{dds_name[d]}» [{dds_pl[d]}] → затраты «{sr_name.get(s, s)}» → факт идёт по ДДС «{dds_name.get(d2, d2)}» [{pl2}]  {mark}")
print(f"всего: {cnt}")

print("\n=== 4. Своя ДДС статьи затрат не выбрана ни в одной PL (факт выпадает из PL) ===")
cnt = 0
for s in sorted(sr_in_pl, key=lambda x: sr_name.get(x, x)):
    d = sr_dds.get(s)
    if d is None:
        cnt += 1
        print(f"  «{sr_name[s]}»: у статьи затрат ПУСТАЯ статья денег")
    elif d not in dds_in_pl:
        cnt += 1
        print(f"  «{sr_name[s]}» → ДДС «{dds_name.get(d, d)}» не выбрана ни в одной PL")
print(f"всего: {cnt}")

print("\n=== 5. ДДС в PL без статьи затрат (пустая А_СтатьяРасходов) ===")
lst = [d for d in dds_in_pl if not dds_sr.get(d)]
print(f"всего: {len(lst)} → " + "; ".join(f"{dds_name[d]} [{dds_pl[d]}]" for d in lst))
