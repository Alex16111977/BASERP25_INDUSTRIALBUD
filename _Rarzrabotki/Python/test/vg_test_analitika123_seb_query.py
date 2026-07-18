# -*- coding: utf-8 -*-
"""Rule #-1 тест: Свод_СебестоимостьТоваров — Аналитика1/2/3 =
Склад.Подразделение / Склад / Номенклатура (было Склад/Номенклатура/"").
Гейт: sum-инвариант ORIG vs MOD (строки, Σ НО/КО) + примеры заполнения."""
import re, sys
import win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BSL = r"C:\Configuration_downloads\BASERP25\Documents\А_ФинРез_Баланс\Ext\ObjectModule.bsl"

text = open(BSL, encoding="utf-8-sig").read()
m = re.search(r"Функция Свод_СебестоимостьТоваров.*?Запрос\.Текст =\s*\n(.*?)\";\s*\n",
              text, re.S)
assert m, "запрос не найден"
q_lines = []
for ln in m.group(1).split("\n"):
    s = ln.strip()
    if s.startswith('"'):
        s = s[1:]
    elif s.startswith('|'):
        s = s[1:]
    q_lines.append(s)
query_orig = "\n".join(q_lines).replace('""', '"')

NEW = {
    'Аналитика1': 'ВЫБОР КОГДА втСеб.Склад.Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка) ТОГДА ПРЕДСТАВЛЕНИЕ(втСеб.Склад.Подразделение) ИНАЧЕ "" КОНЕЦ КАК Аналитика1,',
    'Аналитика2': 'ВЫБОР КОГДА втСеб.Склад <> ЗНАЧЕНИЕ(Справочник.Склады.ПустаяСсылка) ТОГДА ПРЕДСТАВЛЕНИЕ(втСеб.Склад) ИНАЧЕ "" КОНЕЦ КАК Аналитика2,',
    'Аналитика3': 'ВЫБОР КОГДА втСеб.Номенклатура <> ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка) ТОГДА ПРЕДСТАВЛЕНИЕ(втСеб.Номенклатура) ИНАЧЕ "" КОНЕЦ КАК Аналитика3,',
}
mod_lines, replaced = [], 0
for s in query_orig.split("\n"):
    hit = None
    for key, new in NEW.items():
        if s.rstrip().endswith(f"КАК {key},"):
            hit = new
            replaced += 1
            break
    mod_lines.append(hit if hit else s)
assert replaced == 3, f"замен={replaced}, ожидали 3"
query_mod = "\n".join(mod_lines)

def with_dates(q):
    return (q.replace("&НачМес", "ДАТАВРЕМЯ(2026,1,1)")
             .replace("&КонМес", "ДАТАВРЕМЯ(2026,1,31,23,59,59)"))

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
орг = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
искл = erp.NewObject("Массив")
искл.Добавить(erp.ПланыВидовХарактеристик.СтатьиАктивовПассивов.НайтиПоНаименованию("Собственные средства"))

def run(q):
    qq = erp.NewObject("Запрос")
    qq.Text = with_dates(q)
    qq.SetParameter("Орг", орг)
    qq.SetParameter("Искл", искл)
    return qq.Execute().Выгрузить()

def totals(t):
    n, sno, sko = t.Количество(), 0.0, 0.0
    for i in range(n):
        r = t.Получить(i)
        sno += float(r.СуммаНачальныйОстаток or 0)
        sko += float(r.СуммаКонечныйОстаток or 0)
    return n, round(sno, 2), round(sko, 2)

try:
    t_orig = run(query_orig)
except Exception as e:
    print(f"FAIL ORIG: {e.excepinfo[2] if hasattr(e,'excepinfo') and e.excepinfo else e}")
    sys.exit(1)
try:
    t_mod = run(query_mod)
except Exception as e:
    print(f"FAIL MOD: {e.excepinfo[2] if hasattr(e,'excepinfo') and e.excepinfo else e}")
    sys.exit(1)

o, m_ = totals(t_orig), totals(t_mod)
print(f"ORIG: строк={o[0]}  ΣНО={o[1]:,.2f}  ΣКО={o[2]:,.2f}")
print(f"MOD : строк={m_[0]}  ΣНО={m_[1]:,.2f}  ΣКО={m_[2]:,.2f}")
ok = (o == m_)
print(f"SUM-INVARIANT: {'OK' if ok else 'FAIL'}")

shown = 0
print("\nПримеры строк (Аналитика1=подразделение склада | Аналитика2=склад | Аналитика3=номенклатура):")
for i in range(t_mod.Количество()):
    r = t_mod.Получить(i)
    a1 = S(r.Аналитика1)
    if a1 and shown < 5:
        print(f"  [{a1}] | {S(r.Аналитика2)[:30]} | {S(r.Аналитика3)[:40]} | КО={float(r.СуммаКонечныйОстаток or 0):,.2f}")
        shown += 1
from collections import Counter
c = Counter()
for i in range(t_mod.Количество()):
    r = t_mod.Получить(i)
    c[S(r.Аналитика1) or "(пусто)"] += 1
print(f"\nТоп раскладки Аналитика1: {dict(sorted(c.items(), key=lambda x: -x[1])[:8])}")
sys.exit(0 if ok else 1)
