# -*- coding: utf-8 -*-
"""Rule #-1 тест: Свод_РасчетыСПартнерами — переназначение Аналитика1/2/3
(Аналитика1=ПРЕДСТАВЛЕНИЕ(Договор.А_ВидКонтрагента), 2=Контрагент, 3=Договор).

Извлекает запрос из живого ObjectModule.bsl, гоняет ОРИГИНАЛ и МОДИФИКАЦИЮ
на янв2026/ТОВ, сверяет sum-инвариант (Σ НО/КО и число строк не меняются),
показывает примеры заполнения Аналитика1.
"""
import re, sys
import win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BSL = r"C:\Configuration_downloads\BASERP25\Documents\А_ФинРез_Баланс\Ext\ObjectModule.bsl"

# --- извлечь текст запроса функции ---
text = open(BSL, encoding="utf-8-sig").read()
m = re.search(r"Функция Свод_РасчетыСПартнерами.*?Запрос\.Текст =\s*\n(.*?)\";\s*\n",
              text, re.S)
assert m, "запрос не найден"
raw_lines = m.group(1).split("\n")
q_lines = []
for ln in raw_lines:
    s = ln.strip()
    if s.startswith('"'):
        s = s[1:]
    elif s.startswith('|'):
        s = s[1:]
    q_lines.append(s)
query_orig = "\n".join(q_lines).replace('""', '"')

# --- модификация: три строки Аналитика1/2/3 в детальной ветке ---
NEW = {
    'Аналитика1': 'ВЫБОР КОГДА втРасч.Договор.А_ВидКонтрагента <> ЗНАЧЕНИЕ(Справочник.А_ВидыКонтрагентовДляБаланса.ПустаяСсылка) ТОГДА ПРЕДСТАВЛЕНИЕ(втРасч.Договор.А_ВидКонтрагента) ИНАЧЕ "" КОНЕЦ КАК Аналитика1,',
    'Аналитика2': 'ВЫБОР КОГДА втРасч.Контрагент <> ЗНАЧЕНИЕ(Справочник.Контрагенты.ПустаяСсылка) ТОГДА ПРЕДСТАВЛЕНИЕ(втРасч.Контрагент) ИНАЧЕ "" КОНЕЦ КАК Аналитика2,',
    'Аналитика3': 'ВЫБОР КОГДА втРасч.Договор <> ЗНАЧЕНИЕ(Справочник.ДоговорыКонтрагентов.ПустаяСсылка) ТОГДА ПРЕДСТАВЛЕНИЕ(втРасч.Договор) ИНАЧЕ "" КОНЕЦ КАК Аналитика3,',
}
mod_lines, replaced = [], 0
for s in query_orig.split("\n"):
    hit = None
    for key, new in NEW.items():
        if s.rstrip().endswith(f"КАК {key},") and "втРасч." in s:
            hit = new
            replaced += 1
            break
    mod_lines.append(hit if hit else s)
assert replaced == 3, f"замен={replaced}, ожидали 3"
query_mod = "\n".join(mod_lines)

# даты — литералами (без COM datetime / tz)
def with_dates(q):
    return (q.replace("&НачМес", "ДАТАВРЕМЯ(2026,1,1)")
             .replace("&КонМес", "ДАТАВРЕМЯ(2026,1,31,23,59,59)"))

# --- COM ---
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
орг = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
assert erp.ЗначениеЗаполнено(орг), "ТОВ не найдено"
искл = erp.NewObject("Массив")
ст = erp.ПланыВидовХарактеристик.СтатьиАктивовПассивов.НайтиПоНаименованию("Собственные средства")
искл.Добавить(ст)

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

o = totals(t_orig)
m_ = totals(t_mod)
print(f"ORIG: строк={o[0]}  ΣНО={o[1]:,.2f}  ΣКО={o[2]:,.2f}")
print(f"MOD : строк={m_[0]}  ΣНО={m_[1]:,.2f}  ΣКО={m_[2]:,.2f}")
ok = (o == m_)
print(f"SUM-INVARIANT: {'OK' if ok else 'FAIL'}")

# примеры заполнения новой Аналитики
shown = 0
print("\nПримеры строк с Аналитика1 (вид контрагента):")
for i in range(t_mod.Количество()):
    r = t_mod.Получить(i)
    a1 = S(r.Аналитика1)
    if a1 and shown < 5:
        print(f"  [{a1}] | {S(r.Аналитика2)[:35]} | {S(r.Аналитика3)[:45]} | КО={float(r.СуммаКонечныйОстаток or 0):,.2f}")
        shown += 1
# раскладка Аналитика1
from collections import Counter
c = Counter()
for i in range(t_mod.Количество()):
    r = t_mod.Получить(i)
    c[S(r.Аналитика1) or "(пусто)"] += 1
print(f"\nРаскладка Аналитика1: {dict(c)}")
sys.exit(0 if ok else 1)
