# -*- coding: utf-8 -*-
"""Rule #-1: полный пакетный запрос ПолучитьОбъединенныеДанные из нового ObjectModule.bsl,
прогон в COM + регрессия против старой версии (worktree) + Acceptance A на выгрузке.
"""
import win32com.client, sys, re, datetime
from collections import defaultdict
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

NEW_BSL = r"C:\Configuration_downloads\BASERP25\Reports\А_ОтчетPL\Ext\ObjectModule.bsl"
OLD_BSL = (r"C:\Configuration_downloads\BASERP25\.claude\worktrees\baserp-pl-report-split-excel-36baec"
           r"\Reports\А_ОтчетPL\Ext\ObjectModule.bsl")


def extract_query(path):
    src = open(path, encoding='utf-8-sig').read()
    m = re.search(r'Запрос\.Текст =\s*\r?\n(\t"// =+.*?\|";)', src, re.S)
    assert m, f"запрос не найден в {path}"
    lines = []
    for raw in m.group(1).splitlines():
        s = raw.lstrip('\t')
        if s.startswith('"'):
            s = s[1:]
        elif s.startswith('|'):
            s = s[1:]
        if s.endswith('";'):
            s = s[:-2]
        lines.append(s)
    return "\n".join(lines).replace('""', '"')


v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def scalar(text):
    q = erp.NewObject("Запрос")
    q.Text = text
    t = q.Execute().Выгрузить()
    return t.Получить(0).Ссылка if t.Количество() else None


ддс_cogs = scalar("""ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.СтатьиДвиженияДенежныхСредств
ГДЕ А_ПриёмникСебестоимостиПродажPL И НЕ ПометкаУдаления""")
if ддс_cogs is None:
    ддс_cogs = erp.Справочники.СтатьиДвиженияДенежныхСредств.ПустаяСсылка()

товар_сд = scalar("""ВЫБРАТЬ ПЕРВЫЕ 1 СД.Ссылка КАК Ссылка,
    ВЫБОР КОГДА СД.Наименование = "Прочие доходи" ТОГДА 0 ИНАЧЕ 1 КОНЕЦ КАК Приоритет
ИЗ ПланВидовХарактеристик.СтатьиДоходов КАК СД
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.А_Статьи_PL.Статьи КАК СтатьиТЧ
    ПО СтатьиТЧ.СтатьяДвиженияДенежныхСредств = СД.А_СтатьяДвиженияДенежныхСредств
        И СтатьиТЧ.Ссылка.Код = "000000002" И НЕ СтатьиТЧ.Ссылка.ПометкаУдаления
ГДЕ НЕ СД.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Приоритет, СД.Наименование""")
print("ДДСCoGS:", S(ддс_cogs), "| СтатьяДоходовТовар:", S(товар_сд))


def run_full(path, with_tovar_param):
    text = extract_query(path)
    q = erp.NewObject("Запрос")
    q.Text = text
    q.SetParameter("НачалоПериода", datetime.datetime(2026, 6, 1))
    q.SetParameter("КонецПериода", datetime.datetime(2026, 6, 30, 23, 59, 59))
    q.SetParameter("ВключатьДочерние", False)
    q.SetParameter("ДДСCoGS", ддс_cogs)
    q.SetParameter("ПоказPL", True)
    q.SetParameter("ПоказЕРП", True)
    q.SetParameter("ПоказКазна", False)
    if with_tovar_param:
        q.SetParameter("СтатьяДоходовТовар", товар_сд)
    try:
        return q.Execute().Выгрузить()
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        print("FAIL:", msg)
        raise

t_new = run_full(NEW_BSL, True)
print(f"НОВЫЙ запрос: OK, строк={t_new.Количество()}")
t_old = run_full(OLD_BSL, False)
print(f"СТАРЫЙ запрос: OK, строк={t_old.Количество()}")


def totals(t):
    s_erp = s_pl = 0.0
    for i in range(t.Количество()):
        r = t.Получить(i)
        s_erp += float(r.СуммаЕРП)
        s_pl += float(r.СуммаPL)
    return s_erp, s_pl

n_erp, n_pl = totals(t_new)
o_erp, o_pl = totals(t_old)
print(f"РЕГРЕССИЯ СВОДА: СуммаЕРП new={n_erp:,.2f} old={o_erp:,.2f} Δ={n_erp-o_erp:,.2f} "
      f"{'OK' if abs(n_erp-o_erp) < 0.01 else 'FAIL'}")
print(f"                 СуммаPL  new={n_pl:,.2f} old={o_pl:,.2f} Δ={n_pl-o_pl:,.2f} "
      f"{'OK' if abs(n_pl-o_pl) < 0.01 else 'FAIL'}")

# Acceptance A: факт по PL-статьям 1/2 для подразделений Строительства (по выгрузке нового запроса)
q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Ссылка, Код ИЗ Справочник.А_Статьи_PL ГДЕ Код В ("000000001", "000000002")"""
st_map = {}
t = q.Execute().Выгрузить()
for i in range(t.Количество()):
    r = t.Получить(i)
    st_map[S(r.Ссылка)] = S(r.Код)

fact = defaultdict(lambda: defaultdict(float))
for i in range(t_new.Количество()):
    r = t_new.Получить(i)
    if float(r.СуммаЕРП) == 0:
        continue
    код = st_map.get(S(r.СтатьяPL))
    if код:
        fact[S(r.Подразделение)][код] += float(r.СуммаЕРП)

EXPECT = {  # план июнь-2026 (Acceptance A)
    "Глобино-2": ("000000001", 42655266.04),
    "КРИВОЙ РОГ ЦЕМЕНТ-2": ("000000001", 13800460.14),
    "АВРОРА. СТОЯНКА": ("000000001", 8568457.18),
    "Астарта. Тищенки": ("000000001", 110812.70),
}
for подр, (код, plan) in EXPECT.items():
    got = fact[подр][код]
    print(f"{'OK' if abs(got - plan) <= 0.01 else 'FAIL'} {подр}: PL{код[-1]} факт={got:,.2f} план={plan:,.2f}")

# Глобино-2 PL2 = товар 7 224 777.48 + прочие доходы (ДДС Прочие поступления)
got2 = fact["Глобино-2"]["000000002"]
print(f"Глобино-2 PL2 факт={got2:,.2f} (товар 7 224 777.48 + прочие доходы; "
      f"чисто товар {'OK' if abs(got2 - 7224777.48) < 100000 else '??'})")
