# -*- coding: utf-8 -*-
"""Верификация водопада прибыли (формулы Excel) против хранимых ИтогиОбщие.

Формулы (из !PL по компании Червень 2026.xlsx, лист Глобино-2):
  МаржДоход       = ОД − СС
  МаржДоход%      = МаржДоход / ОД
  ОперПрибыль     = МаржДоход − ОПЗ − МЗ − АЗ
  ОперПрибыль%    = ОперПрибыль / ОД
  ПослеФинЗатрат  = ОперПрибыль + ФинДоход(ст.055) − ФинРасходы(ст.056)
  ДоНалогов       = ПослеФинЗатрат − ДР
  ЧистыйДоход     = ДоНалогов − НС − НалогНаПрибыль(ст.057)
  ВРаспоряжении   = ЧистыйДоход − Дивиденды(ст.058)
  Рентабельность% = ВРаспоряжении / ОД
Группы по кодам: ОД=000000006 СС=000000001 ДР=000000007 ОПЗ=000000003 МЗ=000000005 АЗ=000000002 НС=000000008 ФД=000000004
"""
import win32com.client, sys
from collections import defaultdict
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

def fetch(text):
    q = erp.NewObject("Запрос")
    q.Text = text
    r = q.Execute().Выгрузить()
    cols = [r.Колонки.Получить(i).Имя for i in range(r.Колонки.Количество())]
    out = []
    for i in range(r.Количество()):
        row = r.Получить(i)
        out.append({c: getattr(row, c) for c in cols})
    return out

PERIOD = 'ДАТАВРЕМЯ(2026,6,1) И ДАТАВРЕМЯ(2026,6,30,23,59,59)'

# 1. ИтогиОбщие всех документов июня
itogi = fetch(f"""ВЫБРАТЬ
    ТЧ.Ссылка.ПодразделениеСтрока КАК Подр,
    ТЧ.Ссылка.ВключатьДочерние КАК ВД,
    ТЧ.НомерСтроки КАК НС,
    ТЧ.Показатель КАК Показатель,
    ТЧ.ВидПоказателя КАК ВидПоказателя,
    ТЧ.СуммаФ1 КАК СуммаФ1,
    ТЧ.СуммаФ2 КАК СуммаФ2,
    ТЧ.СуммаИтого КАК СуммаИтого
ИЗ Документ.А_ОтчетPL.ИтогиОбщие КАК ТЧ
ГДЕ ТЧ.Ссылка.Дата МЕЖДУ {PERIOD} И НЕ ТЧ.Ссылка.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Подр, НС""")

# 2. Суммы по группам из ДанныеОтчета
gr = fetch(f"""ВЫБРАТЬ
    ТЧ.Ссылка.ПодразделениеСтрока КАК Подр,
    ТЧ.Статья.Группа.Код КАК КодГруппы,
    СУММА(ТЧ.СуммаФома1) КАК Ф1,
    СУММА(ТЧ.СуммаФорма2) КАК Ф2,
    СУММА(ТЧ.Сумма) КАК Итого
ИЗ Документ.А_ОтчетPL.ДанныеОтчета КАК ТЧ
ГДЕ ТЧ.Ссылка.Дата МЕЖДУ {PERIOD} И НЕ ТЧ.Ссылка.ПометкаУдаления
СГРУППИРОВАТЬ ПО ТЧ.Ссылка.ПодразделениеСтрока, ТЧ.Статья.Группа.Код""")

# 3. Суммы по спец-статьям группы ФД
st = fetch(f"""ВЫБРАТЬ
    ТЧ.Ссылка.ПодразделениеСтрока КАК Подр,
    ТЧ.Статья.Код КАК КодСтатьи,
    СУММА(ТЧ.СуммаФома1) КАК Ф1,
    СУММА(ТЧ.СуммаФорма2) КАК Ф2,
    СУММА(ТЧ.Сумма) КАК Итого
ИЗ Документ.А_ОтчетPL.ДанныеОтчета КАК ТЧ
ГДЕ ТЧ.Ссылка.Дата МЕЖДУ {PERIOD} И НЕ ТЧ.Ссылка.ПометкаУдаления
    И ТЧ.Статья.Код В ("000000055","000000056","000000057","000000058","000000070")
СГРУППИРОВАТЬ ПО ТЧ.Ссылка.ПодразделениеСтрока, ТЧ.Статья.Код""")

# --- Свод в словари
G = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, 0.0]))   # подр -> кодгруппы -> [ф1,ф2,итого]
for r in gr:
    G[r['Подр']][S(r['КодГруппы'])] = [float(r['Ф1']), float(r['Ф2']), float(r['Итого'])]
ST = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, 0.0]))
for r in st:
    ST[r['Подр']][S(r['КодСтатьи'])] = [float(r['Ф1']), float(r['Ф2']), float(r['Итого'])]
IT = defaultdict(list)
for r in itogi:
    IT[r['Подр']].append(r)

CODES = dict(OD='000000006', SS='000000001', DR='000000007', OPZ='000000003',
             MZ='000000005', AZ='000000002', NS='000000008', FD='000000004')

def waterfall(p, idx):
    g = lambda code: G[p][code][idx]
    s = lambda code: ST[p][code][idx]
    OD  = g(CODES['OD']); SS = g(CODES['SS']); DR = g(CODES['DR'])
    OPZ = g(CODES['OPZ']); MZ = g(CODES['MZ']); AZ = g(CODES['AZ']); NS = g(CODES['NS'])
    FinD, FinR, Nalog, Div = s('000000055'), s('000000056'), s('000000057'), s('000000058')
    marj = OD - SS
    oper = marj - OPZ - MZ - AZ
    posle_fin = oper + FinD - FinR
    do_nalogov = posle_fin - DR
    chisty = do_nalogov - NS - Nalog
    v_rasp = chisty - Div
    return {
        'Маржинальный доход, грн': marj,
        'Маржинальный доход, %': (marj / OD * 100) if OD else None,
        'Операционная прибыль, грн': oper,
        'Операционная прибыль, %': (oper / OD * 100) if OD else None,
        'Прибыль после вычета финансовых затрат': posle_fin,
        'Прибыль до вычета налогов': do_nalogov,
        'Чистый доход': chisty,
        'Прибыль в распоряжении': v_rasp,
        'Рентабельность продукции, %': (v_rasp / OD * 100) if OD else None,
    }

# --- интересные подразделения: ненулевые МЗ/ДР/НС/налог/дивиденды/ст.70
interesting, rest = [], []
for p in IT:
    score = sum(1 for c in (CODES['MZ'], CODES['DR'], CODES['NS']) if abs(G[p][c][2]) > 0.005) \
          + sum(1 for c in ('000000057', '000000058', '000000070') if abs(ST[p][c][2]) > 0.005)
    (interesting if score > 0 else rest).append((score, p))
interesting.sort(reverse=True)

print(f"Документов июня с ИтогиОбщие: {len(IT)}; с ненулевыми МЗ/ДР/НС/налог/дивиденды/ст70: {len(interesting)}")
print("Топ интересных:", [(p, sc) for sc, p in interesting[:8]])

# --- показать состав ИтогиОбщие первого документа (какие Показатель/ВидПоказателя)
p0 = interesting[0][1] if interesting else list(IT)[0]
print("=" * 100)
print(f"Состав ИтогиОбщие «{p0}»:")
for r in IT[p0]:
    print(f"  НС={int(r['НС'])} Показатель={S(r['Показатель'])!r} Вид={S(r['ВидПоказателя'])!r} "
          f"Ф1={float(r['СуммаФ1']):,.2f} Ф2={float(r['СуммаФ2']):,.2f} Итого={float(r['СуммаИтого']):,.2f}")

# --- верификация: топ-5 интересных + Глобино-2 + все ВД=Истина в интересных
check = [p for _, p in interesting[:5]]
if 'Глобино-2' not in check and 'Глобино-2' in IT:
    check.append('Глобино-2')

TOL = 0.01
for p in check:
    print("=" * 100)
    print(f"ПОДРАЗДЕЛЕНИЕ: {p}")
    print(f"  группы (Итого): " + " ".join(f"{k}={G[p][v][2]:,.0f}" for k, v in CODES.items()))
    print(f"  спецстатьи (Итого): 055={ST[p]['000000055'][2]:,.2f} 056={ST[p]['000000056'][2]:,.2f} "
          f"057={ST[p]['000000057'][2]:,.2f} 058={ST[p]['000000058'][2]:,.2f} 070={ST[p]['000000070'][2]:,.2f}")
    for idx, col in ((0, 'СуммаФ1'), (1, 'СуммаФ2'), (2, 'СуммаИтого')):
        wf = waterfall(p, idx)
        for r in IT[p]:
            pok = S(r['Показатель']).strip()
            stored = float(r[col])
            # матчим по началу наименования
            match = None
            for k in wf:
                if pok.lower().startswith(k.lower()[:20]) or k.lower().startswith(pok.lower()[:20]):
                    match = k
                    break
            if match is None:
                if idx == 2:
                    print(f"  ?? нет формулы для Показатель={pok!r}")
                continue
            calc = wf[match]
            if calc is None:
                continue
            ok = abs(calc - stored) <= TOL
            flag = "OK " if ok else "FAIL"
            if not ok or idx == 2:
                print(f"  [{flag}] {col} {pok!r}: расчёт={calc:,.2f} хранимое={stored:,.2f} Δ={calc-stored:,.2f}")
