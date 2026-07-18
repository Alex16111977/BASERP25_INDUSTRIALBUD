# -*- coding: utf-8 -*-
# GREEN-СИМУЛЯЦИЯ (read-only): доказать, что фикс (Op2b Ф1 = Казна_Ф1 вместо коэффициента)
# устраняет расхождения отчёта "Сверка взаиморасчётов: ЕРП(Форма1) vs Бухгалтерия 661/663".
#
# Логика отчёта (ObjectModule А_СравнитьОстаткиВзаиморасчетыСотрудникиЕРПсBASБухгалтерия):
#   ЕРП:   А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты(ФормаPL=Форма1) -> КонОст = Приход - Расход
#   Бух:   BuhBud Хозрасчетный 661/663, Субконто1=ФЛ -> КонОст = КонОстКт - КонОстДт
#   Разн = КонОст_ЕРП - КонОст_Бух  (Батрак: 2317.12 - 0 = 2317.12)
#
# Фикс меняет ОЗФУ Форма1-начисление с (коэффициент) на Казна_Ф1 => Приход падает на bug_delta
#   => КонОст_ЕРП(после) = КонОст_ЕРП(до) - bug_delta.   bug_delta = текущ.ОЗФУ_Ф1 - Казна_Ф1.
# Для полностью выплаченных (Бух=0, выплата=Казна_Ф1) => остаток -> 0.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected ERP")

def en(v): return erp.XMLСтрока(v) if erp.ЗначениеЗаполнено(v) else ""
def filled(v): return erp.ЗначениеЗаполнено(v)
def uid_of(ref): return erp.String(ref.УникальныйИдентификатор())

NB = "2025-12-01T00:00:00"; NE = "2025-12-31T23:59:59"

# === 1. ЕРП: А_ВзСС Форма1 ОстаткиИОбороты per ФЛ (как в отчёте) ===
q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ Ост.ФизическоеЛицо.А_ИдКод КАК ИдКод, Ост.ФизическоеЛицо.Наименование КАК ФЛ,"
           " СУММА(Ост.СуммаВзаиморасчетовПриход) КАК Нач, СУММА(Ост.СуммаВзаиморасчетовРасход) КАК Выпл,"
           " СУММА(Ост.СуммаВзаиморасчетовКонечныйОстаток) КАК КонОст"
           " ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками.ОстаткиИОбороты("
           "  ДАТАВРЕМЯ(2025,12,1,0,0,0), ДАТАВРЕМЯ(2025,12,31,23,59,59), , ,"
           "  ФормаPL = ЗНАЧЕНИЕ(Перечисление.А_ФормыPL.Форма1)) КАК Ост"
           " СГРУППИРОВАТЬ ПО Ост.ФизическоеЛицо.А_ИдКод, Ост.ФизическоеЛицо.Наименование")
erp_kon = {}; erp_name = {}; erp_nv = {}
sel = q.Выполнить().Выбрать()
while sel.Следующий():
    k = str(sel.ИдКод).strip()
    if not k: continue
    erp_kon[k] = erp_kon.get(k, 0.0) + float(sel.КонОст)
    erp_nv[k] = (float(sel.Нач), float(sel.Выпл))
    erp_name[k] = str(sel.ФЛ)
print(f"ЕРП Форма1: {len(erp_kon)} ФЛ с остатком/оборотом")

# === 2. ЕРП: bug_delta per ФЛ из сохранённых ТЧ док №1 (Op2b: out_f1_zp - kazna_f1; Op2a: 0) ===
qd = erp.NewObject("Запрос")
qd.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Документ.А_ОтражениеЗПпоКазне "
            "ГДЕ Номер = \"000000001\" И Дата МЕЖДУ ДАТАВРЕМЯ(2025,12,1,0,0,0) И ДАТАВРЕМЯ(2025,12,31,23,59,59)")
ds = qd.Выполнить().Выбрать(); ds.Следующий()
obj = ds.С.ПолучитьОбъект()

FL = {}
def rec(uid, name, idk):
    d = FL.get(uid)
    if d is None:
        d = dict(name=name, idk=idk, gross=0.0, ndfl=0.0, netupr=0.0,
                 kf1=0.0, of1zp=0.0, of1ud=0.0, ondfl=0.0)
        FL[uid] = d
    return d
def idkod(fl):
    try:
        v = fl.А_ИдКод
        return str(v).strip() if v else ""
    except: return ""

for r in obj.НачисленияБух:
    if filled(r.ФизЛицо): rec(uid_of(r.ФизЛицо), str(r.ФизЛицо.Наименование), idkod(r.ФизЛицо))['gross'] += float(r.Сумма)
for r in obj.НалогиБухгалтерия:
    if not filled(r.Сотрудник): continue
    fl = r.Сотрудник.ФизическоеЛицо
    if filled(fl) and en(r.ТипНалога) in ("НДФЛ","ВоенныйСбор"):
        rec(uid_of(fl), str(fl.Наименование), idkod(fl))['ndfl'] += float(r.СуммаНДФЛ)
for r in obj.НачисленияУпр:
    if filled(r.ФизЛицо): rec(uid_of(r.ФизЛицо), str(r.ФизЛицо.Наименование), idkod(r.ФизЛицо))['netupr'] += float(r.Сумма)
for r in obj.РаспределениеКазна:
    if not filled(r.Сотрудник): continue
    fl = r.Сотрудник.ФизическоеЛицо
    if not filled(fl): continue
    sn = float(r.СуммаНачисления)
    if sn == 0: continue
    ud = False
    sp = r.СтатьяДвиженияДенежныхСредств
    if filled(sp):
        sp2 = sp.А_СпособОтраженияЗарплатыВБухучете
        if filled(sp2): ud = bool(sp2.А_ЭтоУдержание)
    if ud: continue
    if en(r.ФормаPL) == "Форма1":
        rec(uid_of(fl), str(fl.Наименование), idkod(fl))['kf1'] += sn
for r in obj.НачисленнаяЗарплатаИВзносыПоФизлицам:
    if not filled(r.ФизическоеЛицо): continue
    d = rec(uid_of(r.ФизическоеЛицо), str(r.ФизическоеЛицо.Наименование), idkod(r.ФизическоеЛицо))
    if en(r.ФормаPL) != "Форма1": continue
    summ = float(r.Сумма); vzn = float(r.ВзносыВсего)
    if vzn != 0: continue
    ud = False
    sp = r.СпособОтраженияЗарплатыВБухучете
    if filled(sp): ud = bool(sp.А_ЭтоУдержание)
    if ud: d['of1ud'] += summ
    else: d['of1zp'] += summ
for r in obj.НачисленныйНДФЛ:
    if not filled(r.ФизическоеЛицо): continue
    if en(r.ФормаPL) != "Форма1": continue
    rec(uid_of(r.ФизическоеЛицо), str(r.ФизическоеЛицо.Наименование), idkod(r.ФизическоеЛицо))['ondfl'] += float(r.Сумма)

bug = {}  # idk -> bug_delta
for d in FL.values():
    net = d['netupr'] if d['netupr'] > 0 else d['kf1'] + (0)  # для Op2b net не нужен (clamp не кусает)
    correct = max(0.0, d['kf1'])  # Op2b эталон = Казна_Ф1 (= GROSS-НДФЛ, доказано 0 расхожд.)
    if d['netupr'] > 0:
        correct = max(0.0, min(d['gross'] - d['ndfl'], d['netupr']))  # Op2a — без изменений (bug=0)
    vzss_f1 = (d['of1zp'] + d['of1ud']) - d['ondfl']
    bd = vzss_f1 - correct
    if d['idk']:
        bug[d['idk']] = bug.get(d['idk'], 0.0) + bd

# === 3. BuhBud 661/663 КонОст per ФЛ (как в отчёте), орг из А_ВБалансе ===
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
print("Connected BuhBud")
qo = erp.NewObject("Запрос")
qo.Текст = ("ВЫБРАТЬ Организации.КодПоЕДРПОУ КАК ЕДРПОУ ИЗ Справочник.Организации КАК Организации"
            " ГДЕ Организации.А_ВБалансе = ИСТИНА И Организации.КодПоЕДРПОУ <> \"\"")
mas = buh.NewObject("Массив")
so = qo.Выполнить().Выбрать()
while so.Следующий(): mas.Добавить(str(so.ЕДРПОУ))

qb = buh.NewObject("Запрос")
qb.УстановитьПараметр("масЕДРПОУ", mas)
qb.Текст = ("ВЫБРАТЬ ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод КАК ИдКод,"
            " СУММА(Ост.СуммаКонечныйОстатокКт) КАК КонКт, СУММА(Ост.СуммаКонечныйОстатокДт) КАК КонДт"
            " ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты("
            "  ДАТАВРЕМЯ(2025,12,1,0,0,0), ДАТАВРЕМЯ(2025,12,31,23,59,59)) КАК Ост"
            " ГДЕ Ост.Счет.Код В (\"661\",\"663\") И Ост.Субконто1 ССЫЛКА Справочник.ФизическиеЛица"
            "  И Ост.Организация.КодПоЕДРПОУ В (&масЕДРПОУ)"
            " СГРУППИРОВАТЬ ПО ВЫРАЗИТЬ(Ост.Субконто1 КАК Справочник.ФизическиеЛица).А_ИдКод")
buh_kon = {}
sb = qb.Выполнить().Выбрать()
while sb.Следующий():
    k = str(sb.ИдКод).strip()
    if not k: continue
    buh_kon[k] = buh_kon.get(k, 0.0) + (float(sb.КонКт) - float(sb.КонДт))
buh = None
print(f"BuhBud 661/663: {len(buh_kon)} ФЛ\n")

# === 4. Сравнение ДО/ПОСЛЕ ===
keys = set(erp_kon) | set(buh_kon)
rows = []
for k in keys:
    e = erp_kon.get(k, 0.0); b = buh_kon.get(k, 0.0); bd = bug.get(k, 0.0)
    razn_do = e - b
    razn_posle = (e - bd) - b
    rows.append(dict(idk=k, name=erp_name.get(k, "?"), e=e, b=b, bd=bd,
                     razn_do=razn_do, razn_posle=razn_posle))

sum_abs_do = sum(abs(r['razn_do']) for r in rows)
sum_abs_posle = sum(abs(r['razn_posle']) for r in rows)
sum_bug = sum(bug.values())
resolved = [r for r in rows if abs(r['razn_do']) > 0.01 and abs(r['razn_posle']) <= 0.01]
improved = [r for r in rows if abs(r['razn_posle']) < abs(r['razn_do']) - 0.01]
worsened = [r for r in rows if abs(r['razn_posle']) > abs(r['razn_do']) + 0.01]

print("="*70)
print(f"Σ|Разн.КонОстаток| ДО фикса     = {sum_abs_do:>14,.2f}")
print(f"Σ|Разн.КонОстаток| ПОСЛЕ фикса  = {sum_abs_posle:>14,.2f}")
print(f"Снижение                        = {sum_abs_do - sum_abs_posle:>14,.2f}")
print(f"Σ bug_delta (Op2b Ф1 перекос)   = {sum_bug:>14,.2f}")
print(f"ФЛ полностью устранены (->0):    {len(resolved)}")
print(f"ФЛ улучшены:                     {len(improved)}")
print(f"ФЛ ухудшены (контроль, д.б. 0):  {len(worsened)}")
print("="*70)

print("\n=== БАТРАК ===")
for r in rows:
    if r['name'].startswith("Батрак"):
        print(f"  ЕРП КонОст={r['e']:.2f} Бух={r['b']:.2f} bug_delta={r['bd']:.2f}")
        print(f"  Разн ДО={r['razn_do']:.2f}  ->  Разн ПОСЛЕ={r['razn_posle']:.2f}  {'✓ УСТРАНЕНО' if abs(r['razn_posle'])<=0.01 else ''}")

print("\n=== TOP-15 улучшений (bug-class) ===")
improved.sort(key=lambda r: -(abs(r['razn_do'])-abs(r['razn_posle'])))
print(f"{'ФЛ':40s} {'Разн_до':>11s} {'bug_delta':>11s} {'Разн_после':>11s}")
for r in improved[:15]:
    print(f"{r['name'][:40]:40s} {r['razn_do']:11.2f} {r['bd']:11.2f} {r['razn_posle']:11.2f}")

if worsened:
    print("\n!!! УХУДШЕННЫЕ (требуют внимания):")
    for r in worsened[:10]:
        print(f"  {r['name'][:40]:40s} до={r['razn_do']:.2f} после={r['razn_posle']:.2f} bug={r['bd']:.2f}")

erp = None
print("\nDone (read-only).")
