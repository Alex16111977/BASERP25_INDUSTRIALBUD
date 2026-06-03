# -*- coding: utf-8 -*-
# Решает A vs C: для ВСЕХ Op2b-ФЛ (нет НачисленияУпр) сверяем
#   (1) ФормаPL <=> тип ДокРаспределениеЗП (Ф1<=>А_РаспределениеЗаработнойПлаты, Ф2<=>РаспределениеФ2)
#   (2) Казна_Ф1 (по ФормаPL) == GROSSбух − НДФЛ ?
# Если расхождения есть → фикс ОБЯЗАН идти по Казна.ФормаPL (правило пользователя), не по формуле.
# Read-only (LESSONS §19).
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected ERP")

def en(v): return erp.XMLСтрока(v) if erp.ЗначениеЗаполнено(v) else ""
def filled(v): return erp.ЗначениеЗаполнено(v)
def uid_of(ref): return erp.String(ref.УникальныйИдентификатор())
def metaname(ref):
    if not filled(ref): return ""
    try: return str(ref.Метаданные().Имя)
    except: return "?"

q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Документ.А_ОтражениеЗПпоКазне "
           "ГДЕ Номер = \"000000001\" И Дата МЕЖДУ ДАТАВРЕМЯ(2025,12,1,0,0,0) И ДАТАВРЕМЯ(2025,12,31,23,59,59)")
s = q.Выполнить().Выбрать(); s.Следующий()
obj = s.С.ПолучитьОбъект()

FL = {}
def rec(uid, name):
    d = FL.get(uid)
    if d is None:
        d = dict(name=name, gross=0.0, ndfl=0.0, netupr=0.0,
                 kf1_form=0.0, kf2_form=0.0, kf1_doc=0.0, kf2_doc=0.0)
        FL[uid] = d
    return d

for r in obj.НачисленияБух:
    if filled(r.ФизЛицо): rec(uid_of(r.ФизЛицо), str(r.ФизЛицо.Наименование))['gross'] += float(r.Сумма)
for r in obj.НалогиБухгалтерия:
    if not filled(r.Сотрудник): continue
    fl = r.Сотрудник.ФизическоеЛицо
    if filled(fl) and en(r.ТипНалога) in ("НДФЛ","ВоенныйСбор"):
        rec(uid_of(fl), str(fl.Наименование))['ndfl'] += float(r.СуммаНДФЛ)
for r in obj.НачисленияУпр:
    if filled(r.ФизЛицо): rec(uid_of(r.ФизЛицо), str(r.ФизЛицо.Наименование))['netupr'] += float(r.Сумма)

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
    if ud: continue  # удержания не в «Зарплата управл»
    d = rec(uid_of(fl), str(fl.Наименование))
    f = en(r.ФормаPL)
    if f == "Форма1": d['kf1_form'] += sn
    elif f == "Форма2": d['kf2_form'] += sn
    dt = metaname(r.ДокРаспределениеЗП)
    if dt == "А_РаспределениеЗаработнойПлаты": d['kf1_doc'] += sn
    elif dt == "РаспределениеФ2": d['kf2_doc'] += sn

op2b = [d for d in FL.values() if d['netupr'] == 0]
print(f"Op2b ФЛ (нет НачисленияУпр): {len(op2b)}\n")

mism_form_doc = []   # ФормаPL != тип дока
mism_gross    = []   # Казна_Ф1(форма) != GROSS-НДФЛ
for d in op2b:
    d['gross_ndfl'] = d['gross'] - d['ndfl']
    if abs(d['kf1_form'] - d['kf1_doc']) > 0.01 or abs(d['kf2_form'] - d['kf2_doc']) > 0.01:
        mism_form_doc.append(d)
    if abs(d['kf1_form'] - d['gross_ndfl']) > 0.01:
        mism_gross.append(d)

print(f"(1) ФормаPL != тип ДокРаспределениеЗП: {len(mism_form_doc)} ФЛ")
for d in mism_form_doc[:15]:
    print(f"    {d['name'][:38]:38s} Ф1форма={d['kf1_form']:.2f} Ф1док={d['kf1_doc']:.2f} | Ф2форма={d['kf2_form']:.2f} Ф2док={d['kf2_doc']:.2f}")

print(f"\n(2) Казна_Ф1(форма) != GROSS-НДФЛ: {len(mism_gross)} ФЛ")
mism_gross.sort(key=lambda d: -abs(d['kf1_form']-d['gross_ndfl']))
for d in mism_gross[:25]:
    print(f"    {d['name'][:38]:38s} Казна_Ф1={d['kf1_form']:10.2f} GROSS-НДФЛ={d['gross_ndfl']:10.2f} Δ={d['kf1_form']-d['gross_ndfl']:10.2f} (GROSS={d['gross']:.2f} НДФЛ={d['ndfl']:.2f})")

print(f"\nИТОГ: из {len(op2b)} Op2b-ФЛ — ФормаPL/Док расхождений {len(mism_form_doc)}, Казна_Ф1/GROSS-НДФЛ расхождений {len(mism_gross)}")
print("ВЫВОД:", "Казна.ФормаPL == источник истины, GROSS-НДФЛ совпадает → A и C эквивалентны" if not mism_gross and not mism_form_doc
       else "ЕСТЬ расхождения → фикс по Казна.ФормаPL (C), НЕ по GROSS-НДФЛ")
erp = None
