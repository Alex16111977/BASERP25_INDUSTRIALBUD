# -*- coding: utf-8 -*-
# GREEN: после фикса Op2b — заполнить док №1 В ПАМЯТИ (новый код), сверить ВСЕ ФЛ.
# Критерии: Op2b Ф1==Казна_Ф1 и Ф2==Казна_Ф2; Σ(Ф1+Ф2)==NET (инвариант); Op2a без дефектов;
# Батрак Ф1=15400 Ф2=20300. Без записи (LESSONS §19).
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
def en(v): return erp.XMLСтрока(v) if erp.ЗначениеЗаполнено(v) else ""
def fl(v): return erp.ЗначениеЗаполнено(v)
def uid(r): return erp.String(r.УникальныйИдентификатор())

q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Документ.А_ОтражениеЗПпоКазне "
           "ГДЕ Номер = \"000000001\" И Дата МЕЖДУ ДАТАВРЕМЯ(2025,12,1,0,0,0) И ДАТАВРЕМЯ(2025,12,31,23,59,59)")
s = q.Выполнить().Выбрать(); s.Следующий()
obj = s.С.ПолучитьОбъект()
print("Заполнение В ПАМЯТИ новым кодом...")
obj.ЗаполнитьНачисленияУпризЗуп2()
obj.ЗагрузитьНалогиИУдержанияИзЗарплатыБух()
obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
obj.ЗаполнитьДляОтражения_ДокументОтражениеЗарплатыВФинансовомУчете()

D = {}
def rec(u, n):
    d = D.get(u)
    if d is None:
        d = dict(name=n, gross=0.0, ndfl=0.0, netupr=0.0, kf1=0.0, kf2=0.0,
                 f1zp=0.0, f2zp=0.0, f1ud=0.0, ondfl=0.0)
        D[u] = d
    return d
for r in obj.НачисленияБух:
    if fl(r.ФизЛицо): rec(uid(r.ФизЛицо), str(r.ФизЛицо.Наименование))['gross'] += float(r.Сумма)
for r in obj.НалогиБухгалтерия:
    if fl(r.Сотрудник):
        f = r.Сотрудник.ФизическоеЛицо
        if fl(f) and en(r.ТипНалога) in ("НДФЛ","ВоенныйСбор"):
            rec(uid(f), str(f.Наименование))['ndfl'] += float(r.СуммаНДФЛ)
for r in obj.НачисленияУпр:
    if fl(r.ФизЛицо): rec(uid(r.ФизЛицо), str(r.ФизЛицо.Наименование))['netupr'] += float(r.Сумма)
for r in obj.РаспределениеКазна:
    if not fl(r.Сотрудник): continue
    f = r.Сотрудник.ФизическоеЛицо
    if not fl(f): continue
    sn = float(r.СуммаНачисления)
    if sn == 0: continue
    ud = False; sp = r.СтатьяДвиженияДенежныхСредств
    if fl(sp):
        sp2 = sp.А_СпособОтраженияЗарплатыВБухучете
        if fl(sp2): ud = bool(sp2.А_ЭтоУдержание)
    if ud: continue
    d = rec(uid(f), str(f.Наименование))
    if en(r.ФормаPL) == "Форма1": d['kf1'] += sn
    elif en(r.ФормаPL) == "Форма2": d['kf2'] += sn
for r in obj.НачисленнаяЗарплатаИВзносыПоФизлицам:
    if not fl(r.ФизическоеЛицо): continue
    d = rec(uid(r.ФизическоеЛицо), str(r.ФизическоеЛицо.Наименование))
    f = en(r.ФормаPL); summ = float(r.Сумма); vzn = float(r.ВзносыВсего)
    if vzn != 0: continue
    ud = False; sp = r.СпособОтраженияЗарплатыВБухучете
    if fl(sp): ud = bool(sp.А_ЭтоУдержание)
    if f == "Форма1":
        if ud: d['f1ud'] += summ
        else: d['f1zp'] += summ
    elif f == "Форма2":
        if not ud: d['f2zp'] += summ
for r in obj.НачисленныйНДФЛ:
    if fl(r.ФизическоеЛицо) and en(r.ФормаPL) == "Форма1":
        rec(uid(r.ФизическоеЛицо), str(r.ФизическоеЛицо.Наименование))['ondfl'] += float(r.Сумма)

bad_op2b = []; bad_inv = []; op2a = op2b = 0
for u, d in D.items():
    is2a = d['netupr'] > 0
    if is2a: op2a += 1
    else: op2b += 1
    net = d['netupr'] if is2a else (d['kf1'] + d['kf2'])
    if abs((d['f1zp'] + d['f2zp']) - net) > 0.02:
        bad_inv.append((d['name'], round(d['f1zp'] + d['f2zp'],2), round(net,2)))
    if not is2a:
        if abs(d['f1zp'] - d['kf1']) > 0.02 or abs(d['f2zp'] - d['kf2']) > 0.02:
            bad_op2b.append((d['name'], round(d['f1zp'],2), round(d['kf1'],2), round(d['f2zp'],2), round(d['kf2'],2)))

print(f"\nФЛ всего={len(D)} Op2a={op2a} Op2b={op2b}")
print(f"Op2b с Ф1!=Казна_Ф1 или Ф2!=Казна_Ф2: {len(bad_op2b)}  (ОЖИД 0)")
for x in bad_op2b[:20]: print("   ", x)
print(f"Нарушение инварианта Ф1+Ф2!=NET: {len(bad_inv)}  (ОЖИД 0)")
for x in bad_inv[:20]: print("   ", x)
for u, d in D.items():
    if d['name'].startswith("Батрак"):
        print(f"\nБАТРАК: Ф1zp={d['f1zp']:.2f} (ОЖИД 15400) Ф2zp={d['f2zp']:.2f} (ОЖИД 20300) "
              f"Удерж={d['f1ud']:.2f} НДФЛ={d['ondfl']:.2f}")
        print("  Ф1 take-home =", "OK" if abs(d['f1zp']-15400)<0.02 else "FAIL")
ok = (len(bad_op2b) == 0 and len(bad_inv) == 0)
print("\n=== GREEN ===" if ok else "\n=== FAIL ===")
erp = None
