# -*- coding: utf-8 -*-
# RED: blast-radius для бага Форма1 «Зарплата управл» (Батрак 17717.12 vs 15400).
# Читает СОХРАНЁННЫЕ ТЧ док №1 (состояние, давшее симптом) — БЕЗ заполнения/записи (LESSONS §19).
# Реконсиляция per-ФЛ: GROSSбух, НДФЛ, NETупр, Казна Ф1/Ф2, выход «Зарплата управл» Ф1/Ф2,
# А_ВзСС Ф1 = (ЗПупр Ф1 + Удержание Ф1) − НачисленныйНДФЛ Ф1. Классификация Op2a (NETупр>0) / Op2b.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected ERP")

def en(v):
    return erp.XMLСтрока(v) if erp.ЗначениеЗаполнено(v) else ""
def filled(v):
    return erp.ЗначениеЗаполнено(v)
def uid_of(ref):
    return erp.String(ref.УникальныйИдентификатор())

# --- найти док №1 (дек 2025) ---
q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С, Номер КАК Н ИЗ Документ.А_ОтражениеЗПпоКазне "
           "ГДЕ Номер = \"000000001\" И Дата МЕЖДУ ДАТАВРЕМЯ(2025,12,1,0,0,0) И ДАТАВРЕМЯ(2025,12,31,23,59,59)")
s = q.Выполнить().Выбрать(); s.Следующий()
obj = s.С.ПолучитьОбъект()
print(f"Док №{s.Н} (сохранённое состояние, без заполнения)\n")

FL = {}  # uid -> dict
def rec(uid, name):
    d = FL.get(uid)
    if d is None:
        d = dict(name=name, gross=0.0, ndfl=0.0, netupr=0.0,
                 kazna_f1=0.0, kazna_f2=0.0, out_f1_zp=0.0, out_f1_ud=0.0,
                 out_f2_zp=0.0, out_ndfl_f1=0.0, esv_f1=0.0)
        FL[uid] = d
    return d

# НачисленияБух -> GROSSбух
for r in obj.НачисленияБух:
    if not filled(r.ФизЛицо): continue
    d = rec(uid_of(r.ФизЛицо), str(r.ФизЛицо.Наименование))
    d['gross'] += float(r.Сумма)

# НалогиБухгалтерия -> НДФЛ (СуммаНДФЛ для НДФЛ+ВС)
for r in obj.НалогиБухгалтерия:
    if not filled(r.Сотрудник): continue
    fl = r.Сотрудник.ФизическоеЛицо
    if not filled(fl): continue
    tn = en(r.ТипНалога)
    if tn in ("НДФЛ", "ВоенныйСбор"):
        d = rec(uid_of(fl), str(fl.Наименование))
        d['ndfl'] += float(r.СуммаНДФЛ)

# НачисленияУпр -> NETупр
for r in obj.НачисленияУпр:
    if not filled(r.ФизЛицо): continue
    d = rec(uid_of(r.ФизЛицо), str(r.ФизЛицо.Наименование))
    d['netupr'] += float(r.Сумма)

# РаспределениеКазна -> Казна Ф1/Ф2 (СуммаНачисления, не-удержание)
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
        if filled(sp2):
            ud = bool(sp2.А_ЭтоУдержание)
    if ud: continue
    d = rec(uid_of(fl), str(fl.Наименование))
    if en(r.ФормаPL) == "Форма1": d['kazna_f1'] += sn
    elif en(r.ФормаPL) == "Форма2": d['kazna_f2'] += sn

# НачисленнаяЗарплатаИВзносыПоФизлицам -> выход
for r in obj.НачисленнаяЗарплатаИВзносыПоФизлицам:
    if not filled(r.ФизическоеЛицо): continue
    d = rec(uid_of(r.ФизическоеЛицо), str(r.ФизическоеЛицо.Наименование))
    f = en(r.ФормаPL); summ = float(r.Сумма); vzn = float(r.ВзносыВсего)
    ud = False
    sp = r.СпособОтраженияЗарплатыВБухучете
    if filled(sp): ud = bool(sp.А_ЭтоУдержание)
    if f == "Форма1":
        if vzn != 0: d['esv_f1'] += vzn
        elif ud: d['out_f1_ud'] += summ
        else: d['out_f1_zp'] += summ
    elif f == "Форма2":
        if vzn == 0 and not ud: d['out_f2_zp'] += summ

# НачисленныйНДФЛ Ф1
for r in obj.НачисленныйНДФЛ:
    if not filled(r.ФизическоеЛицо): continue
    if en(r.ФормаPL) != "Форма1": continue
    d = rec(uid_of(r.ФизическоеЛицо), str(r.ФизическоеЛицо.Наименование))
    d['out_ndfl_f1'] += float(r.Сумма)

# --- анализ ---
op2a = op2b = 0
bad = []
for uid, d in FL.items():
    net = d['netupr'] if d['netupr'] > 0 else (d['kazna_f1'] + d['kazna_f2'])
    op = "2a" if d['netupr'] > 0 else "2b"
    if op == "2a": op2a += 1
    else: op2b += 1
    # А_ВзСС Приход Ф1 = (ЗПупр Ф1 + Удержание Ф1) − НачисленныйНДФЛ Ф1
    vzss_f1 = (d['out_f1_zp'] + d['out_f1_ud']) - d['out_ndfl_f1']
    # эталон take-home Ф1 = clamp(GROSSбух − НДФЛ, 0, NET)
    correct = max(0.0, min(d['gross'] - d['ndfl'], net))
    delta = vzss_f1 - correct
    d.update(op=op, net=net, vzss_f1=vzss_f1, correct=correct, delta=delta)
    if abs(delta) > 0.01:
        bad.append(d)

print(f"ФЛ всего: {len(FL)} | Op2a (есть НачисленияУпр): {op2a} | Op2b (нет, fallback Казна): {op2b}\n")
print(f"=== ФЛ с А_ВзСС Ф1 ≠ эталон take-home (|Δ|>0.01): {len(bad)} ===")
bad.sort(key=lambda d: -abs(d['delta']))
print(f"{'ФЛ':40s} {'Op':3s} {'GROSS':>10s} {'НДФЛ':>9s} {'NET':>10s} {'Казна_Ф1':>10s} {'ВзСС_Ф1':>10s} {'эталон':>10s} {'Δ':>10s}")
sum_abs = 0.0
for d in bad:
    sum_abs += abs(d['delta'])
    print(f"{d['name'][:40]:40s} {d['op']:3s} {d['gross']:10.2f} {d['ndfl']:9.2f} {d['net']:10.2f} "
          f"{d['kazna_f1']:10.2f} {d['vzss_f1']:10.2f} {d['correct']:10.2f} {d['delta']:10.2f}")
print(f"\nΣ|Δ| = {sum_abs:,.2f}")
op2a_bad = sum(1 for d in bad if d['op']=='2a')
op2b_bad = sum(1 for d in bad if d['op']=='2b')
print(f"из них Op2a: {op2a_bad} | Op2b: {op2b_bad}")

# Батрак отдельно
print("\n=== БАТРАК ===")
for uid, d in FL.items():
    if d['name'].startswith("Батрак"):
        print(f"  {d['name']} | Op{d['op']}")
        print(f"  GROSSбух={d['gross']:.2f} НДФЛ={d['ndfl']:.2f} NETупр={d['netupr']:.2f} Казна(Ф1={d['kazna_f1']:.2f} Ф2={d['kazna_f2']:.2f})")
        print(f"  выход: ЗПупр Ф1={d['out_f1_zp']:.2f} Удерж Ф1={d['out_f1_ud']:.2f} ЗПупр Ф2={d['out_f2_zp']:.2f} ЕСВ Ф1={d['esv_f1']:.2f} НДФЛ Ф1={d['out_ndfl_f1']:.2f}")
        print(f"  А_ВзСС Ф1={d['vzss_f1']:.2f} | эталон(GROSS−НДФЛ)={d['correct']:.2f} | Δ={d['delta']:.2f}")

erp = None
print("\nDone (read-only, не записан).")
