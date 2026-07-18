# -*- coding: utf-8 -*-
"""
Финальный smoke (структурный, не зависит от дрейфа живой базы):
перепровести декабрь-2025 А_ФинРез_PL и проверить В РЕГИСТРЕ инварианты,
которые верны для ЛЮБОГО снимка данных:
  - ФормаPL заполнена только {Форма1, Форма2}, пустых строк нет;
  - присутствуют ОБЕ формы (факт реально разнесён);
  - построчно СуммаФ1+СуммаФ2 = Сумма и СуммаФ1_Excel+СуммаФ2_Excel = Сумма_Excel;
  - агрегатно Σ(Ф1+Ф2) = Σ Сумма.
Доп.: регистр сверяется со свежим прогоном того же запроса (post == query, нет потери ключа).
"""
import sys, io, re, datetime
import win32com.client, pythoncom

if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

pythoncom.CoInitialize()
ERP = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = ERP.String
WT = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-hawking-4dc5be"
REL = "Documents/А_ФинРез_PL/Ext/ObjectModule.bsl"


def from_bsl_func(src):
    seg = re.search(r'Функция\s+СформироватьЗапросСверткиPL\(\)(.*?)КонецФункции', src, re.S).group(1)
    raw = re.search(r'"(.*)"\s*;', seg, re.S).group(1)
    lines = raw.split("\n")
    return "\n".join([lines[0]] + [re.sub(r'^\t\|', '', x) for x in lines[1:]]).replace('""', '"')


with open(WT + "\\" + REL.replace("/", "\\"), "r", encoding="utf-8-sig") as f:
    MOD_Q = from_bsl_func(f.read())


def dds_cogs():
    q = ERP.NewObject("Запрос")
    q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.СтатьиДвиженияДенежныхСредств "
               "ГДЕ А_ПриёмникСебестоимостиПродажPL И НЕ ПометкаУдаления")
    r = q.Выполнить().Выгрузить()
    return r[0].Ссылка if r.Количество() else ERP.Справочники.СтатьиДвиженияДенежныхСредств.ПустаяСсылка()


org = ERP.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273")
dds = dds_cogs()
nach = datetime.datetime(2025, 12, 1)
konec = datetime.datetime(2025, 12, 31, 23, 59, 59)

# найти и перепровести
doc = ERP.NewObject("Запрос")
doc.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 ФР.Ссылка КАК С ИЗ Документ.А_ФинРез_PL КАК ФР ГДЕ ФР.Проведен И ФР.Месяц = ДАТАВРЕМЯ(2025,12,1)"
ref = doc.Выполнить().Выгрузить()[0].С
print(f"Документ: {S(ref)}")
ref.ПолучитьОбъект().Записать(ERP.РежимЗаписиДокумента.Проведение, ERP.РежимПроведенияДокумента.Неоперативный)
print("перепроведён OK")

# регистр построчно
qr = ERP.NewObject("Запрос")
qr.Текст = """ВЫБРАТЬ Р.ФормаPL КАК Форма, Р.СуммаФ1 КАК Ф1, Р.СуммаФ2 КАК Ф2, Р.Сумма КАК С,
Р.СуммаФ1_Excel КАК Ф1E, Р.СуммаФ2_Excel КАК Ф2E, Р.Сумма_Excel КАК СE
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р ГДЕ Р.Регистратор = &Рег"""
qr.УстановитьПараметр("Рег", ref)
tz = qr.Выполнить().Выгрузить()

f1v, f2v = S(ERP.Перечисления.А_ФормыPL.Форма1), S(ERP.Перечисления.А_ФормыPL.Форма2)
n = tz.Количество()
forms = {}
bad_fact = bad_excel = empty = 0
sF1 = sF2 = sS = 0.0
for i in range(n):
    r = tz[i]
    fv = S(r.Форма)
    forms[fv if fv else "<пусто>"] = forms.get(fv if fv else "<пусто>", 0) + 1
    if not fv:
        empty += 1
    if abs(float(r.Ф1) + float(r.Ф2) - float(r.С)) > 0.01:
        bad_fact += 1
    if abs(float(r.Ф1E) + float(r.Ф2E) - float(r.СE)) > 0.01:
        bad_excel += 1
    sF1 += float(r.Ф1); sF2 += float(r.Ф2); sS += float(r.С)

print(f"\nрегистр: {n} строк, формы={forms}")
print(f"Σ СуммаФ1={sF1:,.2f}  Σ СуммаФ2={sF2:,.2f}  Σ Сумма={sS:,.2f}  Σ(Ф1+Ф2)={sF1+sF2:,.2f}")

# свежий MOD_Q (тот же момент)
q = ERP.NewObject("Запрос")
q.Текст = MOD_Q
for nm, vv in [("НачалоПериода", nach), ("КонецПериода", konec), ("ДДСCoGS", dds), ("Организация", org)]:
    q.УстановитьПараметр(nm, vv)
mq = q.Выполнить().Выгрузить()
mqs = sum(float(mq[i].Сумма) for i in range(mq.Количество()))
print(f"свежий MOD_Q: N={mq.Количество()}  Σ Сумма={mqs:,.2f}  (Δ vs регистр={mqs-sS:,.2f})")

errs = []
if empty: errs.append(f"{empty} строк с пустой ФормаPL")
if "Форма1" not in forms or "Форма2" not in forms: errs.append(f"нет обеих форм: {forms}")
if bad_fact: errs.append(f"{bad_fact} строк ломают Ф1+Ф2=Сумма")
if bad_excel: errs.append(f"{bad_excel} строк ломают Ф1_Excel+Ф2_Excel=Сумма_Excel")
if abs(sF1 + sF2 - sS) > 0.5: errs.append(f"Σ(Ф1+Ф2)={sF1+sF2:.2f} != Σ Сумма={sS:.2f}")

print("\n=== ИТОГ ===")
if errs:
    for e in errs: print(f"  !! {e}")
    print("########## SMOKE FAIL ##########"); sys.exit(1)
print("  [OK] ФормаPL заполнена обеими формами, пустых нет")
print("  [OK] построчно Ф1+Ф2=Сумма и Ф1_Excel+Ф2_Excel=Сумма_Excel")
print("  [OK] Σ(Ф1+Ф2)=Σ Сумма")
print("########## SMOKE OK ##########")
