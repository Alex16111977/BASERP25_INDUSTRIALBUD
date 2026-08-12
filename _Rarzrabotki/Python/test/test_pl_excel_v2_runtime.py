# -*- coding: utf-8 -*-
"""Runtime-приёмка «PL как в Excel (факт)» v2 на рабочей BaseERP (июнь-2026, Глобино-2):
1) прямой вызов ПолучитьОбъединенныеДанные: колонки, водопад Ф1/Ф2/Ит vs независимый
   пересчёт, база %, полный шаблон;
2) ПолныйШаблон=Ложь -> статей меньше;
3) компоновка варианта: последовательность строк 1:1 с листом Excel + хвосты каталога,
   проценты со знаком %, пустой % у ОД;
4) регрессия: свод сумм, 3 старых варианта.
"""
import win32com.client, sys, datetime
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl

XLSX = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Червень_26\!PL по компании Червень 2026.xlsx"
FAILS = []

def num(x):
    return float(x) if x is not None else 0.0

def check(ok, msg):
    print(("OK  " if ok else "FAIL") + " " + msg)
    if not ok:
        FAILS.append(msg)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
D1, D2 = datetime.datetime(2026, 6, 1), datetime.datetime(2026, 6, 30, 23, 59, 59)

# каталог: коды групп/статей, имена, сорт
q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Ст.Наименование КАК Наименование, Ст.Код КАК Код, Ст.Сорт КАК Сорт,
    Ст.Группа.Код КАК КодГруппы
ИЗ Справочник.А_Статьи_PL КАК Ст ГДЕ НЕ Ст.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Ст.Группа.Сорт, Ст.Сорт, Ст.Код"""
t = q.Execute().Выгрузить()
catalog = []
for i in range(t.Количество()):
    r = t.Получить(i)
    catalog.append({"name": S(r.Наименование).strip(), "code": S(r.Код),
                    "sort": float(r.Сорт), "gcode": S(r.КодГруппы)})
st_code_by_name = {}
q2 = erp.NewObject("Запрос")
q2.Text = """ВЫБРАТЬ Ссылка, Код ИЗ Справочник.А_Статьи_PL ГДЕ НЕ ПометкаУдаления"""
t2 = q2.Execute().Выгрузить()
st_code_by_ref = {}
for i in range(t2.Количество()):
    r = t2.Получить(i)
    st_code_by_ref[S(r.Ссылка)] = S(r.Код)
q3 = erp.NewObject("Запрос")
q3.Text = """ВЫБРАТЬ Ссылка, Код ИЗ Справочник.А_ГруппаСтатей_PL"""
t3 = q3.Execute().Выгрузить()
gr_code_by_ref = {}
for i in range(t3.Количество()):
    r = t3.Получить(i)
    gr_code_by_ref[S(r.Ссылка)] = S(r.Код)

# ============ 1. Прямой вызов ============
otchet = erp.Отчеты.А_ОтчетPL.Создать()
tab = otchet.ПолучитьОбъединенныеДанные(D1, D2, False, False, True, False, True, True)
cols = [tab.Колонки.Получить(i).Имя for i in range(tab.Колонки.Количество())]
for c in ("ВидСтроки", "СтрокаЛиста", "СортЛиста", "ПоказыватьСтатью", "БазаПроцента"):
    check(c in cols, f"колонка {c}")

# независимый пересчёт по Глобино-2
K = ("Ф1", "Ф2", "Ит")
acc = defaultdict(lambda: {"Ф1": 0.0, "Ф2": 0.0, "Ит": 0.0})
pok = {}
articles_seen = set()
base_seen = set()
for i in range(tab.Количество()):
    r = tab.Получить(i)
    if S(r.Подразделение) != "Глобино-2":
        continue
    vid = S(r.ВидСтроки)
    if vid != "Статья":
        pok[(int(num(r.СортЛиста)), S(r.СтрокаЛиста))] = (
            num(r.СуммаЕРПФ1), num(r.СуммаЕРПФ2), num(r.СуммаЕРП), vid)
        continue
    if S(r.СтатьяPL):
        articles_seen.add(S(r.СтатьяPL))
    base_seen.add(round(num(r.БазаПроцента), 2))
    f1, f2, it = num(r.СуммаЕРПФ1), num(r.СуммаЕРПФ2), num(r.СуммаЕРП)
    if f1 == 0 and f2 == 0 and it == 0:
        continue
    kс = st_code_by_ref.get(S(r.СтатьяPL))
    kг = gr_code_by_ref.get(S(r.Группа))
    key = None
    if kс in ("000000055", "000000056", "000000057", "000000058"):
        key = "ст" + kс
    elif kг:
        key = "гр" + kг
    if key:
        acc[key]["Ф1"] += f1
        acc[key]["Ф2"] += f2
        acc[key]["Ит"] += it

def wf(col):
    g = lambda k: acc["гр" + k][col]
    s = lambda k: acc["ст" + k][col]
    ОД = g("000000006")
    М = ОД - g("000000001")
    О = М - g("000000003") - g("000000005") - g("000000002")
    ПФ = О + s("000000055") - s("000000056")
    ДН = ПФ - g("000000007")
    Ч = ДН - g("000000008") - s("000000057")
    ВР = Ч - s("000000058")
    return {"Марж": М, "Опер": О, "ПослеФин": ПФ, "ДоНалогов": ДН, "Чистый": Ч, "ВРасп": ВР,
            "МаржП": (М / ОД if ОД else 0), "ОперП": (О / ОД if ОД else 0),
            "РентабП": (ВР / ОД if ОД else 0)}

w = {c: wf(c) for c in K}
EXP = {40: ("Марж", "Показатель"), 50: ("МаржП", "ПоказательПроцент"),
       100: ("Опер", "Показатель"), 110: ("ОперП", "ПоказательПроцент"),
       130: ("ПослеФин", "Показатель"), 150: ("ДоНалогов", "Показатель"),
       160: ("Чистый", "Показатель"), 180: ("ВРасп", "Показатель"),
       190: ("РентабП", "ПоказательПроцент")}
for (srt, name), (f1, f2, it, vid) in sorted(pok.items()):
    key, evid = EXP[srt]
    ok = (abs(f1 - w["Ф1"][key]) < 0.011 and abs(f2 - w["Ф2"][key]) < 0.011
          and abs(it - w["Ит"][key]) < 0.011 and vid == evid)
    check(ok, f"показатель {srt} {name}: Ф1={f1:,.2f}/Ф2={f2:,.2f}/Ит={it:,.2f}")
check(len(pok) == 9, f"строк-показателей 9 (={len(pok)})")

exp_base = sum(acc["гр" + c]["Ит"] for c in
               ("000000001", "000000007", "000000003", "000000005", "000000002", "000000008")) \
    + acc["ст000000056"]["Ит"] + acc["ст000000057"]["Ит"] + acc["ст000000058"]["Ит"]
check(round(exp_base, 2) in base_seen and 0.0 in base_seen,
      f"база %={exp_base:,.2f} (+0 у ОД); в строках={sorted(base_seen)}")
check(len(articles_seen) == len(catalog),
      f"полный шаблон: статей у Глобино-2 {len(articles_seen)} из {len(catalog)}")

# ============ 2. ПолныйШаблон=Ложь ============
tab2 = otchet.ПолучитьОбъединенныеДанные(D1, D2, False, False, True, False, True, False)
arts2 = set()
for i in range(tab2.Количество()):
    r = tab2.Получить(i)
    if S(r.Подразделение) == "Глобино-2" and S(r.ВидСтроки) == "Статья" and S(r.СтатьяPL):
        arts2.add(S(r.СтатьяPL))
check(len(arts2) < len(catalog), f"ПолныйШаблон=Ложь: статей {len(arts2)} < {len(catalog)}")

# ============ 3. Компоновка варианта ============
RANGES = {"000000006": range(6, 8), "000000001": range(9, 34), "000000007": range(36, 37),
          "000000003": range(42, 49), "000000005": range(51, 53), "000000002": range(55, 72),
          "000000008": range(74, 75)}
wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
ws = wb["Глобино-2"]
cells = {}
for idx, row in enumerate(ws.iter_rows(min_col=2, max_col=2, max_row=95), start=1):
    v = row[0].value
    cells[idx] = v.strip() if isinstance(v, str) else ""
excel_arts = {c: [cells[r] for r in rng if cells.get(r)] for c, rng in RANGES.items()}

GROUP_NAMES = {"000000006": "Операционный доход",
               "000000001": "Себестоимость проданой продукции (переменные производственные затраты)",
               "000000007": "Дополнительные  расходы", "000000003": "Общепроизводственные затраты",
               "000000005": "Маркетинговые затраты", "000000002": "Административные затраты",
               "000000008": "Налоги и сборы", "000000004": "Финансовая деятельность"}

def group_block(gcode):
    """Имя группы + статьи: Excel-порядок, затем хвост каталога по Сорт (без 057/058 для ФД)."""
    rows = [GROUP_NAMES[gcode]]
    excl = {"000000057", "000000058"}
    cat = [c for c in catalog if c["gcode"] == gcode and c["code"] not in excl]
    if gcode in excel_arts:
        exc = excel_arts[gcode]
        rows += exc
        exc_low = {e.lower() for e in exc}
        rows += [c["name"] for c in cat if c["name"].lower() not in exc_low]
    else:  # ФД: ФинДоход, ФинРасходы + хвост (070)
        rows += [c["name"] for c in cat]
    return rows

name_057 = next(c["name"] for c in catalog if c["code"] == "000000057")
name_058 = next(c["name"] for c in catalog if c["code"] == "000000058")
expected = (group_block("000000006") + group_block("000000001") + group_block("000000007")
            + ["Маржинальный доход, грн", "Маржинальный доход, %"]
            + group_block("000000003") + group_block("000000005") + group_block("000000002")
            + group_block("000000008")
            + ["Операционная прибыль, грн", "Операционная прибыль, %"]
            + group_block("000000004")
            + ["Прибыль после вычета финансовых затрат", name_057,
               "Прибыль до вычета налогов", "Чистый доход", name_058,
               "Прибыль в распоряжении компании (нераспределенная прибыль), грн",
               "Рентабельность продукции, %"])

схема = otchet.СхемаКомпоновкиДанных
вариант = None
for i in range(схема.ВариантыНастроек.Количество()):
    v = схема.ВариантыНастроек.Получить(i)
    if v.Имя == "А_ОтчетPL_ВидКакВExcel":
        вариант = v
check(вариант is not None, "вариант найден")
otchet.КомпоновщикНастроек.ЗагрузитьНастройки(вариант.Настройки)
настройки = otchet.КомпоновщикНастроек.Настройки
период = erp.NewObject("СтандартныйПериод")
период.ДатаНачала = datetime.datetime(2026, 6, 2)
период.ДатаОкончания = datetime.datetime(2026, 6, 29)
настройки.ПараметрыДанных.УстановитьЗначениеПараметра("Период", период)
qг = erp.NewObject("Запрос")
qг.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.СтруктураПредприятия ГДЕ Наименование = "Глобино-2" И НЕ ПометкаУдаления"""
глобино = qг.Execute().Выгрузить().Получить(0).Ссылка
отбор = настройки.Отбор.Элементы.Получить(0)
отбор.ПравоеЗначение = глобино
отбор.Использование = True

табдок = erp.NewObject("ТабличныйДокумент")
otchet.СкомпоноватьРезультат(табдок, None)
h, wd = табдок.ВысотаТаблицы, табдок.ШиринаТаблицы
rows = []
for rr in range(1, h + 1):
    vals = [табдок.Область(rr, cc, rr, cc).Текст for cc in range(1, min(wd, 6) + 1)]
    rows.append(vals)

# найти строку "Глобино-2", дальше — последовательность строк листа
start = next(i for i, v in enumerate(rows) if v[0].strip() == "Глобино-2") + 1
actual = [v for v in rows[start:] if v[0].strip()]
actual_names = [v[0].strip() for v in actual]
# допускаем опциональный блок «Вне структуры PL» в конце
if "Вне структуры PL" in actual_names:
    cut = actual_names.index("Вне структуры PL")
    extra = actual_names[cut:]
    actual_main = actual_names[:cut]
    print(f"  (в конце присутствует блок: {extra})")
else:
    actual_main = actual_names

seq_ok = actual_main == expected
check(seq_ok, f"последовательность строк листа 1:1 ({len(actual_main)} vs {len(expected)})")
if not seq_ok:
    for i in range(max(len(actual_main), len(expected))):
        a = actual_main[i] if i < len(actual_main) else "<нет>"
        e = expected[i] if i < len(expected) else "<нет>"
        if a != e:
            print(f"   !! poz {i}: отчёт={a!r} эталон={e!r}")

by_name = {v[0].strip(): v for v in actual}
m = by_name.get("Маржинальный доход, %")
check(m is not None and any("%" in x for x in m[1:4]),
      f"«Маржинальный доход, %» со знаком %: {m[1:4] if m else None}")
r_ren = by_name.get("Рентабельность продукции, %")
check(r_ren is not None and any("%" in x for x in r_ren[1:4]),
      f"«Рентабельность продукции, %» со знаком %: {r_ren[1:4] if r_ren else None}")
zp = by_name.get("ЗП производственного персонала")
check(zp is not None and zp[4].endswith("%"), f"колонка % статьи (ЗП произв.): {zp[4] if zp else None!r}")
od = by_name.get("Операционный доход")
check(od is not None and od[4] == "", f"% у ОД пуст: {od[4] if od else None!r}")
mg = by_name.get("Маржинальный доход, грн")
check(mg is not None and mg[3] != "", f"«Маржинальный доход, грн» Итого: {mg[3] if mg else None!r}")

# --- v3: суммы родителей = показатели (у Глобино-2 налог/дивиденды=0 => ПослеФин=ДоНалогов=Чистый=ВРасп)
pf = by_name.get("Прибыль после вычета финансовых затрат")
dn = by_name.get("Прибыль до вычета налогов")
ch = by_name.get("Чистый доход")
vr = by_name.get("Прибыль в распоряжении компании (нераспределенная прибыль), грн")
vals = [x[3] for x in (pf, dn, ch, vr) if x]
check(len(vals) == 4 and vals[0] != "" and len(set(vals)) == 1,
      f"родители = показатели (ПослеФин=ДоНалогов=Чистый=ВРасп): {vals}")

# --- v3: «0%» вместо голого «%» у нулевой статьи
ju = by_name.get("Юридические услуги")
check(ju is not None and ju[4] == "0%", f"нулевая статья %='0%': {ju[4] if ju else None!r}")
zero_pct = [n for n, v in by_name.items() if len(v) > 4 and v[4] == "%"]
check(not zero_pct, f"голых '%' нет: {zero_pct[:5]}")

# --- v3: иерархия — 057/058 идут строками сразу после родителей (подстроки)
i_pf = actual_names.index("Прибыль после вычета финансовых затрат")
check(actual_names[i_pf + 1] == name_057, f"057 сразу после ПослеФин: {actual_names[i_pf + 1]!r}")
i_ch = actual_names.index("Чистый доход")
check(actual_names[i_ch + 1] == name_058, f"058 сразу после Чистого дохода: {actual_names[i_ch + 1]!r}")

# --- v3: цвета заливки уровня-1
def bg(row_name):
    rr = start + 1 + actual_names.index(row_name)
    obl = табдок.Область(rr, 1, rr, 1)
    c = obl.ЦветФона
    try:
        return (int(c.Красный), int(c.Зеленый), int(c.Синий))
    except Exception:
        return ("?", str(c))

for row_name, exp in (("Налоги и сборы", (217, 217, 217)),
                      ("Маржинальный доход, грн", (189, 215, 238)),
                      ("Маржинальный доход, %", (198, 224, 180))):
    got = bg(row_name)
    check(got == exp, f"цвет «{row_name}»: {got} (эталон {exp})")

# ============ 4. Регрессия ============
tot_erp = tot_pl = 0.0
for i in range(tab2.Количество()):
    r = tab2.Получить(i)
    if S(r.ВидСтроки) == "Статья":
        tot_erp += num(r.СуммаЕРП)
        tot_pl += num(r.СуммаPL)
# tab2 собран с ПоказPL=Ложь -> СуммаPL=0; свод ЕРП сравним
check(abs(tot_erp - 149003872.98) < 0.5, f"свод СуммаЕРП={tot_erp:,.2f} (эталон 149 003 872,98)")

for имя in ("Основной", "А_ОтчетPL_РасшифровкаДокументаОтчетPL", "А_ОтчетPL_РасшифровкаПоНаправлению"):
    o2 = erp.Отчеты.А_ОтчетPL.Создать()
    for i in range(o2.СхемаКомпоновкиДанных.ВариантыНастроек.Количество()):
        v = o2.СхемаКомпоновкиДанных.ВариантыНастроек.Получить(i)
        if v.Имя == имя:
            o2.КомпоновщикНастроек.ЗагрузитьНастройки(v.Настройки)
    п = erp.NewObject("СтандартныйПериод")
    п.ДатаНачала = datetime.datetime(2026, 6, 2)
    п.ДатаОкончания = datetime.datetime(2026, 6, 29)
    o2.КомпоновщикНастроек.Настройки.ПараметрыДанных.УстановитьЗначениеПараметра("Период", п)
    td = erp.NewObject("ТабличныйДокумент")
    try:
        o2.СкомпоноватьРезультат(td, None)
        check(td.ВысотаТаблицы > 5, f"вариант «{имя}»: {td.ВысотаТаблицы}x{td.ШиринаТаблицы}")
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        check(False, f"вариант «{имя}»: {msg}")

print()
print("ИТОГ:", "ВСЁ OK" if not FAILS else f"FAIL x{len(FAILS)}")
sys.exit(1 if FAILS else 0)
