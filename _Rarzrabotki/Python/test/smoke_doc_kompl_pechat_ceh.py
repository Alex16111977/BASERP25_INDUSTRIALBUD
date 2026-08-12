# -*- coding: utf-8 -*-
"""Smoke печати «Аналіз для цеху» (МакетАнализЦех) на живом доке №3 (дом №1).

Проверки: колонки, Бойлер (норма/економія только на группе, жирный), заливка только
понаднормовых, Разом = сохранённая ТЧ, заголовок, кросс-сверка норм с отчётом
А_ОтчетПоСписаниюНаПроизводствоБухгалтерский (ПланКол по исполнителю №1).
"""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты\А_ОтчетПоСписаниюНаПроизводствоБухгалтерский.erf"

v8 = win32com.client.Dispatch("V83.COMConnector")
try:
    erp = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
except Exception:
    erp = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = erp.String

FAILS = []
def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f" | {detail}" if detail else ""))
    if not cond:
        FAILS.append(name)

def num(txt):
    txt = (txt or "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None

q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д
ГДЕ Д.Номер = "00000000003" И НЕ Д.ПометкаУдаления"""
doc = q.Execute().Выгрузить().Получить(0).Ссылка.ПолучитьОбъект()
data = doc.ТабличнаяЧастьОстатков.Выгрузить()

td = doc.СформироватьПечатьАнализЦех(data)
check("печать сформирована", td is not None)
H, W = td.ВысотаТаблицы, td.ШиринаТаблицы
print(f"ТабДок {H}x{W}")

def cell(r, c):
    return (td.Область(r, c, r, c).Текст or "").strip()

# 1. Строка заголовков колонок и карта колонок
hdr_row, cols = None, {}
for r in range(1, min(H, 15) + 1):
    row_texts = {c: cell(r, c) for c in range(1, W + 1)}
    if "Видано" in row_texts.values():
        hdr_row = r
        for c, t in row_texts.items():
            if t:
                cols[t] = c
        break
check("строка заголовков найдена", hdr_row is not None, str(cols))
for t in ("Норма СС", "Видано", "Одиниця виміру", "В нормі", "Понад норму",
          "Економія", "Коментар цеху"):
    check(f"колонка {t!r}", t in cols)
for t in ("Згідно з СС", "В нормі, грн", "Факт. списання"):
    check(f"нет колонки {t!r}", t not in cols)

# 2. Заголовок печати
head = " ".join(cell(r, c) for r in range(1, (hdr_row or 8)) for c in range(1, W + 1) if cell(r, c))
check("заголовок: назва", "Аналіз залишків для списання за СС" in head)
check("заголовок: підрозділ", "МД IRS 2026" in head)
check("заголовок: склад №1", "15 м №1" in head)

def find_row(text, after=0):
    for r in range(max((hdr_row or 0) + 1, after + 1), H + 1):
        for c in range(1, min(W, 3) + 1):
            if cell(r, c) == text:
                return r
    return None

def bold(r):
    for c in range(1, min(W, 3) + 1):
        if cell(r, c):
            return bool(td.Область(r, c, r, c).Шрифт.Жирный)
    return False

def fill_rgb(r):
    a = td.Область(r, cols["Видано"], r, cols["Видано"])
    цв = a.ЦветФона
    return (цв.R, цв.G, цв.B)

# 3. Бойлер: группа vs карточка
rg = find_row("Бойлер електричний")
check("группа Бойлер найдена", rg is not None)
if rg:
    check("Бойлер: Норма СС=2", num(cell(rg, cols["Норма СС"])) == 2, cell(rg, cols["Норма СС"]))
    check("Бойлер: Видано=1", num(cell(rg, cols["Видано"])) == 1)
    check("Бойлер: В нормі=1", num(cell(rg, cols["В нормі"])) == 1)
    check("Бойлер: Економія=1", num(cell(rg, cols["Економія"])) == 1)
    check("Бойлер: группа жирная", bold(rg))
    rc = find_row("Електричний водонагрівач O`Pro Slim PC 30, 30л", rg)
    check("карточка водонагрівача найдена", rc is not None)
    if rc:
        check("карточка: Економія пусто", cell(rc, cols["Економія"]) == "")
        check("карточка: Норма СС пусто", cell(rc, cols["Норма СС"]) == "")
        check("карточка: не жирная", not bold(rc))
        check("карточка: Коментар цеху пуст", cell(rc, cols["Коментар цеху"]) == "")

# 4. Нормы других групп (для кросс-сверки)
group_norms = {}
for name in ("Вимикач", "Вітробарєр", "Хомут"):
    r = find_row(name)
    if r:
        group_norms[name] = num(cell(r, cols["Норма СС"]))
check("Вимикач: Норма СС=10", group_norms.get("Вимикач") == 10, str(group_norms.get("Вимикач")))
check("Вітробарєр: Норма СС=42", group_norms.get("Вітробарєр") == 42, str(group_norms.get("Вітробарєр")))
check("Хомут: Норма СС=14", group_norms.get("Хомут") == 14, str(group_norms.get("Хомут")))

# 5. Заливка: понаднормовая жёлтая, нормовая — нет
ra = find_row("АВР")
if ra:
    check("АВР (понад): жёлтая заливка", fill_rgb(ra) == (255, 230, 153), str(fill_rgb(ra)))
rb = find_row("Брус")
if rb:
    check("Брус (в нормі): НЕ жёлтая", fill_rgb(rb) != (255, 230, 153), str(fill_rgb(rb)))

# 6. Разом = сохранённая ТЧ
tot = {"ost": 0.0, "vn": 0.0, "pn": 0.0, "ek": 0.0}
tch = doc.ТабличнаяЧастьОстатков
for i in range(tch.Количество()):
    s = tch.Получить(i)
    tot["ost"] += s.Остаток; tot["vn"] += s.ВНорме
    tot["pn"] += s.ПонадНорму; tot["ek"] += s.Экономия
rr = None
for r in range(H, (hdr_row or 0), -1):
    # COM-сеанс ru -> «Итого», укр-клиент -> «Разом» (локализация платформы)
    if any(("Разом" in cell(r, c)) or ("Итого" in cell(r, c)) for c in range(1, min(W, 3) + 1)):
        rr = r
        break
check("строка Разом найдена", rr is not None)
if rr:
    for title, key in (("Видано", "ost"), ("В нормі", "vn"),
                       ("Понад норму", "pn"), ("Економія", "ek")):
        v = num(cell(rr, cols[title]))
        check(f"Разом {title} = ТЧ", v is not None and abs(v - tot[key]) < 0.01,
              f"{v} vs {tot[key]:.3f}")

# 7. Кросс-сверка норм с отчётом (План, кіл по исполнителю №1)
try:
    rep = erp.ВнешниеОтчеты.Создать(ERF)
    tz = rep.ПолучитьДанные(None)
    plan = {}
    for i in range(tz.Количество()):
        row = tz.Получить(i)
        if S(row.ВидДокумента) == "0. План" and "15 м №1" in S(row.ПодразделениеИсполнитель):
            key = S(row.ОбщееНазвание)
            plan[key] = plan.get(key, 0.0) + row.ПланКол
    for name, expect in (("Бойлер електричний", 2), ("Вимикач", 10), ("Вітробарєр", 42)):
        check(f"отчёт ПланКол {name}={expect}", abs(plan.get(name, 0) - expect) < 0.001,
              str(plan.get(name)))
    check("печать Норма == отчёт План (Вимикач)", group_norms.get("Вимикач") == plan.get("Вимикач"))
    check("печать Норма == отчёт План (Вітробарєр)", group_norms.get("Вітробарєр") == plan.get("Вітробарєр"))
except Exception as e:
    check("кросс-сверка с отчётом", False, repr(e)[:200])

print(f"\n{'='*50}\nИТОГ: {'ALL PASS' if not FAILS else 'FAILS: ' + ', '.join(FAILS)}")
sys.exit(0 if not FAILS else 1)
