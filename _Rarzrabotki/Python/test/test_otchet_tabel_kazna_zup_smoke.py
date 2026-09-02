# -*- coding: utf-8 -*-
"""Смок собранного .erf «Табели Казны и ЗУП» в базе kazna (август 2026).

1. .erf загружается (ВнешниеОтчеты.Создать, БезопасныйРежим = Ложь), основная форма = ОбщаяФорма.ФормаОтчета.
2. Вариант «Сводно по сотрудникам»: строк = сотрудники Казны + ЗУП-люди без табеля (независимый расчёт
   по тем же запросам); строка Чешко С.А.: Дни М = 3, Часы М = 24, Дни работа ЗУП = 21, Часы работа ЗУП = 168,
   Расхождение дней = Р + М + К − дни работы ЗУП.
3. Вариант «Контроль»: у Штиченко/Бутко/Мусійченко/Заєць текст «Нет флага «Трудоустроен»».
4. Вариант «По дням»: компонуется, у Чешко есть строки 11–13.08 со значением «8М».
Печать PASS/FAIL, exit code 1 при провале.
"""
import sys
import os
import datetime
import pywintypes
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
ERF = os.path.normpath(os.path.join(HERE, "..", "..", "Отчеты", "ОтчетПоДаннымТабелейКазныиБазыЗУП.erf"))
CONN_KZ = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'
CONN_ZUP = 'Srvr="localhost";Ref="zup";Usr="cfo";Pwd="2442"'
Y, M = 2026, 8

fails = 0


def check(cond, msg):
    global fails
    print(("PASS " if cond else "FAIL ") + msg)
    if not cond:
        fails += 1


def rows(t):
    cols = [t.Колонки.Получить(i).Имя for i in range(t.Колонки.Количество())]
    return [{c: getattr(t.Получить(i), c) for c in cols} for i in range(t.Количество())]


def run(base, txt, params=None):
    q = base.NewObject("Запрос")
    q.Text = txt
    for k, v in (params or {}).items():
        q.SetParameter(k, v)
    return rows(q.Execute().Выгрузить())


def bounds(base):
    q = base.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ НАЧАЛОПЕРИОДА(&Н, ДЕНЬ) КАК Н, КОНЕЦПЕРИОДА(&К, ДЕНЬ) КАК К"
    q.SetParameter("Н", pywintypes.Time(datetime.datetime(Y, M, 1, 12, 0, 0)))
    q.SetParameter("К", pywintypes.Time(datetime.datetime(Y, M, 31, 12, 0, 0)))
    r = q.Execute().Выгрузить().Получить(0)
    return r.Н, r.К


def expected(kz, zup):
    """Независимый расчёт: множества сотрудников Казны / ЗУП и флаг «нет Трудоустроен»."""
    S = kz.String
    n, k = bounds(kz)
    kzr = run(kz, """ВЫБРАТЬ РАЗЛИЧНЫЕ Т.Сотрудник КАК Сотрудник, Т.Сотрудник.Наименование КАК ФИО, Т.Сотрудник.ИНН КАК ИНН, Т.Сотрудник.Трудоустроен КАК Труд
        ИЗ РегистрНакопления.ДанныеТабелирования КАК Т ГДЕ Т.Период МЕЖДУ &Н И &К И Т.Активность""", {"Н": n, "К": k})
    kz_emp = {S(r["Сотрудник"].УникальныйИдентификатор()): (S(r["ФИО"]), S(r["ИНН"]).strip(), bool(r["Труд"])) for r in kzr}
    kz_inn = {v[1] for v in kz_emp.values() if v[1]}
    # строки Свода: сотрудники с одним ИНН сливаются; Трудоустроен = ИЛИ по слитым записям
    groups = {}
    for uid, (fio, inn, tr) in kz_emp.items():
        groups.setdefault(inn or ("К:" + uid), []).append((fio, inn, tr))
    nz, kk = bounds(zup)
    zt = run(zup, """ВЫБРАТЬ РВ.Сотрудник.Физлицо.КодПоДРФО КАК ДРФО, СУММА(ВЫБОР КОГДА РВ.ВидИспользованияРабочегоВремени.РабочееВремя ТОГДА РВ.Дней ИНАЧЕ 0 КОНЕЦ) КАК ДниРабота
        ИЗ РегистрНакопления.РабочееВремяРаботниковОрганизаций КАК РВ ГДЕ РВ.Период МЕЖДУ &Н И &К И РВ.Активность
        СГРУППИРОВАТЬ ПО РВ.Сотрудник.Физлицо.КодПоДРФО""", {"Н": nz, "К": kk})
    zn = run(zup, """ВЫБРАТЬ Н.Сотрудник.Физлицо.КодПоДРФО КАК ДРФО, СУММА(Н.Результат) КАК Начислено
        ИЗ РегистрРасчета.ОсновныеНачисленияРаботниковОрганизаций КАК Н WHERE Н.ПериодРегистрации МЕЖДУ &Н И &К
        СГРУППИРОВАТЬ ПО Н.Сотрудник.Физлицо.КодПоДРФО""".replace("WHERE", "ГДЕ"), {"Н": nz, "К": kk})
    z_work = {zup.String(r["ДРФО"]).strip(): float(r["ДниРабота"]) for r in zt if zup.String(r["ДРФО"]).strip()}
    z_nach = {zup.String(r["ДРФО"]).strip(): float(r["Начислено"]) for r in zn if zup.String(r["ДРФО"]).strip()}
    zup_only = {d for d in set(z_work) | {d for d, s in z_nach.items() if s != 0} if d not in kz_inn}
    no_flag = sorted(g[0][0] for key, g in groups.items() if g[0][1] and not any(tr for _, _, tr in g) and (z_nach.get(g[0][1], 0) != 0 or z_work.get(g[0][1], 0) > 0))
    print(f"   Казна: записей {len(kz_emp)}, строк после слияния по ИНН {len(groups)}, дублей ИНН {sum(1 for g in groups.values() if len(g) > 1)}")
    return len(groups), len(zup_only), no_flag


def compose(kz, report, variant_name):
    схема = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
    вариант = схема.ВариантыНастроек.Найти(variant_name)
    assert вариант is not None, "вариант %s не найден" % variant_name
    report.КомпоновщикНастроек.ЗагрузитьНастройки(вариант.Настройки)
    период = kz.NewObject("СтандартныйПериод")
    период.ДатаНачала = pywintypes.Time(datetime.datetime(Y, M, 1, 12, 0, 0))
    период.ДатаОкончания = pywintypes.Time(datetime.datetime(Y, M, 31, 12, 0, 0))
    report.КомпоновщикНастроек.Настройки.ПараметрыДанных.УстановитьЗначениеПараметра("Период", период)
    табдок = kz.NewObject("ТабличныйДокумент")
    try:
        report.СкомпоноватьРезультат(табдок, None)
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print("   COMPOSE FAIL [%s]: %s" % (variant_name, msg))
        return [[]]
    h, w = табдок.ВысотаТаблицы, табдок.ШиринаТаблицы
    grid = []
    for r in range(1, h + 1):
        grid.append([(табдок.Область(r, c, r, c).Текст or "") for c in range(1, w + 1)])
    return grid


def num(s):
    t = (s or "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def table_rows(grid):
    """Строки данных: после строки заголовка (содержит «Сотрудник» и «Контроль» или «Казна»), до «Итого»."""
    hdr = None
    for i, row in enumerate(grid):
        if "Сотрудник" in row and ("Контроль" in row or "Казна" in row):
            hdr = i
            break
    assert hdr is not None, "строка заголовка не найдена"
    header = grid[hdr]
    data = []
    for row in grid[hdr + 1:]:
        first = next((c for c in row if c), "")
        if first.startswith("Итого"):
            break
        if any(row):
            data.append(row)
    return header, data


def col(header, title):
    assert title in header, "нет колонки %s в %s" % (title, header)
    return header.index(title)


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kz = v8.Connect(CONN_KZ)
    zup = v8.Connect(CONN_ZUP)
    S = kz.String

    print("ERF:", ERF)
    report = kz.ВнешниеОтчеты.Создать(ERF, False)
    md = report.Метаданные()
    check(S(md.Имя) == "ОтчетПоДаннымТабелейКазныиБазыЗУП", "имя отчёта: " + S(md.Имя))
    check(S(md.ОсновнаяФорма.ПолноеИмя()) == "ОбщаяФорма.ФормаОтчета", "основная форма = ОбщаяФорма.ФормаОтчета (%s)" % S(md.ОсновнаяФорма.ПолноеИмя()))
    сведения = report.СведенияОВнешнейОбработке()
    check(S(сведения.Вид) == "ДополнительныйОтчет" and сведения.БезопасныйРежим is False, "СведенияОВнешнейОбработке: ДополнительныйОтчет, БезопасныйРежим = Ложь")

    n_kz, n_zup_only, no_flag = expected(kz, zup)
    print(f"   ожидание: сотрудников Казны {n_kz}, ЗУП без табеля {n_zup_only}, без флага Трудоустроен {len(no_flag)}: {no_flag}")

    # --- Свод ---
    grid = compose(kz, report, "СводноПоСотрудникам")
    check(len(grid) > 1, "вариант «Сводно по сотрудникам» скомпонован")
    header, data = table_rows(grid)
    print(f"   Свод: таблица {len(grid)} строк, данных {len(data)}, колонок {len(header)}")
    check(len(data) == n_kz + n_zup_only, f"строк Свода {len(data)} = Казна {n_kz} + ЗУП без табеля {n_zup_only}")
    c_fio = col(header, "Сотрудник")
    ch = [r for r in data if r[c_fio] == "Чешко Сергій Анатолійович"]
    check(len(ch) == 1, f"строка Чешко С.А. одна: {len(ch)}")
    if ch:
        r = ch[0]
        vals = {t: num(r[col(header, t)]) or 0 for t in ("Дни Р (по ставке)", "Дни К (командировка)", "Дни М (официально)", "Часы М (официально)", "Дни работа (ЗУП)", "Часы работа (ЗУП)", "Расхождение дней (Казна − ЗУП)", "Начислено (ЗУП)")}
        print("   Чешко:", vals, "| контроль:", r[col(header, "Контроль")])
        check(vals["Дни М (официально)"] == 3 and vals["Часы М (официально)"] == 24, "Чешко: Дни М = 3, Часы М = 24")
        check(vals["Дни работа (ЗУП)"] == 21 and vals["Часы работа (ЗУП)"] == 168, "Чешко: ЗУП 21 день / 168 часов")
        exp_diff = vals["Дни Р (по ставке)"] + vals["Дни М (официально)"] + vals["Дни К (командировка)"] - vals["Дни работа (ЗУП)"]
        check(vals["Расхождение дней (Казна − ЗУП)"] == exp_diff, f"Чешко: расхождение дней = Р+М+К − ЗУП = {exp_diff}")
        check(vals["Начислено (ЗУП)"] == 29500, "Чешко: начислено ЗУП = 29 500")
    zup_only_rows = [r for r in data if not r[col(header, "ИНН")] == "" and "табеля Казны нет" in r[col(header, "Контроль")]]
    check(len(zup_only_rows) == n_zup_only, f"строк «В ЗУП есть, табеля Казны нет» = {len(zup_only_rows)}")

    # --- Контроль ---
    grid = compose(kz, report, "Контроль")
    check(len(grid) > 1, "вариант «Контроль» скомпонован")
    header, data = table_rows(grid)
    c_fio, c_ctl = col(header, "Сотрудник"), col(header, "Контроль")
    print(f"   Контроль: строк {len(data)}")
    check(all(r[c_ctl] for r in data), "в варианте «Контроль» у каждой строки заполнен Контроль")
    got_no_flag = sorted(r[c_fio] for r in data if "Нет флага «Трудоустроен»" in r[c_ctl])
    check(got_no_flag == no_flag, f"без флага Трудоустроен: {got_no_flag}")

    # --- По дням ---
    grid = compose(kz, report, "ПоДням")
    flat = ["|".join(r) for r in grid]
    ch_days = [l for l in flat if "8М" in l and ("11.08.2026" in l or "12.08.2026" in l or "13.08.2026" in l)]
    print(f"   По дням: таблица {len(grid)} строк; строк с «8М» 11–13.08: {len(ch_days)}")
    check(len(grid) > 50, "вариант «По дням» скомпонован")
    check(len(ch_days) >= 3, "по дням есть строки «8М» за 11–13.08 (Чешко и др.)")

    print(f"\nИтого провалов: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
