# -*- coding: utf-8 -*-
"""Смок собранного .erf «Табели Казны и ЗУП» в базе kazna (август 2026).

1. .erf загружается (ВнешниеОтчеты.Создать, БезопасныйРежим = Ложь), основная форма = ОбщаяФорма.ФормаОтчета.
2. Вариант «Сводно по сотрудникам»: строк = сотрудники Казны + ЗУП-люди без табеля (независимый расчёт
   по тем же запросам); строка Чешко С.А.: дни/часы Р и М как в регистре, официальные = Р + М, ЗУП 21 день / 168 часов,
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
    # актуальность ПО РЕГИСТРУ (как ЛичнаяКарточка): есть движения и последнее не увольнение; только трудовые договоры
    za = run(zup, """ВЫБРАТЬ РАЗЛИЧНЫЕ С.Физлицо.КодПоДРФО КАК ДРФО ИЗ Справочник.СотрудникиОрганизаций КАК С
        ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.РаботникиОрганизаций.СрезПоследних(, ) КАК Посл ПО Посл.Сотрудник = С.Ссылка
        ГДЕ НЕ С.ПометкаУдаления И С.Физлицо.КодПоДРФО <> "" И С.ВидДоговора = ЗНАЧЕНИЕ(Перечисление.ВидыДоговоровСФизЛицами.ТрудовойДоговор)
          И НЕ Посл.Регистратор ССЫЛКА Документ.УвольнениеИзОрганизаций""")
    actual = {zup.String(r["ДРФО"]).strip() for r in za}
    zup_only = {d for d in set(z_work) | {d for d, s in z_nach.items() if s != 0} | actual if d not in kz_inn}
    no_flag = sorted(g[0][0] for key, g in groups.items() if g[0][1] and not any(tr for _, _, tr in g) and (z_nach.get(g[0][1], 0) != 0 or z_work.get(g[0][1], 0) > 0 or g[0][1] in actual))
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
    start = hdr + 1
    header2 = None
    if start < len(grid) and "Организация ЗУП" in grid[start]:
        header2 = grid[start]
        start += 1
    data = []
    for row in grid[start:]:
        first = next((c for c in row if c), "")
        if first.startswith("Итого"):
            break
        if any(row):
            data.append(row)
    table_rows.header2 = header2
    if header2 is not None:
        # строки организаций: первая колонка = организация ЗУП («ТОВ …»); строки сотрудников — остальные
        table_rows.all_rows = data
        return header, [r for r in data if not is_org_row(r)]
    return header, data


def is_org_row(row):
    first = next((c for c in row if c), "")
    return first.startswith("ТОВ") or first.startswith("ФОП")


def org_rows_after(all_rows, i_emp):
    out = []
    for rr in all_rows[i_emp + 1:]:
        if not is_org_row(rr):
            break
        out.append(rr)
    return out


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
        vals = {t: num(r[col(header, t)]) or 0 for t in ("Дни Р (по ставке)", "Дни К (командировка)", "Дни М (только буква М)", "Часы М (только буква М)", "Дни официальные (Р+М)", "Часы официальные (Р+М)", "Дни работа (ЗУП)", "Часы работа (ЗУП)", "Расхождение дней (Казна − ЗУП)", "Начислено (ЗУП)")}
        print("   Чешко:", vals, "| контроль:", r[col(header, "Контроль")])
        # ожидание по регистру Казны (табели могут быть перепроведены пользователем)
        n, k = bounds(kz)
        reg = run(kz, """ВЫБРАТЬ Т.Состояние КАК Состояние, КОЛИЧЕСТВО(РАЗЛИЧНЫЕ НАЧАЛОПЕРИОДА(Т.Период, ДЕНЬ)) КАК Дней,
                СУММА(Т.ЧасыФакт) КАК Факт, СУММА(Т.ЧасыОфициальные) КАК Офиц
            ИЗ РегистрНакопления.ДанныеТабелирования КАК Т
            ГДЕ Т.Период МЕЖДУ &Н И &К И Т.Активность И Т.Сотрудник.ИНН = "3406013131" И (Т.Состояние = "М" ИЛИ (Т.Состояние = "Р" И Т.ЧасыФакт > 0))
            СГРУППИРОВАТЬ ПО Т.Состояние""", {"Н": n, "К": k})
        reg = {S(x["Состояние"]): (float(x["Дней"]), float(x["Факт"]), float(x["Офиц"])) for x in reg}
        exp_m_days, _, exp_m_hours = reg.get("М", (0, 0, 0))
        exp_r_days, exp_r_hours, _ = reg.get("Р", (0, 0, 0))
        print(f"   регистр Чешко: М {exp_m_days} дн / {exp_m_hours} ч, Р {exp_r_days} дн / {exp_r_hours} ч")
        check(vals["Дни М (только буква М)"] == exp_m_days and vals["Часы М (только буква М)"] == exp_m_hours, f"Чешко: Дни М = {exp_m_days}, Часы М = {exp_m_hours} (по регистру)")
        check(vals["Дни Р (по ставке)"] == exp_r_days and num(r[col(header, "Часы Р (по ставке)")]) == exp_r_hours, f"Чешко: Дни Р = {exp_r_days}, Часы Р = {exp_r_hours} (по регистру)")
        check(vals["Дни официальные (Р+М)"] == vals["Дни Р (по ставке)"] + vals["Дни М (только буква М)"] and vals["Часы официальные (Р+М)"] == vals["Часы М (только буква М)"] + num(r[col(header, "Часы Р (по ставке)")]),
              f"Чешко: официальные = Р + М (дни {vals['Дни официальные (Р+М)']}, часы {vals['Часы официальные (Р+М)']})")
        check(vals["Дни работа (ЗУП)"] == 21 and vals["Часы работа (ЗУП)"] == 168, "Чешко: ЗУП 21 день / 168 часов")
        exp_diff = vals["Дни официальные (Р+М)"] + vals["Дни К (командировка)"] - vals["Дни работа (ЗУП)"]
        check(vals["Расхождение дней (Казна − ЗУП)"] == exp_diff, f"Чешко: расхождение дней = официальные+К − ЗУП = {exp_diff}")
        check(vals["Начислено (ЗУП)"] == 29500, "Чешко: начислено ЗУП = 29 500")
        print("   Чешко приём/увольнение:", r[col(header, "Дата приёма (ЗУП)")], "|", r[col(header, "Дата увольнения (ЗУП)")])
        check(r[col(header, "Дата приёма (ЗУП)")] == "23.07.2025" and r[col(header, "Дата увольнения (ЗУП)")] == "", "Чешко: приём 23.07.2025 (актуальная запись), увольнения нет")
    zup_only_rows = [r for r in data if not r[col(header, "ИНН")] == "" and "табеля Казны нет" in r[col(header, "Контроль")]]
    check(len(zup_only_rows) == n_zup_only, f"строк «В ЗУП есть, табеля Казны нет» = {len(zup_only_rows)}")

    # --- уровень организаций ЗУП под сотрудником (связанный набор СводОрг) ---
    all_rows = getattr(table_rows, "all_rows", data)
    h2 = table_rows.header2
    check(h2 is not None, "есть вторая строка заголовка с колонками организаций ЗУП")
    ci_org = col(h2, "Организация ЗУП")
    i_ch = next(i for i, rr in enumerate(all_rows) if rr[c_fio] == "Чешко Сергій Анатолійович")
    org_rows = org_rows_after(all_rows, i_ch)
    print("   строки организаций Чешко:", [(rr[ci_org], rr[col(h2, "Дни работа (орг)")], rr[col(h2, "Часы работа (орг)")], rr[col(h2, "Начислено (орг)")], rr[col(h2, "Приём в организацию")]) for rr in org_rows])
    ind = [rr for rr in org_rows if "ІНДАСТРІАЛБУД" in rr[ci_org]]
    stl = [rr for rr in org_rows if "ІНДЕПТ СТІЛ" in rr[ci_org]]
    check(len(org_rows) == 2 and len(ind) == 1 and len(stl) == 1, "Чешко: две организации ЗУП под сотрудником (действующая и уволенная)")
    if ind:
        check(num(ind[0][col(h2, "Дни работа (орг)")]) == 21 and num(ind[0][col(h2, "Часы работа (орг)")]) == 168 and num(ind[0][col(h2, "Начислено (орг)")]) == 29500, "Чешко/ІНДАСТРІАЛБУД: 21 дн / 168 ч / 29 500")
        check(ind[0][col(h2, "Приём в организацию")] == "23.07.2025" and ind[0][col(h2, "Договор (ЗУП)")] == "Трудовой" and ind[0][col(h2, "Источник даты приёма")] == "" and ind[0][col(h2, "Увольнение из организации")] == "", "Чешко/ІНДАСТРІАЛБУД: приём 23.07.2025 по кадровому регистру, трудовой договор, не уволен")
    if stl:
        check(stl[0][col(h2, "Приём в организацию")] == "07.05.2025" and stl[0][col(h2, "Увольнение из организации")] == "22.07.2025", "Чешко/ІНДЕПТ СТІЛ: приём 07.05.2025, увольнение 22.07.2025 по регистру")
    n_org_rows = sum(1 for rr in all_rows if is_org_row(rr))
    print(f"   всего строк организаций: {n_org_rows}")
    check(n_org_rows >= len(data) - 150, "строки организаций есть у большинства сотрудников")
    check(not any(t in header for t in ("Есть контроль", "Нет флага Трудоустроен", "Нет табеля Казны", "Организации (Казна)")), "служебные флаги и «Организации (Казна)» в выводе не показаны")

    # --- Беспалько: трудовой договор без приказа о приёме, дата из «даты начала», плюс договор ГПХ ---
    bes = [r for r in data if r[col(header, "ИНН")] == "2627718590"]
    check(len(bes) == 1, f"Беспалько О.Ю. в Своде одной строкой: {len(bes)}")
    if bes:
        r = bes[0]
        i_b = next(i for i, rr in enumerate(all_rows) if rr[col(header, "ИНН")] == "2627718590")
        b_org = org_rows_after(all_rows, i_b)
        txt = "; ".join(rr[ci_org] + ": " + rr[col(h2, "Договор (ЗУП)")] + " / " + rr[col(h2, "Источник даты приёма")] for rr in b_org)
        print("   Беспалько:", r[col(header, "Дата приёма (ЗУП)")], "|", r[col(header, "Актуален в ЗУП (по кадровому регистру)")], "|", txt, "| контроль:", r[col(header, "Контроль")])
        check(r[col(header, "Дата приёма (ЗУП)")] == "24.04.2025" and r[col(header, "Актуален в ЗУП (по кадровому регистру)")] == "Нет", "Беспалько: дата приёма 24.04.2025 из приказа, по регистру НЕ актуален")
        check("приказ не проведён" in txt and "ГПХ" in txt, "Беспалько: у организации источник «приказ не проведён», договор «Трудовой + ГПХ»")
        check("Приказ о приёме в ЗУП не проведён" in r[col(header, "Контроль")] and "ІНД00075" in r[col(header, "Контроль")] and "помечен на удаление" in r[col(header, "Контроль")], "Беспалько: контроль «приказ ІНД00075 не проведён, помечен на удаление»")

    # --- Сотрудник ЗУП в двух организациях: дни по уникальным датам, часы/начисления суммой ---
    nz, kk = bounds(zup)
    multi = run(zup, """ВЫБРАТЬ ПЕРВЫЕ 1 Вл.ДРФО КАК ДРФО, Вл.Орг КАК Орг, Вл.Дней КАК Дней ИЗ (ВЫБРАТЬ РВ.Сотрудник.Физлицо.КодПоДРФО КАК ДРФО,
            КОЛИЧЕСТВО(РАЗЛИЧНЫЕ РВ.Организация) КАК Орг,
            КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫБОР КОГДА РВ.ВидИспользованияРабочегоВремени.РабочееВремя ТОГДА НАЧАЛОПЕРИОДА(РВ.Период, ДЕНЬ) КОНЕЦ) КАК Дней
        ИЗ РегистрНакопления.РабочееВремяРаботниковОрганизаций КАК РВ ГДЕ РВ.Период МЕЖДУ &Н И &К И РВ.Активность
        СГРУППИРОВАТЬ ПО РВ.Сотрудник.Физлицо.КодПоДРФО ИМЕЮЩИЕ КОЛИЧЕСТВО(РАЗЛИЧНЫЕ РВ.Организация) > 1) КАК Вл УПОРЯДОЧИТЬ ПО Вл.ДРФО""", {"Н": nz, "К": kk})
    if multi:
        drfo, n_org, n_days = zup.String(multi[0]["ДРФО"]).strip(), int(multi[0]["Орг"]), int(multi[0]["Дней"])
        mrow = [r for r in data if r[col(header, "ИНН")] == drfo]
        check(len(mrow) == 1, f"сотрудник ЗУП с {n_org} организациями (ДРФО {drfo}) в Своде одной строкой: {len(mrow)}")
        if mrow:
            i_m = next(i for i, rr in enumerate(all_rows) if rr[col(header, "ИНН")] == drfo)
            m_org = org_rows_after(all_rows, i_m)
            print("   мульти-орг организации:", [(rr[ci_org], rr[col(h2, "Дни работа (орг)")], rr[col(h2, "Начислено (орг)")]) for rr in m_org])
            check(len(m_org) == n_org, f"под сотрудником {n_org} строк организаций")
        if mrow:
            r = mrow[0]
            print("   мульти-орг:", r[c_fio], "| контроль:", r[col(header, "Контроль")])
            check(num(r[col(header, "Организаций (ЗУП)")]) == n_org, f"Организаций (ЗУП) = {n_org}")
            check(num(r[col(header, "Дни работа (ЗУП)")]) == n_days, f"Дни работа (ЗУП) по уникальным датам = {n_days} (не сумма по организациям)")
            check("несколько организаций" in r[col(header, "Контроль")], "в Контроле пометка про несколько организаций")
    else:
        print("   сотрудников ЗУП с 2+ организациями за период нет — проверка пропущена")

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
