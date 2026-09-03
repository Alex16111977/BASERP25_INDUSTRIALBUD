# -*- coding: utf-8 -*-
"""Смок собранного .erf «Табели Казны и ЗУП» в базе kazna (август 2026).

Варианты: «Сводно по сотрудникам» (сотрудник компактно + организации ЗУП, итоги), «Контроль»,
«Сотрудник → табели Казны», «Подробно», «По дням». Ожидания считаются из регистров Казны и ЗУП.
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
CHESHKO = "Чешко Сергій Анатолійович"
COMPACT = ["Сотрудник", "ИНН", "Трудоустроен (Казна)", "В ЗУП (актуален по регистру)", "Приём (ЗУП)", "Увольнение (ЗУП)",
           "Расхождение дней (Казна − ЗУП)", "Контроль", "Дни Казна (Р+М+К)", "Часы Казна (Р+М)", "Дни ЗУП (работа)", "Часы ЗУП (работа)"]

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


def num(s):
    t = (s or "").replace("\xa0", "").replace(" ", "").replace(",", ".").replace("−", "-")
    try:
        return float(t)
    except ValueError:
        return None


def expected(kz, zup):
    S = kz.String
    n, k = bounds(kz)
    kzr = run(kz, """ВЫБРАТЬ РАЗЛИЧНЫЕ Т.Сотрудник КАК Сотрудник, Т.Сотрудник.Наименование КАК ФИО, Т.Сотрудник.ИНН КАК ИНН, Т.Сотрудник.Трудоустроен КАК Труд
        ИЗ РегистрНакопления.ДанныеТабелирования КАК Т ГДЕ Т.Период МЕЖДУ &Н И &К И Т.Активность""", {"Н": n, "К": k})
    kz_emp = {S(r["Сотрудник"].УникальныйИдентификатор()): (S(r["ФИО"]), S(r["ИНН"]).strip(), bool(r["Труд"])) for r in kzr}
    kz_inn = {v[1] for v in kz_emp.values() if v[1]}
    groups = {}
    for uid, (fio, inn, tr) in kz_emp.items():
        groups.setdefault(inn or ("К:" + uid), []).append((fio, inn, tr))
    nz, kk = bounds(zup)
    zt = run(zup, """ВЫБРАТЬ РВ.Сотрудник.Физлицо.КодПоДРФО КАК ДРФО, СУММА(ВЫБОР КОГДА РВ.ВидИспользованияРабочегоВремени.РабочееВремя ТОГДА РВ.Дней ИНАЧЕ 0 КОНЕЦ) КАК ДниРабота
        ИЗ РегистрНакопления.РабочееВремяРаботниковОрганизаций КАК РВ ГДЕ РВ.Период МЕЖДУ &Н И &К И РВ.Активность
        СГРУППИРОВАТЬ ПО РВ.Сотрудник.Физлицо.КодПоДРФО""", {"Н": nz, "К": kk})
    zn = run(zup, """ВЫБРАТЬ Н.Сотрудник.Физлицо.КодПоДРФО КАК ДРФО, СУММА(Н.Результат) КАК Начислено
        ИЗ РегистрРасчета.ОсновныеНачисленияРаботниковОрганизаций КАК Н ГДЕ Н.ПериодРегистрации МЕЖДУ &Н И &К
        СГРУППИРОВАТЬ ПО Н.Сотрудник.Физлицо.КодПоДРФО""", {"Н": nz, "К": kk})
    za = run(zup, """ВЫБРАТЬ РАЗЛИЧНЫЕ С.Физлицо.КодПоДРФО КАК ДРФО ИЗ Справочник.СотрудникиОрганизаций КАК С
        ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.РаботникиОрганизаций.СрезПоследних(, ) КАК Посл ПО Посл.Сотрудник = С.Ссылка
        ГДЕ НЕ С.ПометкаУдаления И С.Физлицо.КодПоДРФО <> "" И С.ВидДоговора = ЗНАЧЕНИЕ(Перечисление.ВидыДоговоровСФизЛицами.ТрудовойДоговор)
          И НЕ Посл.Регистратор ССЫЛКА Документ.УвольнениеИзОрганизаций""")
    z_work = {zup.String(r["ДРФО"]).strip(): float(r["ДниРабота"]) for r in zt if zup.String(r["ДРФО"]).strip()}
    z_nach = {zup.String(r["ДРФО"]).strip(): float(r["Начислено"]) for r in zn if zup.String(r["ДРФО"]).strip()}
    actual = {zup.String(r["ДРФО"]).strip() for r in za}
    zup_only = {d for d in set(z_work) | {d for d, s in z_nach.items() if s != 0} | actual if d not in kz_inn}
    no_flag = sorted(g[0][0] for key, g in groups.items() if g[0][1] and not any(tr for _, _, tr in g)
                     and (z_nach.get(g[0][1], 0) != 0 or z_work.get(g[0][1], 0) > 0 or g[0][1] in actual))
    print(f"   Казна: записей {len(kz_emp)}, строк после слияния по ИНН {len(groups)}")
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
    return [[(табдок.Область(r, c, r, c).Текст or "") for c in range(1, w + 1)] for r in range(1, h + 1)]


def parse(grid):
    """Заголовки (по строке на уровень), строки данных до «Итого», строка итогов."""
    hdr = next(i for i, row in enumerate(grid) if "Сотрудник" in row and "Контроль" in row)
    headers = [grid[hdr]]
    i = hdr + 1
    while i < len(grid) and ("Организация ЗУП" in grid[i] or "Табель Казны" in grid[i]):
        headers.append(grid[i]); i += 1
    data, total = [], None
    for row in grid[i:]:
        f = first(row)
        if f.startswith("Итого"):
            total = row
            break
        if any(row):
            data.append(row)
    return headers, data, total


def col(header, title):
    assert title in header, "нет колонки %s в %s" % (title, header)
    return header.index(title)


def first(row):
    return next((c for c in row if c), "")


def is_child(row):
    f = first(row)
    return f.startswith("ТОВ") or f.startswith("ФОП") or f.startswith("Табель учета")


def children_after(data, i):
    out = []
    for rr in data[i + 1:]:
        if not is_child(rr):
            break
        out.append(rr)
    return out


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kz = v8.Connect(CONN_KZ)
    zup = v8.Connect(CONN_ZUP)
    S = kz.String

    print("ERF:", ERF)
    report = kz.ВнешниеОтчеты.Создать(ERF, False)
    md = report.Метаданные()
    check(S(md.Имя) == "ОтчетПоДаннымТабелейКазныиБазыЗУП", "имя отчёта: " + S(md.Имя))
    check(S(md.ОсновнаяФорма.ПолноеИмя()) == "ОбщаяФорма.ФормаОтчета", "основная форма = ОбщаяФорма.ФормаОтчета")
    схема0 = report.ПолучитьМакет("ОсновнаяСхемаКомпоновкиДанных")
    имена_параметров = [S(схема0.Параметры.Получить(i).Имя) for i in range(схема0.Параметры.Количество())]
    check(имена_параметров == ["Период"], f"параметры СКД только Период: {имена_параметров}")
    варианты = [S(схема0.ВариантыНастроек.Получить(i).Имя) for i in range(схема0.ВариантыНастроек.Количество())]
    check(варианты == ["СводноПоСотрудникам", "Контроль", "ПоТабелямКазны", "Подробно", "ПоДням"], f"варианты: {варианты}")
    сведения = report.СведенияОВнешнейОбработке()
    check(S(сведения.Вид) == "ДополнительныйОтчет" and сведения.БезопасныйРежим is False, "СведенияОВнешнейОбработке: ДополнительныйОтчет, БезопасныйРежим = Ложь")

    n_kz, n_zup_only, no_flag = expected(kz, zup)
    print(f"   ожидание: сотрудников Казны {n_kz}, ЗУП без табеля {n_zup_only}, без флага Трудоустроен {len(no_flag)}: {no_flag}")

    # --- Сводно: сотрудник компактно + организации ЗУП ---
    grid = compose(kz, report, "СводноПоСотрудникам")
    check(len(grid) > 1, "вариант «Сводно по сотрудникам» скомпонован")
    headers, data_all, total = parse(grid)
    header = headers[0]
    h2 = headers[1] if len(headers) > 1 else None
    data = [r for r in data_all if not is_child(r)]
    print(f"   Свод: таблица {len(grid)} строк, сотрудников {len(data)}, дочерних {len(data_all) - len(data)}; заголовков {len(headers)}")
    check([c for c in header if c] == COMPACT, f"уровень сотрудника: 12 компактных колонок в нужном порядке: {[c for c in header if c]}")
    check(h2 is not None and "Организация ЗУП" in h2, "уровень организаций ЗУП присутствует в основном варианте")
    check(len(data) == n_kz + n_zup_only, f"строк сотрудников {len(data)} = Казна {n_kz} + ЗУП без табеля {n_zup_only}")
    c_fio, c_inn, c_ctl = col(header, "Сотрудник"), col(header, "ИНН"), col(header, "Контроль")
    c_dk, c_dz, c_diff = col(header, "Дни Казна (Р+М+К)"), col(header, "Дни ЗУП (работа)"), col(header, "Расхождение дней (Казна − ЗУП)")
    check(total is not None, "есть строка итогов")
    if total is not None:
        s_dk = sum(num(r[c_dk]) or 0 for r in data); s_dz = sum(num(r[c_dz]) or 0 for r in data)
        print(f"   итоги: Дни Казна {num(total[c_dk])} (Σ строк {s_dk}), Дни ЗУП {num(total[c_dz])} (Σ строк {s_dz})")
        check(num(total[c_dk]) == s_dk and num(total[c_dz]) == s_dz, "итоги дней Казны и ЗУП = суммы по сотрудникам")
    c_akt, c_tr = col(header, "В ЗУП (актуален по регистру)"), col(header, "Трудоустроен (Казна)")
    linked = [r for r in data if r[c_inn] and "нет табеля Казны" not in r[c_ctl] and (r[c_akt] == "Да" or r[c_tr] == "Да")]
    check(all((num(r[c_diff]) or 0) == (num(r[c_dk]) or 0) - (num(r[c_dz]) or 0) for r in linked), f"расхождение = Дни Казна − Дни ЗУП у связанных с ЗУП ({len(linked)} строк)")
    got_no_flag = sorted(r[c_fio] for r in data if "нет флага Трудоустроен" in r[c_ctl])
    check(got_no_flag == no_flag, f"без флага Трудоустроен: {got_no_flag}")
    n_zo = sum(1 for r in data if "нет табеля Казны" in r[c_ctl])
    check(n_zo == n_zup_only, f"строк «нет табеля Казны» = {n_zo}")
    i_ch = next((i for i, r in enumerate(data_all) if r[c_fio] == CHESHKO), None)
    check(i_ch is not None, "строка Чешко есть")
    if i_ch is not None:
        r = data_all[i_ch]
        print("   Чешко (свод):", [r[i] for i in range(len(header)) if header[i]])
        check(num(r[c_dz]) == 21 and num(r[col(header, 'Часы ЗУП (работа)')]) == 168, "Чешко: ЗУП 21 дн / 168 ч")
        check(num(r[c_diff]) == (num(r[c_dk]) or 0) - 21 and ("ЗУП > Казна на" in r[c_ctl] or num(r[c_diff]) == 0), "Чешко: расхождение с текстом «ЗУП > Казна на N дн»")
        if h2:
            ci_org = col(h2, "Организация ЗУП")
            orgs = [rr for rr in children_after(data_all, i_ch) if first(rr).startswith("ТОВ")]
            print("   организации Чешко:", [(rr[ci_org], rr[col(h2, 'Дни работа (орг)')], rr[col(h2, 'Приём в организацию')], rr[col(h2, 'Увольнение из организации')]) for rr in orgs])
            check(len(orgs) == 1 and "ІНДАСТРІАЛБУД" in orgs[0][ci_org] and num(orgs[0][col(h2, "Дни работа (орг)")]) == 21, "Чешко: только действующая ІНДАСТРІАЛБУД (ІНДЕПТ СТІЛ, уволен 22.07.2025, скрыт)")
    if h2:
        ci_org = col(h2, "Организация ЗУП")
        i_au = next((i for i, rr in enumerate(data_all) if rr[c_fio].startswith("Аулова Маргарита")), None)
        if i_au is not None:
            au = [rr[ci_org] for rr in children_after(data_all, i_au) if first(rr).startswith("ТОВ")]
            print("   организации Ауловой:", au)
            check(len(au) == 3 and not any("ІНДЕП ТРАНС" in a or "ІНДЕПТ СТІЛ" in a for a in au), "Аулова: три действующие организации, уволенные в 2024 скрыты")
        multi = run(zup, """ВЫБРАТЬ ПЕРВЫЕ 1 Вл.ДРФО КАК ДРФО, Вл.Орг КАК Орг, Вл.Дней КАК Дней ИЗ (ВЫБРАТЬ РВ.Сотрудник.Физлицо.КодПоДРФО КАК ДРФО,
                КОЛИЧЕСТВО(РАЗЛИЧНЫЕ РВ.Организация) КАК Орг,
                КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫБОР КОГДА РВ.ВидИспользованияРабочегоВремени.РабочееВремя ТОГДА НАЧАЛОПЕРИОДА(РВ.Период, ДЕНЬ) КОНЕЦ) КАК Дней
            ИЗ РегистрНакопления.РабочееВремяРаботниковОрганизаций КАК РВ ГДЕ РВ.Период МЕЖДУ &Н И &К И РВ.Активность
            СГРУППИРОВАТЬ ПО РВ.Сотрудник.Физлицо.КодПоДРФО ИМЕЮЩИЕ КОЛИЧЕСТВО(РАЗЛИЧНЫЕ РВ.Организация) > 1) КАК Вл УПОРЯДОЧИТЬ ПО Вл.ДРФО""", {"Н": bounds(zup)[0], "К": bounds(zup)[1]})
        if multi:
            drfo, n_org, n_days = zup.String(multi[0]["ДРФО"]).strip(), int(multi[0]["Орг"]), int(multi[0]["Дней"])
            i_m = next((i for i, rr in enumerate(data_all) if rr[c_inn] == drfo), None)
            check(i_m is not None, f"мульти-орг {drfo} в отчёте")
            if i_m is not None:
                mo = [rr for rr in children_after(data_all, i_m) if first(rr).startswith("ТОВ")]
                check(len(mo) >= n_org and num(data_all[i_m][c_dz]) == n_days and f"{n_org} орг. ЗУП" in data_all[i_m][c_ctl], f"мульти-орг: {n_org} организации, дни по уникальным датам = {n_days}, пометка в контроле")
    b = [r for r in data if r[c_inn] == "2627718590"]
    check(len(b) == 1, "Беспалько в своде")
    if b:
        r = b[0]
        print("   Беспалько:", r[col(header, "Приём (ЗУП)")], "|", r[col(header, "В ЗУП (актуален по регистру)")], "|", r[c_ctl])
        check(r[col(header, "Приём (ЗУП)")] == "24.04.2025" and r[col(header, "В ЗУП (актуален по регистру)")] == "Нет" and "риказ" not in r[c_ctl], "Беспалько: 24.04.2025 из даты начала, по регистру не актуален, без приказов")

    # --- Подробно ---
    grid = compose(kz, report, "Подробно")
    headers, data_d, _ = parse(grid)
    hd = headers[0]
    check(len(headers) == 1, "Подробно: плоский, одна строка заголовка")
    r = next((x for x in data_d if x[col(hd, "Сотрудник")] == CHESHKO), None)
    check(r is not None, "Подробно: строка Чешко есть")
    if r:
        n, k = bounds(kz)
        reg = run(kz, """ВЫБРАТЬ Т.Состояние КАК Состояние, КОЛИЧЕСТВО(РАЗЛИЧНЫЕ НАЧАЛОПЕРИОДА(Т.Период, ДЕНЬ)) КАК Дней,
                СУММА(Т.ЧасыФакт) КАК Факт, СУММА(Т.ЧасыОфициальные) КАК Офиц
            ИЗ РегистрНакопления.ДанныеТабелирования КАК Т
            ГДЕ Т.Период МЕЖДУ &Н И &К И Т.Активность И Т.Сотрудник.ИНН = "3406013131" И (Т.Состояние = "М" ИЛИ (Т.Состояние = "Р" И Т.ЧасыФакт > 0))
            СГРУППИРОВАТЬ ПО Т.Состояние""", {"Н": n, "К": k})
        reg = {S(x["Состояние"]): (float(x["Дней"]), float(x["Факт"]), float(x["Офиц"])) for x in reg}
        m_days, _, m_hours = reg.get("М", (0, 0, 0)); r_days, r_hours, _ = reg.get("Р", (0, 0, 0))
        print(f"   регистр Чешко: М {m_days} дн / {m_hours} ч, Р {r_days} дн / {r_hours} ч")
        check(num(r[col(hd, "Дни М (только буква М)")]) == m_days and num(r[col(hd, "Часы М (только буква М)")]) == m_hours, "Чешко: дни/часы М по регистру")
        check(num(r[col(hd, "Дни Р (по ставке)")]) == r_days and num(r[col(hd, "Часы Р (по ставке)")]) == r_hours, "Чешко: дни/часы Р по регистру")
        check(num(r[col(hd, "Часы Казна (Р+М)")]) == r_hours + m_hours and num(r[col(hd, "Дни Казна (Р+М+К)")]) == num(r[col(hd, "Дни официальные (Р+М)")]) + num(r[col(hd, "Дни К (командировка)")]), "Чешко: часы Казна = Р + М, дни Казна = официальные + К")
        check(num(r[col(hd, "Начислено (ЗУП)")]) == 29500 and r[col(hd, "Приём (ЗУП)")] == "23.07.2025" and r[col(hd, "Увольнение (ЗУП)")] == "", "Чешко: начислено 29 500, приём 23.07.2025 по регистру, не уволен")

    # --- Табели Казны ---
    grid = compose(kz, report, "ПоТабелямКазны")
    headers, data_t, _ = parse(grid)
    h1, h3 = headers[0], headers[1] if len(headers) > 1 else None
    check(h3 is not None and "Табель Казны" in h3, "вариант по табелям: вторая строка заголовка")
    if h3:
        i_ch = next(i for i, rr in enumerate(data_t) if rr[col(h1, "Сотрудник")] == CHESHKO)
        docs = [rr for rr in children_after(data_t, i_ch) if first(rr).startswith("Табель учета")]
        print("   табели Чешко:", [(rr[col(h3, 'Табель Казны')][:45], rr[col(h3, 'Часы официальные (табель)')]) for rr in docs])
        check(any("000003260" in rr[col(h3, "Табель Казны")] for rr in docs), "Чешко: табель 000003260 ссылкой")
        check(sum(num(rr[col(h3, "Часы официальные (табель)")]) or 0 for rr in docs) == num(data_t[i_ch][col(h1, "Часы Казна (Р+М)")]), "Чешко: сумма часов по табелям = часы Казна")

    # --- Контроль ---
    grid = compose(kz, report, "Контроль")
    headers, data_c, _ = parse(grid)
    emp_c = [r for r in data_c if not is_child(r)]
    check(len(emp_c) > 0 and all(r[col(headers[0], "Контроль")] for r in emp_c), f"вариант «Контроль»: {len(emp_c)} сотрудников, у всех заполнен контроль")

    # --- По дням ---
    grid = compose(kz, report, "ПоДням")
    flat = ["|".join(r) for r in grid]
    ch_days = [l for l in flat if "8М" in l and ("11.08.2026" in l or "12.08.2026" in l or "13.08.2026" in l)]
    check(len(grid) > 50 and len(ch_days) >= 3, "вариант «По дням» скомпонован, строки «8М» есть")
    check(any("Табель учета рабочего времени 000003260" in l for l in ch_days), "по дням: ссылка на табель 000003260")

    print(f"\nИтого провалов: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
