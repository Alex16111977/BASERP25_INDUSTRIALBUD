# -*- coding: utf-8 -*-
"""Смок .erf «Табели Казны и ЗУП» (kazna, август 2026), форма v3: один набор, ресурсы по уровням,
стеки «дни / часы», организации в тех же колонках, что и сотрудник. Печать PASS/FAIL, exit code 1 при провале."""
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
MAIN_TITLES = ["Сотрудник", "ИНН", "Трудоустроен (Казна)", "Актуален в ЗУП", "Приём (ЗУП)", "Увольнение (ЗУП)",
               "Расхождение дней (Казна − ЗУП)", "Контроль", "Казна: дни / часы", "ЗУП: дни / часы (по орг. — норма)",
               "Договор (ЗУП)", "Отпуск ЗУП (дн)", "Больничный ЗУП (дн)", "Без оплаты ЗУП (дн)", "Начислено (ЗУП)"]
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


def stack(s):
    """«21\\n168» → (21.0, 168.0); «4099 / 32000» → тоже."""
    parts = [p for p in (s or "").replace("/", "\n").split("\n") if p.strip()]
    return (num(parts[0]), num(parts[1])) if len(parts) == 2 else (None, None)


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


def first(row):
    return next((c for c in row if c), "")


def parse(grid):
    hdr = next(i for i, row in enumerate(grid) if "Сотрудник" in row and "Контроль" in row)
    headers = [grid[hdr]]
    i = hdr + 1
    while i < len(grid) and ("Организация ЗУП" in grid[i] or "Табель Казны" in grid[i]):
        headers.append(grid[i]); i += 1
    data, total = [], None
    for row in grid[i:]:
        if first(row).startswith("Итого"):
            total = row
            break
        if any(row):
            data.append(row)
    return headers, data, total


def col(header, title):
    assert title in header, "нет колонки %s в %s" % (title, header)
    return header.index(title)


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
    параметры = [S(схема0.Параметры.Получить(i).Имя) for i in range(схема0.Параметры.Количество())]
    check(параметры == ["Период"], f"параметры СКД только Период: {параметры}")
    варианты = [S(схема0.ВариантыНастроек.Получить(i).Имя) for i in range(схема0.ВариантыНастроек.Количество())]
    check(варианты == ["СводноПоСотрудникам", "Контроль", "ПоТабелямКазны", "Подробно", "ПоДням"], f"варианты: {варианты}")

    n_kz, n_zup_only, no_flag = expected(kz, zup)
    print(f"   ожидание: сотрудников Казны {n_kz}, ЗУП без табеля {n_zup_only}, без флага Трудоустроен {len(no_flag)}: {no_flag}")

    # --- Сводно ---
    grid = compose(kz, report, "СводноПоСотрудникам")
    check(len(grid) > 1, "вариант «Сводно по сотрудникам» скомпонован")
    headers, data_all, total = parse(grid)
    header = headers[0]
    emp = [r for r in data_all if not is_child(r)]
    print(f"   Свод: таблица {len(grid)} строк, сотрудников {len(emp)}, строк организаций {len(data_all) - len(emp)}, заголовков {len(headers)}")
    check(len(headers) == 2 and len([c for c in headers[1] if c]) == 1, "заголовок: колонки одни на оба уровня, вторая строка только подпись «Организация ЗУП»")
    check([c for c in header if c] == MAIN_TITLES, f"колонки основного варианта: {[c for c in header if c]}")
    check(len(emp) == n_kz + n_zup_only, f"строк сотрудников {len(emp)} = Казна {n_kz} + ЗУП без табеля {n_zup_only}")
    c = {t: col(header, t) for t in MAIN_TITLES}
    check(total is not None, "есть строка итогов")
    if total is not None:
        tk, tz = stack(total[c["Казна: дни / часы"]]), stack(total[c["ЗУП: дни / часы (по орг. — норма)"]])
        sk = sum((stack(r[c["Казна: дни / часы"]])[0] or 0) for r in emp); sz = sum((stack(r[c["ЗУП: дни / часы (по орг. — норма)"]])[0] or 0) for r in emp)
        print(f"   итоги: Казна {total[c['Казна: дни / часы']]!r} (Σ дней по строкам {sk}), ЗУП {total[c['ЗУП: дни / часы (по орг. — норма)']]!r} (Σ {sz})")
        check(tk[0] == sk and tz[0] == sz, "итог «дни / часы» = сумма по сотрудникам (Казна и ЗУП)")
    got_no_flag = sorted(r[c["Сотрудник"]] for r in emp if "нет флага Трудоустроен" in r[c["Контроль"]])
    check(got_no_flag == no_flag, f"без флага Трудоустроен: {got_no_flag}")
    check(sum(1 for r in emp if "нет табеля Казны" in r[c["Контроль"]]) == n_zup_only, "строк «нет табеля Казны» = ЗУП без табеля")
    linked = [r for r in emp if r[c["ИНН"]] and "нет табеля Казны" not in r[c["Контроль"]] and (r[c["Актуален в ЗУП"]] == "Да" or r[c["Трудоустроен (Казна)"]] == "Да")]
    check(all((num(r[c["Расхождение дней (Казна − ЗУП)"]]) or 0) == (stack(r[c["Казна: дни / часы"]])[0] or 0) - (stack(r[c["ЗУП: дни / часы (по орг. — норма)"]])[0] or 0) for r in linked),
          f"расхождение = дни Казна − дни ЗУП у связанных с ЗУП ({len(linked)} строк)")
    i_ch = next((i for i, r in enumerate(data_all) if r[c["Сотрудник"]] == CHESHKO), None)
    check(i_ch is not None, "строка Чешко есть")
    if i_ch is not None:
        r = data_all[i_ch]
        print("   Чешко:", [r[c[t]] for t in MAIN_TITLES])
        check(stack(r[c["ЗУП: дни / часы (по орг. — норма)"]]) == (21.0, 168.0) and "\n" in r[c["ЗУП: дни / часы (по орг. — норма)"]], "Чешко: ЗУП «21 / 168» двумя строками в одной ячейке")
        check(r[c["Приём (ЗУП)"]] == "23.07.2025" and r[c["Увольнение (ЗУП)"]] == "" and r[c["Актуален в ЗУП"]] == "Да", "Чешко: приём 23.07.2025, не уволен, актуален")
        orgs = children_after(data_all, i_ch)
        print("   организации Чешко:", [(first(o), o[c["Приём (ЗУП)"]], o[c["Увольнение (ЗУП)"]], o[c["ЗУП: дни / часы (по орг. — норма)"]], o[c["Договор (ЗУП)"]], o[c["Начислено (ЗУП)"]]) for o in orgs])
        check(len(orgs) == 1 and "ІНДАСТРІАЛБУД" in first(orgs[0]), "Чешко: одна действующая организация (уволенная до периода скрыта)")
        if orgs:
            o = orgs[0]
            check(o[c["Приём (ЗУП)"]] == "23.07.2025" and stack(o[c["ЗУП: дни / часы (по орг. — норма)"]]) == (21.0, 168.0) and o[c["Договор (ЗУП)"]] == "Трудовой" and num(o[c["Начислено (ЗУП)"]]) == 29500,
                  "Чешко/ІНДАСТРІАЛБУД: приём, норма 21/168, договор, начислено — в тех же колонках, что у сотрудника")
            check(o[c["ИНН"]] == "" and o[c["Контроль"]] == "" and o[c["Казна: дни / часы"]] == "", "строка организации: колонки сотрудника пустые")
    i_au = next((i for i, r in enumerate(data_all) if r[c["Сотрудник"]].startswith("Аулова Маргарита")), None)
    if i_au is not None:
        au = [first(o) for o in children_after(data_all, i_au)]
        print("   организации Ауловой:", au)
        check(len(au) == 3 and not any("ІНДЕП ТРАНС" in a or "ІНДЕПТ СТІЛ" in a for a in au), "Аулова: три действующие организации, уволенные в 2024 скрыты")
    b = [r for r in emp if r[c["ИНН"]] == "2627718590"]
    check(len(b) == 1, "Беспалько в своде")
    if b:
        r = b[0]
        print("   Беспалько:", r[c["Приём (ЗУП)"]], "|", r[c["Актуален в ЗУП"]], "|", r[c["Контроль"]])
        check(r[c["Приём (ЗУП)"]] == "24.04.2025" and r[c["Актуален в ЗУП"]] == "Нет" and "риказ" not in r[c["Контроль"]], "Беспалько: 24.04.2025, не актуален, без приказов")

    # --- Подробно ---
    grid = compose(kz, report, "Подробно")
    headers, data_d, total_d = parse(grid)
    hd = headers[0]
    check(len(data_d) == len(emp), f"Подробно: одна строка на сотрудника ({len(data_d)})")
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
        check(num(r[col(hd, "Дни М (только буква М)")]) == m_days and num(r[col(hd, "Часы М (только буква М)")]) == m_hours, "Чешко: дни/часы М по регистру")
        check(num(r[col(hd, "Дни Р (по ставке)")]) == r_days and num(r[col(hd, "Часы Р (по ставке)")]) == r_hours, "Чешко: дни/часы Р по регистру")
        check(num(r[col(hd, "Часы Казна (Р+М)")]) == r_hours + m_hours and num(r[col(hd, "Дни Казна (Р+М+К)")]) == num(r[col(hd, "Дни официальные (Р+М)")]) + num(r[col(hd, "Дни К (командировка)")]), "Чешко: часы Казна = Р + М, дни Казна = официальные + К")
        check(num(r[col(hd, "Начислено (ЗУП)")]) == 29500 and r[col(hd, "Приём (ЗУП, сотрудник)")] == "23.07.2025", "Чешко: начислено 29 500, приём 23.07.2025")
    if total_d is not None:
        check(num(total_d[col(hd, "Дни ЗУП (работа, уник. даты)")]) == sum(num(x[col(hd, "Дни ЗУП (работа, уник. даты)")]) or 0 for x in data_d), "Подробно: итог дней ЗУП = сумма строк")

    # --- Табели ---
    grid = compose(kz, report, "ПоТабелямКазны")
    headers, data_t, _ = parse(grid)
    h1 = h3 = headers[0]
    check(len(headers) == 2 and "Часы официальные (табель)" in h3, "вариант по табелям: колонки одни на оба уровня, табели в первой колонке под сотрудником")
    if "Часы официальные (табель)" in h3:
        i_ch = next(i for i, rr in enumerate(data_t) if rr[col(h1, "Сотрудник")] == CHESHKO)
        docs = [rr for rr in children_after(data_t, i_ch) if first(rr).startswith("Табель учета")]
        print("   табели Чешко:", [(first(rr)[:45], rr[col(h3, "Часы официальные (табель)")]) for rr in docs])
        check(any("000003260" in first(rr) for rr in docs), "Чешко: табель 000003260 ссылкой в первой колонке")
        check(sum(num(rr[col(h3, "Часы официальные (табель)")]) or 0 for rr in docs) == stack(data_t[i_ch][col(h1, "Казна: дни / часы")])[1], "Чешко: сумма часов по табелям = часы Казна")

    # --- Контроль ---
    grid = compose(kz, report, "Контроль")
    headers, data_c, _ = parse(grid)
    emp_c = [r for r in data_c if not is_child(r)]
    check(len(emp_c) > 0 and all(r[col(headers[0], "Контроль")] for r in emp_c), f"вариант «Контроль»: {len(emp_c)} сотрудников, у всех заполнен контроль")

    # --- По дням ---
    grid = compose(kz, report, "ПоДням")
    flat = ["|".join(r) for r in grid]
    ch_days = [l for l in flat if "8М" in l and ("11.08.2026" in l or "12.08.2026" in l or "13.08.2026" in l)]
    check(len(grid) > 50 and len(ch_days) >= 3 and any("Табель учета рабочего времени 000003260" in l for l in ch_days), "вариант «По дням»: строки «8М» со ссылкой на табель 000003260")

    print(f"\nИтого провалов: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
