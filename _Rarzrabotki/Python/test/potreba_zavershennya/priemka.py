# -*- coding: utf-8 -*-
"""Приёмка отчёта «Прогноз виконання СС за потребою» (BaseERP) — только инварианты, данные не меняются
(сценарии с изменением документа — в транзакции с откатом).

    python priemka.py --erf <путь .erf>      # внешний носитель
    python priemka.py --builtin              # встроенный отчёт конфигурации
"""
import argparse
import datetime as dt
import json
import sys
from collections import defaultdict
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
PROJ = json.loads(Path(r"C:\Configuration_downloads\BASERP25\.v8-project.json").read_text(encoding="utf-8-sig"))
SHIFT = dt.timedelta(hours=3)
D_REPORT = dt.datetime(2026, 10, 8, 12) + SHIFT
PROJECTS = ("МД IRC 2026", "МД ООН Админ")
FAILS, OKS = [], []


def check(cond, msg):
    (OKS if cond else FAILS).append(msg)
    print(("  OK   " if cond else "  FAIL ") + msg)


def connect():
    db = next(d for d in PROJ["databases"] if d["id"] == "erp")
    return win32com.client.Dispatch("V83.COMConnector").Connect(
        f'Srvr="localhost";Ref="{db["ref"]}";Usr="{db["user"]}";Pwd="{db["password"]}"')


erp = connect()
S = erp.String


def rows_of(tz, cols):
    out = []
    for i in range(tz.Количество()):
        r = tz.Получить(i)
        out.append({c: getattr(r, c) for c in cols})
    return out


def norm(s):
    return erp.Документы.А_ПотребаЗавершенняПроекту.НормализоватьТекст(s)


def report_object(args):
    if args.builtin:
        return erp.Отчеты.А_ПрогнозВиконанняССЗаПотребою.Создать()
    return erp.ВнешниеОтчеты.Создать(args.erf, False)


REP_COLS = ["Подразделение", "ДокументПотребы", "ЕтапНазва", "НазваЗагальна", "ОбщееНазвание", "ОзнакаСС", "ВидРядка", "Аналитика",
            "Документ", "Дома", "ПланКол", "ПланГрн", "КупленоКол", "КупленоГрн", "ПослеГрн", "ЗалишокСС", "ПотребаКол", "ПотребаГрн",
            "ЗалишокПотребы", "НеДокупаем"]


def report_rows(rep, date):
    tz = rep.ПолучитьДанные(date)
    rows = rows_of(tz, REP_COLS)
    for r in rows:
        r["Подр"] = S(r["Подразделение"]).strip()
        r["Назва"] = S(r["НазваЗагальна"])
        r["Вид"] = S(r["ВидРядка"])
        for k in REP_COLS[10:]:
            r[k] = float(r[k])
    return rows


def planfact(date_end):
    per = erp.NewObject("СтандартныйПериод", dt.datetime(2000, 1, 1, 12), date_end)
    tz = erp.Отчеты.А_ПланФактныйПроизводствоПолный.Создать().ПолучитьДанные(per)
    out = defaultdict(lambda: defaultdict(float))
    for r in rows_of(tz, ["Подразделение", "Блок", "ПланГрн", "ФактГрн", "ПланНаФактГрн", "ПрогнозГрн"]):
        if S(r["Блок"]) != "2. Матеріали":
            continue
        p = S(r["Подразделение"]).strip()
        for k in ("ПланГрн", "ФактГрн", "ПланНаФактГрн", "ПрогнозГрн"):
            out[p][k] += float(r[k])
    return out


def docs():
    q = erp.NewObject("Запрос")
    q.Текст = """ВЫБРАТЬ ДокПотр.Ссылка КАК Ссылка, ДокПотр.Подразделение КАК Подразделение, ДокПотр.СуммаДокумента КАК Сумма,
        ДокПотр.Дата КАК Дата ИЗ Документ.А_ПотребаЗавершенняПроекту КАК ДокПотр ГДЕ ДокПотр.Проведен И НЕ ДокПотр.ПометкаУдаления"""
    return {S(r["Подразделение"]).strip(): r for r in rows_of(q.Выполнить().Выгрузить(), ["Ссылка", "Подразделение", "Сумма", "Дата"])}


def by_project(rows):
    out = defaultdict(lambda: defaultdict(float))
    for r in rows:
        for k in REP_COLS[10:]:
            out[r["Подр"]][k] += r[k]
    return out


def by_group(rows):
    g = defaultdict(lambda: defaultdict(float))
    for r in rows:
        key = (r["Подр"], r["Назва"])
        for k in REP_COLS[10:]:
            g[key][k] += r[k]
        g[key]["_ознака"] = S(r["ОзнакаСС"])
    return g


def section_basic(rep):
    print("\n[1] Инварианты на дату отчёта 08.10.2026 (дата списка = 08.10.2026 11:36:56)")
    rows = report_rows(rep, D_REPORT)
    pf = planfact(D_REPORT)
    dd = docs()
    pr = by_project(rows)
    check(set(pr) == set(PROJECTS), f"проекты в отчёте = {sorted(pr)}")
    for p in PROJECTS:
        a, b = pr[p], pf[p]
        check(abs(a["ПланГрн"] - b["ПланГрн"]) < 0.01, f"{p}: Σ План {a['ПланГрн']:.2f} = План-факт {b['ПланГрн']:.2f}")
        check(abs(a["КупленоГрн"] - b["ФактГрн"]) < 0.01, f"{p}: Σ Куплено {a['КупленоГрн']:.2f} = факт План-факт {b['ФактГрн']:.2f}")
        check(abs(a["ЗалишокСС"] - b["ПланНаФактГрн"]) < 0.01, f"{p}: Σ Залишок за СС {a['ЗалишокСС']:.2f} = ПланНаФакт {b['ПланНаФактГрн']:.2f}")
        check(abs(a["КупленоГрн"] + a["ЗалишокСС"] - b["ПрогнозГрн"]) < 0.01,
              f"{p}: Прогноз за СС {a['КупленоГрн'] + a['ЗалишокСС']:.2f} = «План на факт» {b['ПрогнозГрн']:.2f}")
        check(abs(a["ПотребаГрн"] - float(dd[p]["Сумма"])) < 0.01, f"{p}: Σ Потреба {a['ПотребаГрн']:.2f} = сумма документа {float(dd[p]['Сумма']):.2f}")
        check(abs(a["ПослеГрн"]) < 0.005, f"{p}: Куплено після списку = 0 (дата списка = дата отчёта)")
        check(abs(a["ЗалишокПотребы"] - a["ПотребаГрн"]) < 0.01, f"{p}: Залишок потреби = Потреба ({a['ЗалишокПотребы']:.2f})")
        econ = a["ПланГрн"] - a["КупленоГрн"] - a["ЗалишокПотребы"]
        print(f"      {p}: План {a['ПланГрн']:,.2f} | Куплено {a['КупленоГрн']:,.2f} | Потреба {a['ПотребаГрн']:,.2f} | "
              f"Прогноз {a['КупленоГрн'] + a['ЗалишокПотребы']:,.2f} | Економія {econ:,.2f} ({econ / a['ПланГрн'] * 100:.1f} %) | "
              f"Економія за СС {a['ПланГрн'] - a['КупленоГрн'] - a['ЗалишокСС']:,.2f} | Не докуповуємо {a['НеДокупаем']:,.2f}".replace(",", " "))
    g = by_group(rows)
    multi = [k for k, v in g.items() if v["ПотребаГрн"] > 0 and abs(v["ЗалишокПотребы"] - v["ПотребаГрн"]) > 0.01]
    check(not multi, f"на дату списка Залишок потреби = Потреба в каждой группе (нарушений {len(multi)})")
    bad_nd = [k for k, v in g.items() if v["ПотребаГрн"] == 0 and v["НеДокупаем"] > 0 and abs(v["НеДокупаем"] - v["ЗалишокСС"]) > 0.01]
    check(not bad_nd, f"группы без потреби: Не докуповуємо = Залишок за СС (нарушений {len(bad_nd)})")
    neg = [k for k, v in g.items() if v["НеДокупаем"] < -0.005 or v["НеДокупаем"] - v["ЗалишокСС"] > 0.01]
    check(not neg, f"0 <= Не докуповуємо <= Залишок за СС в каждой группе (нарушений {len(neg)})")
    for name in ("Вікно металопластикове", "Двері металопластикові"):
        v = g.get(("МД ООН Админ", name))
        check(v is not None and v["НеДокупаем"] == 0, f"Админ «{name}»: покрыто складеною строкою, Не докуповуємо = 0")
    one_doc = defaultdict(set)
    for r in rows:
        one_doc[r["Подр"]].add(S(r["ДокументПотребы"]))
    check(all(len(v) == 1 for v in one_doc.values()), "ДокументПотребы заполнен одинаково во всех строках проекта")
    irc = pr["МД IRC 2026"]
    print(f"      ориентир письма IRC: Прогноз 9 808 849,48 / Економія 1 487 875,76 → отчёт "
          f"{irc['КупленоГрн'] + irc['ЗалишокПотребы']:,.2f} / {irc['ПланГрн'] - irc['КупленоГрн'] - irc['ЗалишокПотребы']:,.2f}".replace(",", " "))
    return rows


def last_purchase_before(project):
    """Момент списка для сценария списания: начало дня перед последней закупкой проекта по ОН из списка."""
    q = erp.NewObject("Запрос")
    q.Текст = """ВЫБРАТЬ ПЕРВЫЕ 1 Зак.Ссылка.Дата КАК Дата
    ИЗ Документ.ПриобретениеТоваровУслуг.Товары КАК Зак
    ГДЕ Зак.Ссылка.Проведен И Зак.Ссылка.Подразделение В ИЕРАРХИИ (&Подр)
      И Зак.Номенклатура.А_ОбщееНазваниеНоменклатуры В (ВЫБРАТЬ СтрПотр.ОбщееНазвание ИЗ Документ.А_ПотребаЗавершенняПроекту.Потреба КАК СтрПотр
                                                         ГДЕ СтрПотр.Ссылка.Подразделение = &Подр)
    УПОРЯДОЧИТЬ ПО Зак.Ссылка.Дата УБЫВ"""
    q.УстановитьПараметр("Подр", docs()[project]["Подразделение"])
    t = q.Выполнить().Выгрузить()
    return None if t.Количество() == 0 else t.Получить(0).Дата


def doc_date(ref):
    """Дата документа компонентами из запроса (без COM-сдвига дат)."""
    q = erp.NewObject("Запрос")
    q.Текст = ("ВЫБРАТЬ ГОД(Д.Дата) КАК Г, МЕСЯЦ(Д.Дата) КАК М, ДЕНЬ(Д.Дата) КАК Дн, ЧАС(Д.Дата) КАК Ч, МИНУТА(Д.Дата) КАК Мин, "
               "СЕКУНДА(Д.Дата) КАК С ИЗ Документ.А_ПотребаЗавершенняПроекту КАК Д ГДЕ Д.Ссылка = &Ссылка")
    q.УстановитьПараметр("Ссылка", ref)
    r = q.Выполнить().Выгрузить().Получить(0)
    return dt.datetime(int(r.Г), int(r.М), int(r.Дн), int(r.Ч), int(r.Мин), int(r.С))


def section_consumption(rep):
    print("\n[2] Списание потреби закупками после даты списка (транзакция с откатом)")
    p = "МД IRC 2026"
    last = last_purchase_before(p)
    if last is None:
        print("  SKIP нет закупок по ОН списка")
        return
    last_local = dt.datetime(last.year, last.month, last.day) - dt.timedelta(days=1)
    moment = dt.datetime(last_local.year, last_local.month, last_local.day, 12)
    print(f"      последняя закупка по ОН списка {S(last)} → момент списка {moment:%d.%m.%Y %H:%M}")
    before = docs()[p]
    orig_sum, orig_date = float(before["Сумма"]), doc_date(before["Ссылка"])
    erp.НачатьТранзакцию()
    try:
        obj = before["Ссылка"].ПолучитьОбъект()
        obj.Дата = moment + SHIFT
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        rows = report_rows(rep, D_REPORT)
        g = {k: v for k, v in by_group(rows).items() if k[0] == p}
        after = {k: v for k, v in g.items() if v["ПослеГрн"] > 0 and v["ПотребаГрн"] > 0}
        check(len(after) > 0, f"есть группы потреби с закупками после списка: {len(after)}")
        out_of = [k for k, v in g.items() if v["ЗалишокПотребы"] < -0.005 or v["ЗалишокПотребы"] - v["ПотребаГрн"] > 0.01]
        check(not out_of, f"0 <= Залишок потреби <= Потреба в каждой группе (нарушений {len(out_of)})")
        untouched = [k for k, v in g.items() if v["ПотребаГрн"] > 0 and v["ПослеГрн"] == 0 and abs(v["ЗалишокПотребы"] - v["ПотребаГрн"]) > 0.01]
        check(not untouched, f"группы без закупок после списка: Залишок = Потреба (нарушений {len(untouched)})")
        recompute = expected_remainders(p, moment)
        diff = []
        for k, v in after.items():
            exp = recompute.get(k[1])
            if exp is None or abs(exp[0] - v["ЗалишокПотребы"]) > 0.02:
                diff.append(f"{k[1]}: отчёт {v['ЗалишокПотребы']:.2f}, Python {exp}")
        check(not diff, f"Залишок потреби = независимый пересчёт Python в {len(after)} группах (расхождений {len(diff)})")
        for d_ in diff[:10]:
            print("        ", d_)
        for k, v in sorted(after.items(), key=lambda kv: -kv[1]["ПослеГрн"])[:8]:
            print(f"        {k[1]}: потреба {v['ПотребаГрн']:,.2f}, куплено після {v['ПослеГрн']:,.2f} → залишок {v['ЗалишокПотребы']:,.2f}"
                  f" [{recompute.get(k[1], ('?', '?'))[1]}]".replace(",", " "))
        # итог не зависит от ОН на дату списка; после неё текстовые строки не списываются
        tot_with = sum(v["ЗалишокПотребы"] for v in g.values())
        for st in [obj.Потреба.Получить(i) for i in range(obj.Потреба.Количество())]:
            st.ОбщееНазвание = erp.Справочники.А_ОбщиеНазванияНоменклатуры.ПустаяСсылка()
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        rows2 = report_rows(rep, D_REPORT)
        g2 = {k: v for k, v in by_group(rows2).items() if k[0] == p}
        tot_without = sum(v["ЗалишокПотребы"] for v in g2.values())
        tot_potr = sum(v["ПотребаГрн"] for v in g2.values())
        check(abs(tot_without - tot_potr) < 0.01 and tot_without >= tot_with - 0.01,
              f"без ОН потреба не списывается: Σ залишок {tot_without:,.2f} = Σ потреба {tot_potr:,.2f} ≥ с ОН {tot_with:,.2f}".replace(",", " "))
        obj.Дата = orig_date + SHIFT
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        rows3 = report_rows(rep, D_REPORT)
        a3 = by_project(rows3)[p]
        a1 = by_project(report_rows_cache[0])[p]
        check(abs((a3["КупленоГрн"] + a3["ЗалишокПотребы"]) - (a1["КупленоГрн"] + a1["ЗалишокПотребы"])) < 0.01,
              "на дату списка итог Прогноза не зависит от сопоставления ОН")
    finally:
        erp.ОтменитьТранзакцию()
    after_rb = docs()[p]
    check(abs(float(after_rb["Сумма"]) - orig_sum) < 0.01 and doc_date(after_rb["Ссылка"]) == orig_date,
          f"после отката документ IRC не изменился (сумма {orig_sum:,.2f}, дата {orig_date:%d.%m.%Y %H:%M:%S})".replace(",", " "))


def expected_remainders(project, moment):
    """Независимый пересчёт залишку потреби (правила промта 3.4 шаг 4) по исходным данным."""
    per = erp.NewObject("СтандартныйПериод", dt.datetime(2000, 1, 1, 12), D_REPORT)
    tz = erp.Отчеты.А_ПланФактныйПроизводствоПолный.Создать().ПолучитьДанные(per)
    pf = [r for r in rows_of(tz, ["Подразделение", "Блок", "ОбщееНазвание", "Документ", "ФактКол", "ФактГрн", "ПланКол", "ПланГрн", "Единица"])
          if S(r["Блок"]) == "2. Матеріали" and S(r["Подразделение"]).strip() == project]
    q = erp.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Зак.Ссылка КАК Ссылка, Зак.Дата КАК Дата ИЗ Документ.ПриобретениеТоваровУслуг КАК Зак ГДЕ Зак.Дата > &М"
    q.УстановитьПараметр("М", moment + SHIFT)
    after_docs = {S(r["Ссылка"]) for r in rows_of(q.Выполнить().Выгрузить(), ["Ссылка", "Дата"])}
    grp = defaultdict(lambda: {"после_кол": 0.0, "после_грн": 0.0, "ед": set(), "план_кол": 0.0, "план_грн": 0.0, "ед_плана": None, "макс": -1})
    for r in pf:
        name = S(r["ОбщееНазвание"]).strip() if S(r["ОбщееНазвание"]) else None
        if not name:
            continue
        g = grp[name]
        doc = r["Документ"]
        g["план_кол"] += float(r["ПланКол"])
        g["план_грн"] += float(r["ПланГрн"])
        if float(r["ПланГрн"]) > g["макс"] and float(r["ПланКол"]) + float(r["ПланГрн"]) != 0:
            g["макс"] = float(r["ПланГрн"])
            g["ед_плана"] = norm(S(r["Единица"]))
        if S(doc) in after_docs:
            g["после_кол"] += float(r["ФактКол"])
            g["после_грн"] += float(r["ФактГрн"])
            g["ед"].add(norm(S(r["Единица"])))
    q = erp.NewObject("Запрос")
    q.Текст = """ВЫБРАТЬ СтрПотр.ОбщееНазвание КАК ОН, СтрПотр.Количество КАК Кол, СтрПотр.Единица КАК Ед, СтрПотр.КоличествоСС КАК КолСС,
        СтрПотр.ЕдиницаСС КАК ЕдСС, СтрПотр.Сумма КАК Сумма, СтрПотр.Предупреждение КАК Пред, СтрПотр.НазваниеСС КАК НазваСС
        ИЗ Документ.А_ПотребаЗавершенняПроекту.Потреба КАК СтрПотр ГДЕ СтрПотр.Ссылка.Подразделение = &Подр И СтрПотр.Ссылка.Проведен
        УПОРЯДОЧИТЬ ПО СтрПотр.НомерСтроки"""
    q.УстановитьПараметр("Подр", docs()[project]["Подразделение"])
    potr = defaultdict(list)
    for r in rows_of(q.Выполнить().Выгрузить(), ["ОН", "Кол", "Ед", "КолСС", "ЕдСС", "Сумма", "Пред", "НазваСС"]):
        if S(r["ОН"]) and "/" not in S(r["НазваСС"]):
            potr[S(r["ОН"]).strip()].append(r)
    res = {}
    for name, lines in potr.items():
        g = grp[name]
        p_grn = sum(float(x["Сумма"]) for x in lines)
        if g["после_грн"] == 0:
            res[name] = (round(p_grn, 2), "без закупок")
            continue
        branch, rem = "гроші", max(0.0, p_grn - g["после_грн"])
        ok = g["после_кол"] > 0 and len(g["ед"]) == 1 and not any(("Одиниця" in S(x["Пред"]) or "кількість" in S(x["Пред"])) for x in lines)
        if ok:
            ed_after = next(iter(g["ед"]))
            ed_ss, ed_f = norm(S(lines[-1]["ЕдСС"])), norm(S(lines[-1]["Ед"]))
            k_ss, k_f = sum(float(x["КолСС"]) for x in lines), sum(float(x["Кол"]) for x in lines)
            if ed_after == ed_ss and k_ss > 0:
                kol, fld = k_ss, "КолСС"
            elif ed_after == ed_f and k_f > 0:
                kol, fld = k_f, "Кол"
            else:
                kol = None
            if kol:
                price = p_grn / kol
                homo = all(float(x[fld]) > 0 and 0.95 * price <= float(x["Сумма"]) / float(x[fld]) <= 1.05 * price for x in lines)
                fact_ok = 0.67 * price <= g["после_грн"] / g["после_кол"] <= 1.5 * price
                plan_ok = True
                if fld == "КолСС" and g["план_кол"] > 0 and g["ед_плана"] == ed_ss:
                    pp = g["план_грн"] / g["план_кол"]
                    plan_ok = 0.67 * pp <= price <= 1.5 * pp
                if homo and fact_ok and plan_ok:
                    branch, rem = "кількість", max(0.0, kol - g["после_кол"]) * price
        res[name] = (round(rem, 2), branch)
    return res


def section_compose(rep):
    print("\n[3] Компоновка всех вариантов (СформироватьОтчет)")
    schema = rep.СхемаКомпоновкиДанных
    bad_types = []
    for i in range(schema.НаборыДанных.Получить(0).Поля.Количество()):
        f = schema.НаборыДанных.Получить(0).Поля.Получить(i)
        if f.ТипЗначения.Типы().Количество() == 0:
            bad_types.append(S(f.ПутьКДанным))
    check(not bad_types, f"у всех полей набора тип задан (пустых {bad_types})")
    for i in range(schema.ВариантыНастроек.Количество()):
        var = schema.ВариантыНастроек.Получить(i)
        rep.КомпоновщикНастроек.Инициализировать(erp.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema))
        rep.КомпоновщикНастроек.ЗагрузитьНастройки(var.Настройки)
        rep.КомпоновщикНастроек.Настройки.ПараметрыДанных.УстановитьЗначениеПараметра("ДатаОтчета", D_REPORT)
        td = erp.NewObject("ТабличныйДокумент")
        rep.СформироватьОтчет(td, erp.NewObject("ДанныеРасшифровкиКомпоновкиДанных"))
        check(td.ВысотаТаблицы > 5, f"вариант «{S(var.Представление)}»: строк {td.ВысотаТаблицы}, колонок {td.ШиринаТаблицы}")
    # пустая дата (так форма открывается по умолчанию)
    rep.КомпоновщикНастроек.ЗагрузитьНастройки(schema.НастройкиПоУмолчанию)
    td = erp.NewObject("ТабличныйДокумент")
    rep.СформироватьОтчет(td, erp.NewObject("ДанныеРасшифровкиКомпоновкиДанных"))
    check(td.ВысотаТаблицы > 5, f"пустая дата отчёта (текущая): строк {td.ВысотаТаблицы}")
    # дата отчёта как в форме БСП — СтандартнаяДатаНачала
    rep.КомпоновщикНастроек.Настройки.ПараметрыДанных.УстановитьЗначениеПараметра(
        "ДатаОтчета", erp.NewObject("СтандартнаяДатаНачала", D_REPORT))
    td = erp.NewObject("ТабличныйДокумент")
    rep.СформироватьОтчет(td, erp.NewObject("ДанныеРасшифровкиКомпоновкиДанных"))
    check(td.ВысотаТаблицы > 5, f"дата как СтандартнаяДатаНачала: строк {td.ВысотаТаблицы}")


report_rows_cache = []
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--erf")
    ap.add_argument("--builtin", action="store_true")
    ap.add_argument("--date", default="2026-10-08", help="дата отчёта ГГГГ-ММ-ДД (по умолчанию дата списка директора)")
    args = ap.parse_args()
    D_REPORT = dt.datetime.strptime(args.date, "%Y-%m-%d").replace(hour=12) + SHIFT
    rep = report_object(args)
    report_rows_cache.append(section_basic(rep))
    section_consumption(rep)
    section_compose(rep)
    print(f"\nИТОГ: OK {len(OKS)}, FAIL {len(FAILS)}")
    sys.exit(1 if FAILS else 0)
