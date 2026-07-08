# -*- coding: utf-8 -*-
"""Smoke внешней обработки «Загрузка материалов в СС» на базе BuhBud.
Ожидание вычисляется из того же Excel через openpyxl, факт — из ТЧ Комплектующие
временной СтруктураСебестоимости после вызова движка ЗагрузитьМатериалыИзМассива."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
import openpyxl

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Загрузка СС\Загрузка материалов в СС.epf"
XLSX = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Загрузка СС\Виробництво\IRC 15м2 НОВИЙ ШАблон  СС Виробництво 16-06-2026_.xlsx"
TESTNAME = "__ТЕСТ Загрузка материалов СС"
CONN = 'Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"'

def num(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except Exception:
        return 0.0

def read_expected(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Проект СС"]
    def g(r, c): return ws.cell(row=r, column=c).value
    rows = []
    for r in range(2, ws.max_row + 1):
        c = g(r, 3)
        if (str(c).strip() if c is not None else "") != "Матеріал":
            continue
        gname = (str(g(r, 7)).strip() if g(r, 7) is not None else "")
        if gname == "":
            continue
        rows.append({
            "Этап": (str(g(r, 5)).strip() if g(r, 5) is not None else ""),
            "НоменклатураСС": gname,
            "ЕдиницаСС": (str(g(r, 9)).strip() if g(r, 9) is not None else ""),
            "Количество": num(g(r, 10)),
            "Сумма": num(g(r, 15)),
        })
    return rows

def get_or_create_struct(erp):
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ Ссылка ИЗ Справочник.СтруктураСебестоимости ГДЕ Наименование = &Н"
    q.УстановитьПараметр("Н", TESTNAME)
    sel = q.Execute().Выбрать()
    if sel.Следующий():
        return sel.Ссылка
    obj = erp.Справочники.СтруктураСебестоимости.СоздатьЭлемент()
    obj.Наименование = TESTNAME
    obj.Записать()
    return obj.Ссылка

def build_array(erp, rows):
    mass = erp.NewObject("Массив")
    for row in rows:
        s = erp.NewObject("Структура")
        s.Вставить("Этап", row["Этап"])
        s.Вставить("НоменклатураСС", row["НоменклатураСС"])
        s.Вставить("ЕдиницаСС", row["ЕдиницаСС"])
        s.Вставить("Количество", row["Количество"])
        s.Вставить("Сумма", row["Сумма"])
        mass.Добавить(s)
    return mass

def tc_totals(erp, ref):
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, СУММА(Сумма) КАК С, "
              "СУММА(ВЫБОР КОГДА Этап <> ЗНАЧЕНИЕ(Справочник.ЭтапыРабот.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СЭтап, "
              "СУММА(ВЫБОР КОГДА НоменклатураСС <> \"\" ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНаим "
              "ИЗ Справочник.СтруктураСебестоимости.Комплектующие ГДЕ Ссылка = &Р")
    q.УстановитьПараметр("Р", ref)
    sel = q.Execute().Выбрать()
    sel.Следующий()
    return int(sel.К or 0), float(sel.С or 0), int(sel.СЭтап or 0), int(sel.СНаим or 0)

def main():
    assert os.path.exists(EPF), f"НЕ найден .epf: {EPF} (собери обработку — Task 5)"
    exp = read_expected(XLSX)
    exp_cnt = len(exp)
    exp_sum = round(sum(r["Сумма"] for r in exp), 2)
    print(f"ОЖИДАНИЕ (Excel): строк={exp_cnt}, ΣСумма={exp_sum}")

    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)
    ref = get_or_create_struct(erp)
    mass = build_array(erp, exp)

    obr = erp.ВнешниеОбработки.Создать(EPF, False)
    res = obr.ЗагрузитьМатериалыИзМассива(ref, mass)
    print(f"ПРОТОКОЛ: Ошибка={res.Ошибка} Текст='{res.ТекстОшибки}' "
          f"Загружено={res.Загружено} СозданоЭтапов={res.СозданоЭтапов} Пропущено={res.Пропущено}")
    assert not res.Ошибка, f"движок вернул ошибку: {res.ТекстОшибки}"
    assert int(res.Загружено) == exp_cnt, f"Загружено {res.Загружено} != {exp_cnt}"
    assert int(res.Пропущено) == 0, f"Пропущено {res.Пропущено} != 0 (массив предфильтрован)"

    k, s, s_etap, s_naim = tc_totals(erp, ref)
    print(f"ФАКТ (ТЧ): строк={k}, ΣСумма={round(s,2)}, сЭтапом={s_etap}, сНаим={s_naim}")
    assert k == exp_cnt, f"строк в ТЧ {k} != {exp_cnt}"
    assert abs(s - exp_sum) < 0.05, f"ΣСумма {s} != {exp_sum}"
    assert s_naim == exp_cnt, "не все НоменклатураСС заполнены"
    assert s_etap == exp_cnt, "не все Этап заполнены (ссылка)"

    # идемпотентность: повторный прогон — те же строки, 0 новых этапов
    res2 = obr.ЗагрузитьМатериалыИзМассива(ref, mass)
    assert not res2.Ошибка, f"повторный прогон вернул ошибку: {res2.ТекстОшибки}"
    assert int(res2.Загружено) == exp_cnt, f"идемпотентность: загружено {res2.Загружено} != {exp_cnt}"
    k2, s2, _, _ = tc_totals(erp, ref)
    assert k2 == exp_cnt, f"идемпотентность: строк {k2} != {exp_cnt}"
    assert abs(s2 - exp_sum) < 0.05, f"идемпотентность: ΣСумма {s2} != {exp_sum}"
    assert int(res2.СозданоЭтапов) == 0, f"идемпотентность: создано этапов {res2.СозданоЭтапов} != 0"
    print("OK: smoke пройден (загрузка + суммы + этапы + идемпотентность)")

if __name__ == "__main__":
    main()
