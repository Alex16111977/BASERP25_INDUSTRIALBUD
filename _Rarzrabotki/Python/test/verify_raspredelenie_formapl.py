# -*- coding: utf-8 -*-
"""
Rule #-1 / bas-verify для Части A (ФормаPL в РаспределениеПрочихЗатрат).

Тест-документ: РаспределениеПрочихЗатрат №00000000279 от 31.12.2025.
  Распределяет статью "Зарплата управленческого персонала" (ОЗФУ-источник),
  одно движение Расход Σ=1 374 953,46, сейчас ФормаPL=(пусто).
  Источник (ОЗФУ Приход) в декабре 2025 по этой (статья,орг,подр):
      Форма1 = 342 280,10 (10 движ.)
      Форма2 = 1 032 673,36 (9 движ.)
      ИТОГО  = 1 374 953,46  == Σ распределения.

Ожидание после Части A (перепроведение): одно (пусто)-движение распадётся на
ДВА — Форма1=342 280,10 + Форма2=1 032 673,36, Σ неизменна.

Режимы:
  python verify_raspredelenie_formapl.py          # ЭТАЛОН: снять текущее состояние + источник
  python verify_raspredelenie_formapl.py --repost # SMOKE: перепровести 279 (COM) и сверить разрез
"""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
DOC_NUM = "00000000279"
# границы декабря 2025 (серверно через ДАТАВРЕМЯ в запросе)
TOL = 0.01
EXP_F1 = 342280.10
EXP_F2 = 1032673.36
EXP_SUM = 1374953.46


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN)


def doc_ref(erp):
    q = erp.NewObject("Запрос")
    q.Текст = (
        'ВЫБРАТЬ Р.Ссылка КАК Ссылка ИЗ Документ.РаспределениеПрочихЗатрат КАК Р '
        'ГДЕ Р.Номер = "%s" И Р.Дата = ДАТАВРЕМЯ(2025,12,31,23,59,59)' % DOC_NUM
    )
    sel = q.Выполнить().Выбрать()
    if not sel.Следующий():
        raise RuntimeError("Тест-документ %s не найден" % DOC_NUM)
    return sel.Ссылка


def forma_str(erp, formaPL):
    """Безопасное имя формы (перечисление через XMLСтрока, не str())."""
    if not erp.ЗначениеЗаполнено(formaPL):
        return "(пусто)"
    x = erp.XMLСтрока(formaPL)  # ...>Форма1</...
    return "Форма1" if "Форма1" in x else ("Форма2" if "Форма2" in x else x)


def movements_by_form(erp, ref):
    """Движения ПрочиеРасходы тест-документа, сгруппированные по ФормаPL."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Рег", ref)
    q.Текст = (
        "ВЫБРАТЬ П.ФормаPL КАК ФормаPL, КОЛИЧЕСТВО(*) КАК К, "
        "СУММА(П.Сумма) КАК Сумма, СУММА(П.СуммаУпр) КАК СуммаУпр, СУММА(П.СуммаРегл) КАК СуммаРегл "
        "ИЗ РегистрНакопления.ПрочиеРасходы КАК П "
        "ГДЕ П.Регистратор = &Рег И П.Активность "
        "СГРУППИРОВАТЬ ПО П.ФормаPL"
    )
    res = {}
    sel = q.Выполнить().Выбрать()
    while sel.Следующий():
        res[forma_str(erp, sel.ФормаPL)] = (int(sel.К), float(sel.Сумма), float(sel.СуммаУпр), float(sel.СуммаРегл))
    return res


def source_split(erp, ref):
    """Источник (Приход, НЕ распределение) по (статья,орг,подр) тест-документа в дек.2025, разрез по форме."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Рег", ref)
    q.Текст = """
    ВЫБРАТЬ
        П.ФормаPL КАК ФормаPL, КОЛИЧЕСТВО(*) КАК К, СУММА(П.Сумма) КАК Сумма
    ИЗ РегистрНакопления.ПрочиеРасходы КАК П
    ГДЕ П.Период МЕЖДУ ДАТАВРЕМЯ(2025,12,1,0,0,0) И ДАТАВРЕМЯ(2025,12,31,23,59,59)
        И П.Активность
        И П.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
        И (П.СтатьяРасходов, П.Организация, П.Подразделение) В
            (ВЫБРАТЬ РАЗЛИЧНЫЕ Р.СтатьяРасходов, Р.Организация, Р.Подразделение
             ИЗ РегистрНакопления.ПрочиеРасходы КАК Р ГДЕ Р.Регистратор = &Рег)
    СГРУППИРОВАТЬ ПО П.ФормаPL
    """
    res = {}
    sel = q.Выполнить().Выбрать()
    while sel.Следующий():
        res[forma_str(erp, sel.ФормаPL)] = (int(sel.К), float(sel.Сумма))
    return res


def print_state(title, mv):
    print("\n=== %s ===" % title)
    total = 0.0
    for f in sorted(mv):
        k, s = mv[f][0], mv[f][1]
        total += s
        print("  %-8s К=%-4d Сумма=%15.2f" % (f, k, s))
    print("  ИТОГО Σ = %.2f" % total)
    return total


def snapshot(erp):
    ref = doc_ref(erp)
    print("Тест-документ:", erp.XMLСтрока(ref))
    mv = movements_by_form(erp, ref)
    tot_mv = print_state("ЭТАЛОН: движения распределения (сейчас)", mv)
    src = source_split(erp, ref)
    tot_src = print_state("ИСТОЧНИК: ОЗФУ-приход по форме (дек.2025)", src)
    print("\nПроверки эталона:")
    print("  Σ движений == Σ источника: %.2f vs %.2f  -> %s" %
          (tot_mv, tot_src, "OK" if abs(tot_mv - tot_src) <= TOL else "FAIL"))
    print("  Источник Форма1 == %.2f -> %s" % (EXP_F1, "OK" if abs(src.get("Форма1", (0, 0))[1] - EXP_F1) <= TOL else "FAIL"))
    print("  Источник Форма2 == %.2f -> %s" % (EXP_F2, "OK" if abs(src.get("Форма2", (0, 0))[1] - EXP_F2) <= TOL else "FAIL"))
    return ref


def repost_and_verify(erp):
    ref = doc_ref(erp)
    before = movements_by_form(erp, ref)
    print_state("ДО перепроведения", before)
    obj = ref.ПолучитьОбъект()
    print("\nПерепроведение через COM ...")
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    after = movements_by_form(erp, ref)
    tot = print_state("ПОСЛЕ перепроведения", after)
    f1 = after.get("Форма1", (0, 0.0))[1]
    f2 = after.get("Форма2", (0, 0.0))[1]
    pusto = after.get("(пусто)", (0, 0.0))[1]
    print("\nРЕЗУЛЬТАТ SMOKE:")
    print("  Σ == эталон %.2f: %s" % (EXP_SUM, "OK" if abs(tot - EXP_SUM) <= TOL else "FAIL"))
    print("  Форма1 == %.2f: %s" % (EXP_F1, "OK" if abs(f1 - EXP_F1) <= TOL else "FAIL"))
    print("  Форма2 == %.2f: %s" % (EXP_F2, "OK" if abs(f2 - EXP_F2) <= TOL else "FAIL"))
    print("  Остаток (пусто) == 0: %s" % ("OK" if abs(pusto) <= TOL else "FAIL (%.2f)" % pusto))


def main():
    erp = connect()
    if "--repost" in sys.argv:
        repost_and_verify(erp)
    else:
        snapshot(erp)


if __name__ == "__main__":
    main()
