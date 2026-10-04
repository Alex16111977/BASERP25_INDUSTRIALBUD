# -*- coding: utf-8 -*-
"""Этап 2, партия 2A через COM: факт и цены (t_fakt), «Склад з документа комплектації» Р31 (t_kn), «Заповнити з СС» на 30 м
(t_fill: 170 позиций, 17 зелёных, ни одной жёлтой на ОН с 2+ позициями), «Позначити зелені перевіреними», «Взяти з оригіналу
Excel» (кутник 10 м), подбор вне ОН позиции (Р35)."""
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import OUT, arr, connect, q, same, tx  # noqa: E402
from fixtures import card, kn, new_spec, row_index_by_ss, srv, ss  # noqa: E402


def e0():
    return json.load(open(sorted(glob.glob(os.path.join(OUT, "e0_snimok_*.json")))[-1], encoding="utf-8"))["control_sums"]


def t_fakt(B):
    s = srv(B)
    s9, s8 = ss(B, "000000009"), ss(B, "000000008")
    d9, d8 = s.ДомаТипа(s9), s.ДомаТипа(s8)
    assert (d9.Количество(), d8.Количество()) == (7, 6)
    f8 = s.ФактСписанийДома(d8)
    total = round(sum(f8.Получить(i).Сумма for i in range(f8.Количество())), 2)
    assert abs(total - e0()["written_off_15_uah_net"]) < 0.01, total
    ceny = s.ЦеныПоследнейЗакупки(arr(B, [card(B, "00-00010526")]), s9.ПодразделениеЗакупки)
    c = ceny.Получить(card(B, "00-00010526"))
    assert round(c.ЦенаБезНДС, 2) == 2679.16 and c.ЦенаСНДС > c.ЦенаБезНДС
    ost = s.ОстаткиСкладаПроекта(s9.ПодразделениеЗакупки)
    assert ost.Количество() > 0
    kand = s.КандидатыКарточек(s9)
    assert kand.Количество() == 170
    print(f"OK факт 15 м {total}, цена днища 2679,16, склад проекта карточек {ost.Количество()}")


def t_kn(B):
    s = srv(B)
    sost = s.СоставИзДокументаКомплектации(kn(B, "ІБ00-000174"))
    assert round(sost.КоличествоВыпускаДокумента, 3) == 0.5
    qty = sorted(round(sost.Строки.Получить(i).Количество, 6) for i in range(sost.Строки.Количество()))
    assert qty == [1, 1, 1, 2], qty
    half = s.СоставИзДокументаКомплектации(kn(B, "ІБ00-000174"), 2)
    assert sorted(round(half.Строки.Получить(i).Количество, 6) for i in range(half.Строки.Количество())) == [0.5, 0.5, 0.5, 1]
    docs = s.ДокументыКомплектацииПозиций(ss(B, "000000009"))
    assert any(int(docs.Получить(i).НомерСтроки) == 170 for i in range(docs.Количество()))
    print("OK Р31: ІБ00-000174 → 1/2/1/1 (выпуск 0,5 не делит состав)")


def t_fill(B):
    s = srv(B)
    with tx(B):
        h = new_spec(B, "Дом", "еталон 30 м", ss_ref=ss(B, "000000009"))
        p = s.ЗаполнитьИзСС(h)
        assert (p.Строк, p.Зеленых, p.ПропущеноНулевых) == (170, 17, 0), (p.Строк, p.Зеленых, p.ПропущеноНулевых)
        assert round(p.Сумма, 2) == 1009864.74 and p.Групп == 9, (p.Сумма, p.Групп)
        kand = s.КандидатыКарточек(ss(B, "000000009"))
        on_positions = {}
        for i in range(kand.Количество()):
            r = kand.Получить(i)
            if r.Количество != 0:
                on_positions[B.String(r.ОбщееНазвание)] = on_positions.get(B.String(r.ОбщееНазвание), 0) + 1
        yellow = [h.Состав.Получить(i) for i in range(h.Состав.Количество())
                  if B.String(h.Состав.Получить(i).ИсточникКарточки) == "Списувалась на будинки типу"]
        assert yellow and all(on_positions[B.String(r.ОбщееНазвание)] == 1 for r in yellow), "жёлтая на ОН с 2+ позициями"
        marked = s.ОтметитьЗеленыеПроверенными(h)
        assert 0 < marked <= 17
        i13 = row_index_by_ss(B, h, "000000009", 13)
        h.Состав.Получить(i13).Номенклатура = card(B, "00-00022213")
        assert s.ВзятьИзОригиналаExcel(h, i13)
        r13 = h.Состав.Получить(i13)
        assert round(r13.Количество, 6) == 10 and r13.Коэффициент == 1 and B.String(r13.ИсточникКоэффициента) == "Одиниці збігаються", \
            (r13.Количество, r13.Коэффициент, B.String(r13.ИсточникКоэффициента))
        print(f"OK «Заповнити з СС»: 170 / 17 зелёных / жёлтых {len(yellow)} (все на ОН с одной позицией) / отмечено {marked}; кутник 10 м")


def t_podbor(B):
    s = srv(B)
    s9 = ss(B, "000000009")
    on42 = q(B, "ВЫБРАТЬ К.ОбщееНазвание КАК ОН ИЗ Справочник.СтруктураСебестоимости.Комплектующие КАК К ГДЕ К.Ссылка = &С И К.НомерСтроки = 42", С=s9)[0]["ОН"]
    only_on = s.ПредложитьКарточки(s9, on42, "Вітробарєр", True, "")
    assert not any(same(B, only_on.Получить(i).Номенклатура, card(B, "00-00006711")) for i in range(only_on.Количество()))
    wide = s.ПредложитьКарточки(s9, on42, "Вітробарєр", False, "Мембрана супердифузійна")
    rows = [wide.Получить(i) for i in range(wide.Количество())]
    assert any(same(B, r.Номенклатура, card(B, "00-00006711")) for r in rows), "подбор вне ОН (E3 §3.5)"
    qty = [r.СписаноКоличество for r in rows]
    assert qty == sorted(qty, reverse=True), "сортировка по факту"
    print("OK подбор: мембрана находится только вне ОН позиции")


if __name__ == "__main__":
    B = connect()
    t_fakt(B)
    t_kn(B)
    t_fill(B)
    t_podbor(B)
    print("OK test_e2a_zapolnit")
