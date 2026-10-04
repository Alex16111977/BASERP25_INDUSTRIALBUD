# -*- coding: utf-8 -*-
"""Единицы этапа 2 через COM: ЕдиницаПоТексту (кэш классификатора), ПодсказкаКоэффициента (Р23, E3 §4.2),
РазмерНеПротиворечит (Р25)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect  # noqa: E402
from fixtures import card, srv, unit  # noqa: E402


def main():
    B = connect()
    s = srv(B)
    for text, code in ((" т", "0306"), ("т", "0306"), ("шт.", "2009"), ("пач", "2110"), ("м.п.", "0101"), ("м", "0101"), ("балон", "2009")):
        ref = s.ЕдиницаПоТексту(text)
        assert ref is not None and ref.Код.strip() == code, (text, ref.Код if ref else None)
    assert s.ЕдиницаПоТексту("непонятно") is None
    print("OK ЕдиницаПоТексту")
    m2, m = unit(B, "0123"), unit(B, "0101")
    for code, row_unit, k in (("00-00006711", m2, 75), ("00-00009980", m, 3), ("00-00013067", m, 3),
                              ("00-00021890", m2, 12.2), ("00-00022590", m2, 5), ("00-00022213", m, 1)):
        h = s.ПодсказкаКоэффициента(card(B, code), row_unit)
        assert abs(h.Коэффициент - k) < 1e-9, (code, h.Коэффициент, h.Текст)
    assert s.ПодсказкаКоэффициента(card(B, "00-00010526"), m2).Коэффициент == 0
    print("OK ПодсказкаКоэффициента: мембрана 75, плінтус 3, труба 3, EKOBOARD 12,2, Кнауф 5")
    assert s.РазмерНеПротиворечит("Мінеральна вата (50 мм)", "EKOBOARD 50*610*1250 мм 12.2 м2 1/36")
    assert not s.РазмерНеПротиворечит("Мінеральна вата (50 мм)", "EKOBOARD 100*610*1250 мм 6,1 м2 1/36")
    assert s.РазмерНеПротиворечит("Замок", "Замок врізний 85 мм")
    print("OK РазмерНеПротиворечит")
    print("OK test_e2a_edinicy")


if __name__ == "__main__":
    main()
