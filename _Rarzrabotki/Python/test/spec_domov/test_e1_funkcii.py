# -*- coding: utf-8 -*-
"""Функции этапа 1 общего модуля через COM: хеш, единицы (Р23), разузлование (§6, Р26) на тестовом дереве в транзакции."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect, same, tx  # noqa: E402
from fixtures import add_row, card, new_spec, srv, ss, unit, write  # noqa: E402

SEPTIK = [("00-00010526", 1, 2834.672495), ("00000017651", 2, 2834.683076), ("00-00006019", 1, 661.278277),
          ("00-00012879", 1, 2834.683076)]


def t_hash_units(B):
    s = srv(B)
    h = s.ХешMD5("а/б")
    assert len(h) == 32 and h == s.ХешMD5("а/б") and h != s.ХешMD5("а/в") and h == h.lower()
    for text, key in (("пач", "пак"), ("м.п.", "м"), ("пог.м", "м"), ("балон", "шт"), ("паков", "паков"), (" т", "т"), ("шт.", "шт")):
        assert s.КлючЕдиницыСпецификации(text) == key, (text, s.КлючЕдиницыСпецификации(text))
    assert s.ЕдиницыСовпадают("пог.м", unit(B, "0101")) and s.ЕдиницыСовпадают("пач", "пак")
    assert not s.ЕдиницыСовпадают("т", "кг") and not s.ЕдиницыСовпадают("", "")
    print("OK хеш и единицы")


def t_unfold(B):
    with tx(B):
        u = unit(B, "2938")
        sep = new_spec(B, "Узел", "Септик", unit_ref=u)
        for code, qty, price in SEPTIK:
            add_row(B, sep, card_ref=card(B, code), qty=qty, price=price)
        sep_ref = write(sep)
        gen = new_spec(B, "Узел", "Генератор", unit_ref=u)
        gen.НоменклатураВузла = card(B, "00-00023184")
        gen_ref = write(gen)
        zm = new_spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        add_row(B, zm, "Узел", node=sep_ref, unit_ref=u)
        add_row(B, zm, card_ref=card(B, "00-00022147"), qty=18.75)
        zm_ref = write(zm)
        h = new_spec(B, "Дом", "розгортання", ss_ref=ss(B, "000000009"))
        add_row(B, h, card_ref=card(B, "00-00009980"), qty=92, unit_ref=unit(B, "0101"), k=3, ksrc="Вручную")
        add_row(B, h, "Узел", node=zm_ref, unit_ref=u)
        add_row(B, h, "Узел", node=gen_ref, unit_ref=u, ready=True)
        add_row(B, h, pos="не постачається", ne_post=True, comment="тест", ksrc="НеНайден")
        h_ref = write(h)
        r = srv(B).РазвернутьДоКарточек(h_ref)
        lv = [r.Листья.Получить(i) for i in range(r.Листья.Количество())]
        ek = [r.Экземпляры.Получить(i) for i in range(r.Экземпляры.Количество())]
        assert len(lv) == 7 and len(ek) == 2, (len(lv), len(ek))
        assert len({x.КлючЛиста for x in lv}) == 7 and all(len(x.КлючЛиста) == 32 for x in lv)
        plint = [x for x in lv if same(B, x.Номенклатура, card(B, "00-00009980"))][0]
        assert round(plint.Количество, 4) == 30.6667 and plint.Уровень == 0 and plint.КлючЭкземпляра == ""
        gl = [x for x in lv if same(B, x.Номенклатура, card(B, "00-00023184"))]
        assert len(gl) == 1, "«купується готовим» — лист с карточкой узла"
        sep_l = [x for x in lv if x.Уровень == 2]
        assert sorted(round(x.Количество, 6) for x in sep_l) == [1, 1, 1, 2]
        assert round(sum(x.СуммаНормы for x in sep_l), 2) == 12000.00
        codes = f"{zm_ref.Код.strip()}/{sep_ref.Код.strip()}"
        assert all(x.ПутьУзла == codes for x in sep_l), [x.ПутьУзла for x in sep_l]
        e_sep = [e for e in ek if same(B, e.Узел, sep_ref)][0]
        assert e_sep.Уровень == 2 and round(e_sep.СуммаНормы, 2) == 12000.00
        assert all(x.КлючЭкземпляра == e_sep.КлючЭкземпляра for x in sep_l)
        r2 = srv(B).РазвернутьДоКарточек(h_ref, True)
        assert r2.Листья.Количество() == 8, "строка «Не постачається» — только с флагом"
        ya = [r2.ЯвныеПозиции.Получить(i).ПозицияСС for i in range(r2.ЯвныеПозиции.Количество())]
        assert ya == ["не постачається"], ya
    print("OK разузлование")


def t_ploshchadnoy(B):
    with tx(B):
        m2 = unit(B, "0123")
        pid = new_spec(B, "Узел", "Підлога", base_qty=18, unit_ref=m2)
        add_row(B, pid, card_ref=card(B, "00-00009980"), qty=36)
        pid_ref = write(pid)
        h = new_spec(B, "Дом", "площадний", ss_ref=ss(B, "000000009"))
        add_row(B, h, "Узел", node=pid_ref, qty=18, unit_ref=m2)
        r = srv(B).РазвернутьДоКарточек(write(h))
        assert round(r.Листья.Получить(0).Количество, 6) == 36, "Р26: состав на базовую площадь 18 м²"
    print("OK площадной узел")


if __name__ == "__main__":
    B = connect()
    t_hash_units(B)
    t_unfold(B)
    t_ploshchadnoy(B)
    print("OK test_e1_funkcii")
