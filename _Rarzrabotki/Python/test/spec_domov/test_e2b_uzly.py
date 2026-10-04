# -*- coding: utf-8 -*-
"""Узлы через COM в транзакции: «Склад з документа комплектації» ІБ00-000174 → септик 1/2/1/1 (Р31), узел в строку позиции 170,
«Обʼєднати у вузол» водостік (поз. 67–74) и «Зовнішні мережі» ⊃ «Септик» (узел в узле), «Розформувати» — норма по явным позициям
не меняется ни на одном шаге; «Входимість»; имена узлов строителя (§9)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, same, tx  # noqa: E402
from fixtures import kn, new_spec, row_index_by_ss, srv, ss, unit, write  # noqa: E402


def explicit_norm(B, spec):
    r = srv(B).РазвернутьДоКарточек(spec, True)
    m = B.А_СтруктураСебестоимостиСервер
    out = {}
    for i in range(r.ЯвныеПозиции.Количество()):
        x = r.ЯвныеПозиции.Получить(i)
        k = m.КлючПозиции(x.ЭтапСС, x.ПозицияСС)
        out[k] = round(out.get(k, 0) + x.КоличествоЕдСтроки, 6)
    return out


def row_index_by_node(B, obj, node):
    for i in range(obj.Состав.Количество()):
        if obj.Состав.Получить(i).Узел is not None and same(B, obj.Состав.Получить(i).Узел, node):
            return i
    raise AssertionError("нет строки узла")


def main():
    B = connect()
    s = srv(B)
    names = [s.ИменаУзловСтроителя().Получить(i) for i in range(s.ИменаУзловСтроителя().Количество())]
    assert len(names) == 21 and "Септик" in names and "Кріплення загального призначення" in names
    with tx(B):
        s9 = ss(B, "000000009")
        kompl = unit(B, "2938")
        h = new_spec(B, "Дом", "вузли 30 м", ss_ref=s9)
        s.ЗаполнитьИзСС(h)
        node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
        lv = s.РазвернутьДоКарточек(node).Листья
        assert sorted(round(lv.Получить(i).Количество, 6) for i in range(lv.Количество())) == [1, 1, 1, 2], "не 2/4/2/2"
        s.ПоставитьУзелВСтроку(h, row_index_by_ss(B, h, "000000009", 170), node)
        write(h)
        before = explicit_norm(B, h.Ссылка)
        idx = [row_index_by_ss(B, h, "000000009", n) for n in range(67, 75)]
        vod = s.ОбъединитьВУзел(h, arr(B, idx), "Водостік 30 м", 1, kompl)
        write(h)
        assert explicit_norm(B, h.Ссылка) == before, "водостік изменил норму"
        idx = [row_index_by_node(B, h, node)] + [row_index_by_ss(B, h, "000000009", n) for n in (169, 145, 142)]
        zm = s.ОбъединитьВУзел(h, arr(B, idx), "Зовнішні мережі", 1, kompl)
        write(h)
        assert explicit_norm(B, h.Ссылка) == before, "узел в узле изменил норму"
        vh = s.Входимость(node)
        lv = {B.String(vh.Получить(i).Спецификация): vh.Получить(i).Уровень for i in range(vh.Количество())}
        assert lv == {"Зовнішні мережі": 1, "__ТЕСТ_П1 вузли 30 м": 2}, lv
        r = s.РазвернутьДоКарточек(h.Ссылка)
        sep = [r.Листья.Получить(i) for i in range(r.Листья.Количество()) if r.Листья.Получить(i).Уровень == 2]
        assert len(sep) == 4, "детали септика на уровне 2"
        assert s.РасформироватьУзел(h, row_index_by_node(B, h, zm))
        write(h)
        assert explicit_norm(B, h.Ссылка) == before, "расформирование изменило норму"
        assert s.РасформироватьУзел(h, row_index_by_node(B, h, vod))
        write(h)
        assert explicit_norm(B, h.Ссылка) == before
    print("OK test_e2b_uzly: Р31 1/2/1/1, объединение/расформирование без изменения нормы, узел в узле, входимость")


if __name__ == "__main__":
    main()
