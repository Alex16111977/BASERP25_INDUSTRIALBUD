# -*- coding: utf-8 -*-
"""Приёмка этапа 2 на 30 м (§10.1): 170 позиций; мембрана 1,093 рул; плінтус 30,6667 шт; труба 110/3000 1,1667 шт; кутник 10 м;
ІБ00-000174 → септик 1/2/1/1, Σ 12 000,00; Σ норм = 1 009 864,74 ± 1,00 (И4); «Порівняти з СС» — без сирот и пустых позиций."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect, tx  # noqa: E402
from fixtures import card, kn, new_spec, row_index_by_ss, srv, ss, write  # noqa: E402


def set_card(B, h, n, code, expect_k):
    r = h.Состав.Получить(row_index_by_ss(B, h, "000000009", n))
    r.Номенклатура = card(B, code)
    hint = srv(B).ПодсказкаКоэффициента(r.Номенклатура, r.ЕдиницаСтроки)
    assert abs(hint.Коэффициент - expect_k) < 1e-9, (n, code, hint.Коэффициент)
    r.Коэффициент = hint.Коэффициент
    r.ИсточникКоэффициента = B.Перечисления.А_ИсточникиКоэффициентаПересчетаСС.Вручную


def main():
    B = connect()
    s = srv(B)
    with tx(B):
        s9 = ss(B, "000000009")
        h = new_spec(B, "Дом", "еталон 30 м", ss_ref=s9)
        p = s.ЗаполнитьИзСС(h)
        assert p.Строк == 170
        set_card(B, h, 42, "00-00006711", 75)
        # «Заповнити з СС» уже поставила ту же карточку мембраны на «Паро-гідро бар'єр» (2,8): 1,093 — норма позиции 42.
        pos42 = B.String(h.Состав.Получить(row_index_by_ss(B, h, "000000009", 42)).ПозицияСС)
        set_card(B, h, 53, "00-00009980", 3)
        set_card(B, h, 117, "00-00013067", 3)
        set_card(B, h, 118, "00-00013067", 3)
        i13 = row_index_by_ss(B, h, "000000009", 13)
        h.Состав.Получить(i13).Номенклатура = card(B, "00-00022213")
        assert s.ВзятьИзОригиналаExcel(h, i13)
        node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
        s.ПоставитьУзелВСтроку(h, row_index_by_ss(B, h, "000000009", 170), node)
        write(h)
        r = s.РазвернутьДоКарточек(h.Ссылка)
        by_card = {}
        membrana42 = 0
        for i in range(r.Листья.Количество()):
            x = r.Листья.Получить(i)
            if x.Номенклатура is not None and not x.Номенклатура.Пустая():
                by_card[x.Номенклатура.Код.strip()] = by_card.get(x.Номенклатура.Код.strip(), 0) + x.Количество
                if x.Номенклатура.Код.strip() == "00-00006711" and B.String(x.ПозицияСС) == pos42:
                    membrana42 += x.Количество
        assert round(membrana42, 4) == 1.093, membrana42
        assert round(by_card["00-00009980"], 4) == 30.6667, by_card["00-00009980"]
        assert round(by_card["00-00013067"], 4) == 1.1667, by_card["00-00013067"]
        assert round(by_card["00-00022213"], 4) == 10, by_card["00-00022213"]
        sep = [r.Листья.Получить(i) for i in range(r.Листья.Количество()) if r.Листья.Получить(i).Уровень == 1]
        assert sorted(round(x.Количество, 6) for x in sep) == [1, 1, 1, 2], "септик 1/2/1/1, не 2/4/2/2"
        pr = s.РассчитатьЦеныНормы(h)
        write(h)
        assert round(pr.СуммаСС, 2) == 1009864.74 and abs(pr.И4) <= 1.00 and pr.БезЛистьев.Количество() == 0, (pr.СуммаСС, pr.И4)
        r2 = s.РазвернутьДоКарточек(h.Ссылка, True)
        total = round(sum(r2.Листья.Получить(i).СуммаНормы for i in range(r2.Листья.Количество())), 2)
        assert abs(total - 1009864.74) <= 1.00, total
        sep_sum = round(sum(r2.Листья.Получить(i).СуммаНормы for i in range(r2.Листья.Количество())
                            if r2.Листья.Получить(i).Уровень == 1), 2)
        assert abs(sep_sum - 12000.00) <= 0.01, sep_sum
        cmp = s.СравнитьСоСтруктуройСебестоимости(h.Ссылка)
        assert cmp.Сироты.Количество() == 0 and cmp.ПозицийБезСтрок == 0 and abs(cmp.Разница) <= 1.00, \
            (cmp.Сироты.Количество(), cmp.ПозицийБезСтрок, cmp.Разница)
        st = {int(cmp.Позиции.Получить(i).НомерСтроки): cmp.Позиции.Получить(i).Статус for i in range(cmp.Позиции.Количество())}
        # Глобальные функции языка (НСтр, Формат, ТипЗнч) через COM недоступны (AttributeError, память com_date_shift_3_hours):
        # статус «Совпадает» сравнивается с обоими языковыми вариантами НСтр модуля.
        match = {"Збіг", "Совпадает"}
        assert st[13] in match and st[42] in match and st[170] in match, (st[13], st[42], st[170])
    print(f"OK test_e2b_ceny: И4 = {round(pr.И4, 2)}, септик {sep_sum}")


if __name__ == "__main__":
    main()
