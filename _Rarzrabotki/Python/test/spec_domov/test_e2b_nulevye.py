# -*- coding: utf-8 -*-
"""Р22 на 15 м: 22 нулевые позиции, у 18 есть факт; после строк-алиасов 131, 52, 109, 70 ([Z3] §2.6) закрыто алиасами 4
(169 → 131, 173 → 52, 183 → 109, 187 → 70); печать для экономиста содержит 18 строк."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect, q, tx, vidy_subkonto, zapros  # noqa: E402
from fixtures import new_spec, row_index_by_ss, srv, ss, write  # noqa: E402


def rows(t):
    return [t.Получить(i) for i in range(t.Количество())]


def main():
    B = connect()
    s = srv(B)
    s8 = ss(B, "000000008")
    fact = q(B, zapros("nulevye_pozicii"), СС=s8, ПодразделениеОбъекта=s8.ПодразделениеОбъекта, ВидыСубконто=vidy_subkonto(B))
    first_card = {}
    for r in fact:
        if r["Номенклатура"] is not None:
            first_card.setdefault(int(r["НомерСтроки"]), r["Номенклатура"])
    with tx(B):
        h = new_spec(B, "Дом", "еталон 15 м", ss_ref=s8)
        s.ЗаполнитьИзСС(h)
        write(h)
        t = rows(s.НулевыеПозицииСФактом(h.Ссылка))
        assert len(t) == 22 and sum(1 for r in t if r.ЕстьФакт) == 18, (len(t), sum(1 for r in t if r.ЕстьФакт))
        closed_before = {int(r.НомерСтроки) for r in t if r.ЗакрытаАлиасом}
        for zero_n, alias_n in ((169, 131), (173, 52), (183, 109), (187, 70)):
            h.Состав.Получить(row_index_by_ss(B, h, "000000008", alias_n)).Номенклатура = first_card[zero_n]
        write(h)
        t = rows(s.НулевыеПозицииСФактом(h.Ссылка))
        closed = {int(r.НомерСтроки) for r in t if r.ЗакрытаАлиасом}
        # 183 может быть закрыта уже после «Заповнити з СС» (карточка факта зелёная на 109 «Канал d100/1»).
        assert closed == {169, 173, 183, 187} and closed_before <= closed, (sorted(closed_before), sorted(closed))
        doc = s.НулевыеПозицииВТабличныйДокумент(h.Ссылка)
        assert doc.ВысотаТаблицы == 3 + 18, doc.ВысотаТаблицы
    print("OK test_e2b_nulevye: 22 / 18 с фактом / 4 закрыты алиасами")


if __name__ == "__main__":
    main()
