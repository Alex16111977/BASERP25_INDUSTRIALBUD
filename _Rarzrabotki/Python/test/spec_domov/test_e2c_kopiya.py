# -*- coding: utf-8 -*-
"""Приёмка этапа 2 на 15 м (§10.1, Р32, Т11): 15 м копией из 30 м — каждая скопированная карточка, расходящаяся с фактом 15 м,
жёлтая (независимая сверка с фактом); поз. 44, 45 (пач/пак) и 27 (балон) — k = 1 без ручного ввода; Септик переиспользован в
поз. 168 (похожий текст), «Водостік 15 м» скопирован; норма по явным позициям не изменилась; Σ = 704 612,01, И4 = 0,00;
строки «Не постачається» не мешают «Діє»."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, q, same, tx  # noqa: E402
from fixtures import (add_row, card, enum, kn, new_spec, row_index_by_ss, srv, ss, ss_row,  # noqa: E402
                      unit, write)
from test_e2b_uzly import explicit_norm  # noqa: E402


def names(B):
    """Представления значений перечислений в языке сеанса — сравнение не зависит от языка базы."""
    return (B.String(enum(B, "А_ИсточникиКарточкиСпецификации", "КопияИзДругойСпецификации")),
            B.String(enum(B, "А_ИсточникиКарточкиСпецификации", "Вручную")),
            B.String(enum(B, "А_ИсточникиКоэффициентаПересчетаСС", "ЕдиницыСовпадают")),
            B.String(enum(B, "А_ВидыСтрокСпецификацииДомов", "Узел")))


def build_source(B, s, s9):
    src = new_spec(B, "Дом", "джерело 30 м", ss_ref=s9)
    s.ЗаполнитьИзСС(src)
    for n30, code in ((27, "00000000449"), (45, "00-00006990"), (46, "00-00021890")):
        r = src.Состав.Получить(row_index_by_ss(B, src, "000000009", n30))
        r.Номенклатура = card(B, code)
        s.УстановитьКоэффициентПоЕдиницам(r)
    node = s.УзелИзДокументаКомплектации(kn(B, "ІБ00-000174"), "Септик", 1, None, s9)
    s.ПоставитьУзелВСтроку(src, row_index_by_ss(B, src, "000000009", 170), node)
    idx = [row_index_by_ss(B, src, "000000009", n) for n in range(67, 75)]
    vod_src = s.ОбъединитьВУзел(src, arr(B, idx), "Водостік 30 м", 1, unit(B, "2938"))
    return write(src), node, vod_src


def t_copy(B):
    s = srv(B)
    YELLOW, MANUAL, MATCH, NODE = names(B)
    with tx(B):
        s9, s8 = ss(B, "000000009"), ss(B, "000000008")
        src, septik, vod_src = build_source(B, s, s9)
        dst = new_spec(B, "Дом", "еталон 15 м", ss_ref=s8)
        s.ЗаполнитьИзСС(dst)
        before = explicit_norm(B, dst)
        p = s.СкопироватьИзДругойСпецификации(dst, src)
        assert explicit_norm(B, dst) == before, "копия изменила норму по позициям"
        assert p.УзловПереиспользовано >= 1 and p.УзловСкопировано >= 1, (p.УзловПереиспользовано, p.УзловСкопировано)
        r168 = dst.Состав.Получить(row_index_by_ss(B, dst, "000000008", 168))
        assert B.String(r168.ВидСтроки) == NODE and same(B, r168.Узел, septik), "Т11: септик в поз. 168"
        vod = q(B, """ВЫБРАТЬ
	Спец.Ссылка КАК Ссылка
ИЗ
	Справочник.А_СпецификацииДомов КАК Спец
ГДЕ
	Спец.Наименование = &Имя
	И Спец.Основание = &Основание""", Имя="Водостік 15 м", Основание=vod_src)
        assert len(vod) == 1, "копия «Водостік 15 м» с основанием «Водостік 30 м» теста (боевые узлы бухгалтера не мешают)"
        for n in (27, 44, 45):
            r = dst.Состав.Получить(row_index_by_ss(B, dst, "000000008", n))
            assert not r.Номенклатура.Пустая() and r.Коэффициент == 1 and B.String(r.ИсточникКоэффициента) == MATCH, \
                (n, B.String(r.Номенклатура), r.Коэффициент, B.String(r.ИсточникКоэффициента))
        fakt = s.ФактСписанийДома(s.ДомаТипа(s8))
        by_on = {}
        for i in range(fakt.Количество()):
            f = fakt.Получить(i)
            by_on.setdefault(B.String(f.ОбщееНазвание), set()).add(f.Номенклатура.Код.strip())
        checked = 0
        for i in range(dst.Состав.Количество()):
            r = dst.Состав.Получить(i)
            src_name = B.String(r.ИсточникКарточки)
            if src_name not in (YELLOW, MANUAL):
                continue
            cards = by_on.get(B.String(r.ОбщееНазвание))
            must_be_yellow = bool(cards) and r.Номенклатура.Код.strip() not in cards
            assert (src_name == YELLOW) == must_be_yellow, (r.ПозицияСС, B.String(r.Номенклатура), src_name)
            checked += 1
        assert checked > 0
        write(dst)
        pc = s.РассчитатьЦеныНормы(dst)
        assert round(pc.СуммаСС, 2) == 704612.01 and round(pc.И4, 2) == 0.00, (pc.СуммаСС, pc.И4)
    print(f"OK копия: проверено скопированных карточек {checked}, совпало {p.Совпало}, жёлтых {p.Отличий}; Σ 704 612,01, И4 0,00")


def t_diye_ne_postach(B):
    with tx(B):
        s8 = ss(B, "000000008")
        h = new_spec(B, "Дом", "діє з не постачається", ss_ref=s8, st="Действует")
        # В СС 15 м карточек нет (0 из 190 строк) — берём первую ненулевую позицию и карточку из этого же теста.
        n1 = q(B, """ВЫБРАТЬ ПЕРВЫЕ 1
	Компл.НомерСтроки КАК Н
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код
	И Компл.Количество > 0
УПОРЯДОЧИТЬ ПО
	Компл.НомерСтроки""", Код="000000008")[0]["Н"]
        r1 = ss_row(B, "000000008", int(n1))
        r2 = ss_row(B, "000000008", 2 if int(r1["НомерСтроки"]) == 1 else 1)
        add_row(B, h, card_ref=card(B, "00000000449"), qty=r1["Количество"], stage=r1["Этап"], pos=r1["НоменклатураСС"], checked=True)
        add_row(B, h, qty=r2["Количество"] or 1, stage=r2["Этап"], pos=r2["НоменклатураСС"], ne_post=True,
                comment="не постачається — тест приймання", ksrc="НеНайден")
        write(h)
    print("OK «Не постачається» не мешает «Діє»")



if __name__ == "__main__":
    B = connect()
    t_copy(B)
    t_diye_ne_postach(B)
    print("OK test_e2c_kopiya")
