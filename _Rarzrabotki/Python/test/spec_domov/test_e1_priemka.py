# -*- coding: utf-8 -*-
"""Приёмка этапа 1 (§10.1) — Python COM, каждый негативный кейс в своей транзакции (откат).
Кейсы: k0 запись в транзакции откатывается; k1 цикл не записывается; k2 цикл тихой записью ловится «Перевірити цикли» и при «Діє»;
k3 9 уровней узлов пишутся и разворачиваются; k4 10-й уровень не записывается; k5 10-й уровень тихой записью — исключение
разузлования; k6 Б1 — эталон на СС 003: 214/214 карточек, регистрация к обмену не тронута; k7 эталон только из «Діє»;
k8 Т9 — деталь без карточки; k9 закупочная без проверенных строк не в «Діє»; k10 обмен не тронут за всю приёмку."""
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import OUT, call, changes_rows, connect, count_rows, q, same, tx  # noqa: E402
from fixtures import (PFX, add_row, arr_list, card, expect_write_error, house, new_spec, srv, ss,  # noqa: E402
                      ss_first_row_with_card, ss_row, unit, write)

RES = {}


def k0_tx(B):
    with tx(B):
        write(new_spec(B, "Узел", "откат", unit_ref=unit(B, "2938")))
        assert len(q(B, "ВЫБРАТЬ Т.Ссылка КАК С ИЗ Справочник.А_СпецификацииДомов КАК Т ГДЕ Т.Наименование = &Н", Н=PFX + "откат")) == 1
    assert len(q(B, "ВЫБРАТЬ Т.Ссылка КАК С ИЗ Справочник.А_СпецификацииДомов КАК Т ГДЕ Т.Наименование = &Н", Н=PFX + "откат")) == 0


def k1_cycle_write(B):
    with tx(B):
        u = unit(B, "2938")
        septik = new_spec(B, "Узел", "Септик", unit_ref=u)
        s_ref = write(septik)
        zm = new_spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        add_row(B, zm, "Узел", node=s_ref, unit_ref=u)
        zm_ref = write(zm)
        add_row(B, septik, "Узел", node=zm_ref, unit_ref=u)
        expect_write_error(B, septik, "входить сам у себе", f"{PFX}Септик → {PFX}Зовнішні мережі → {PFX}Септик")


def k2_cycle_dije(B):
    with tx(B):
        u = unit(B, "2938")
        septik = new_spec(B, "Узел", "Септик", unit_ref=u)
        s_ref = write(septik)
        zm = new_spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        add_row(B, zm, "Узел", node=s_ref, unit_ref=u, checked=True)
        zm_ref = write(zm)
        add_row(B, septik, "Узел", node=zm_ref, unit_ref=u, checked=True)
        write(septik, silent=True)                       # «два пользователя»: тихая запись обходит ПередЗаписью
        errs = [str(x) for x in arr_list(srv(B).ПроверитьЦиклы())]
        assert any("входить сам у себе" in e for e in errs), errs
        h = new_spec(B, "Дом", "дім", ss_ref=ss(B, "000000009"), st="Действует")
        add_row(B, h, "Узел", node=zm_ref, unit_ref=u, checked=True)
        expect_write_error(B, h, "входить сам у себе")


def chain9(B):
    """Дом → У1 → … → У9 (узлы снизу вверх); возвращает (дом, объекты узлов по номеру)."""
    u = unit(B, "2938")
    objs, prev = {}, None
    for i in range(9, 0, -1):
        n = new_spec(B, "Узел", f"У{i}", unit_ref=u)
        if prev is not None:
            add_row(B, n, "Узел", node=prev, unit_ref=u)
        prev = write(n)
        objs[i] = n
    h = new_spec(B, "Дом", "глибина 9", ss_ref=ss(B, "000000009"))
    add_row(B, h, "Узел", node=prev, unit_ref=u)
    return write(h), objs


def k3_depth9(B):
    with tx(B):
        h, _ = chain9(B)
        r = srv(B).РазвернутьДоКарточек(h)
        levels = [r.Экземпляры.Получить(i).Уровень for i in range(r.Экземпляры.Количество())]
        assert len(levels) == 9 and max(levels) == 9, levels


def k4_depth10_write(B):
    with tx(B):
        h, objs = chain9(B)
        u10 = write(new_spec(B, "Узел", "У10", unit_ref=unit(B, "2938")))
        add_row(B, objs[9], "Узел", node=u10, unit_ref=unit(B, "2938"))
        expect_write_error(B, objs[9], "Глибина вкладеності")


def k5_depth10_unfold(B):
    with tx(B):
        h, objs = chain9(B)
        u10 = write(new_spec(B, "Узел", "У10", unit_ref=unit(B, "2938")))
        add_row(B, objs[9], "Узел", node=u10, unit_ref=unit(B, "2938"))
        write(objs[9], silent=True)
        res, e = call(srv(B).РазвернутьДоКарточек, h)
        assert res is None and "Глибина вкладеності" in e, e


def etalon_ss003(B, st):
    s3 = ss(B, "000000003")
    row = ss_first_row_with_card(B, "000000003")
    h = new_spec(B, "Дом", "еталон СС 003", ss_ref=s3, st=st)
    add_row(B, h, "Деталь", card_ref=row["Номенклатура"], qty=row["Количество"], stage=row["Этап"],
            pos=row["НоменклатураСС"], checked=True)
    return s3, write(h)


def k6_b1(B):
    cards_q = """ВЫБРАТЬ
	КОЛИЧЕСТВО(Компл.НомерСтроки) КАК Строк,
	СУММА(ВЫБОР КОГДА Компл.Номенклатура = ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка) ТОГДА 0 ИНАЧЕ 1 КОНЕЦ) КАК СКарточкой
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код"""
    with tx(B):
        s3 = ss(B, "000000003")
        before_chg, before_nom = changes_rows(B, s3), count_rows(B, "Справочник.Номенклатура.Изменения")
        _, h = etalon_ss003(B, "Действует")
        assert srv(B).СделатьЭталономСС(h)
        c = q(B, cards_q, Код="000000003")[0]
        assert (int(c["Строк"]), int(c["СКарточкой"])) == (214, 214), c
        assert changes_rows(B, s3) == before_chg, (before_chg, changes_rows(B, s3))
        assert count_rows(B, "Справочник.Номенклатура.Изменения") == before_nom
        e = q(B, "ВЫБРАТЬ СС.ЭталоннаяСпецификация КАК Э ИЗ Справочник.СтруктураСебестоимости КАК СС ГДЕ СС.Код = &Код",
              Код="000000003")[0]["Э"]
        assert same(B, e, h)
        RES["b1"] = {"cards": [214, 214], "changes": before_chg}


def k7_etalon_only_dije(B):
    with tx(B):
        _, h = etalon_ss003(B, "Черновик")
        res, e = call(srv(B).СделатьЭталономСС, h)
        assert res is None and "переведіть специфікацію в «Діє»" in e, e


def k8_t9(B):
    s9 = ss(B, "000000009")
    row = ss_row(B, "000000009", 2)
    common = dict(stage=row["Этап"], pos=row["НоменклатураСС"], qty=row["Количество"])
    with tx(B):
        h = new_spec(B, "Дом", "чернетка без картки", ss_ref=s9)
        add_row(B, h, ksrc="НеНайден", **common)
        write(h)
    with tx(B):
        h = new_spec(B, "Дом", "діє без картки", ss_ref=s9, st="Действует")
        add_row(B, h, checked=True, ksrc="НеНайден", **common)
        expect_write_error(B, h, "деталь без картки")
    with tx(B):
        h = new_spec(B, "Дом", "нп без причини", ss_ref=s9)
        add_row(B, h, ne_post=True, ksrc="НеНайден", **common)
        expect_write_error(B, h, "потрібна причина")
    with tx(B):
        h = new_spec(B, "Дом", "нп з причиною", ss_ref=s9, st="Действует")
        add_row(B, h, ne_post=True, comment="тест: не постачається", ksrc="НеНайден", **common)
        write(h)


def k9_zakupochna(B):
    """Закупочная: без проверенных строк не в «Діє»; дом факта другого типа (15 м в закупочной 30 м) — отказ (§4.5);
    проверенная строка и дом своего типа — записывается."""
    s9 = ss(B, "000000009")
    row = ss_row(B, "000000009", 42)
    for checked, home_ss, needle in ((False, "000000009", "не перевірено"), (True, "000000008", "не будинок цього типу"),
                                     (True, "000000009", None)):
        with tx(B):
            et = new_spec(B, "Дом", "еталон", ss_ref=s9, st="Действует")
            add_row(B, et, card_ref=card(B, "00-00006711"), qty=1, stage=row["Этап"], pos=row["НоменклатураСС"], checked=True)
            et_ref = write(et)
            z = new_spec(B, "Закупочная", "закупівельна", ss_ref=s9, st="Действует")
            z.Основание = et_ref
            z.ДомаФакта.Добавить().Подразделение = house(B, home_ss, 1)
            zr = z.ЗакупочныйСостав.Добавить()
            zr.Номенклатура = card(B, "00-00006711")
            zr.Количество = 1
            zr.Проверено = checked
            if needle is None:
                write(z)
            else:
                expect_write_error(B, z, needle)


EXCHANGE_Q = """ВЫБРАТЬ
	Изм.Ссылка.Код КАК Код,
	ПРЕДСТАВЛЕНИЕ(Изм.Узел) КАК Узел,
	ЕСТЬNULL(Изм.НомерСообщения, -1) КАК НомерСообщения
ИЗ
	Справочник.СтруктураСебестоимости.Изменения КАК Изм
УПОРЯДОЧИТЬ ПО
	Код"""
START = {}


def exchange_snapshot(B):
    return [{"code": r["Код"], "node": r["Узел"], "msg": int(r["НомерСообщения"])} for r in q(B, EXCHANGE_Q)]


def k10_exchange_untouched(B):
    """Регистрация СС к обмену в конце прогона = в его начале (START заполняет main). Инвариант прогона, а не снимок
    этапа 0: между днями таблицу меняют обмен и правки пользователей (память feedback_invariantnye_testy_ne_baseline).
    Расхождение — сперва проверить, не правил ли кто-то СС во время прогона (код и узел — в тексте ошибки)."""
    now = exchange_snapshot(B)
    assert now == START["exchange"], ([x for x in now if x not in START["exchange"]],
                                      [x for x in START["exchange"] if x not in now])


CASES = [k0_tx, k1_cycle_write, k2_cycle_dije, k3_depth9, k4_depth10_write, k5_depth10_unfold, k6_b1,
         k7_etalon_only_dije, k8_t9, k9_zakupochna, k10_exchange_untouched]


def main():
    B = connect()
    START["exchange"] = exchange_snapshot(B)
    fails = {}
    for fn in CASES:
        try:
            fn(B)
            RES[fn.__name__] = "OK"
            print("OK  ", fn.__name__)
        except Exception as e:                         # noqa: BLE001 — тест собирает все провалы
            fails[fn.__name__] = repr(e)[:500]
            print("FAIL", fn.__name__, repr(e)[:500])
        assert not B.TransactionActive()
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, f"e1_priemka_{datetime.datetime.now():%Y%m%d_%H%M}.json")
    json.dump({"ok": RES, "fail": fails}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("->", out, "ПРОВАЛОВ:", len(fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
