# -*- coding: utf-8 -*-
"""Rule #-1 для запросов этапа 2: каждый zaprosy/*.sql исполняется на bas_industrialbud и сверяется с независимыми числами
(свежий снимок этапа 0 — тот же день; E3 §3–§4; [M5] §4.5). Только чтение, кроме двух тестовых элементов в транзакции."""
import datetime
import glob
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, OUT, arr, connect, q, same, tx, vidy_subkonto, zapros  # noqa: E402
from fixtures import card, kn, ss  # noqa: E402

# «На сейчас» без переноса даты через COM: каждая дата из Python уезжает в 1С на −3 ч (память com_date_shift_3_hours),
# а ТекущаяДатаСеанса(), прочитанная и отданная обратно, отрезала бы движения последних часов. Дальняя дата = все движения.
FAR = datetime.datetime(2100, 1, 1)


def fresh_e0():
    files = sorted(glob.glob(os.path.join(OUT, "e0_snimok_*.json")))
    stamp = datetime.datetime.strptime(os.path.basename(files[-1])[10:23], "%Y%m%d_%H%M") if files else None
    if stamp is None or datetime.datetime.now() - stamp > datetime.timedelta(hours=2):
        subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "e0_snimok.py")], check=True)
        files = sorted(glob.glob(os.path.join(OUT, "e0_snimok_*.json")))
    return json.load(open(files[-1], encoding="utf-8"))["control_sums"]


def houses(B, code):
    rows = q(B, zapros("doma_tipa"), ПодразделениеОбъекта=ss(B, code).ПодразделениеОбъекта)
    return [r["Ссылка"] for r in rows]


def main():
    B = connect()
    e0 = fresh_e0()
    h8, h9 = houses(B, "000000008"), houses(B, "000000009")
    assert (len(h8), len(h9)) == (6, 7), (len(h8), len(h9))
    print("OK doma_tipa: 6 и 7 домов")

    for hs, key in ((h8, "written_off_15_uah_net"), (h9, "written_off_30_uah_net")):
        rows = q(B, zapros("fakt_spisaniy_domov"), Дома=arr(B, hs), ВидыСубконто=vidy_subkonto(B),
                 БезОтбораНоменклатуры=True, Номенклатура=arr(B, []))
        total = round(sum(r["Сумма"] for r in rows), 2)
        assert abs(total - e0[key]) < 0.01, (key, total, e0[key])
        print(f"OK fakt_spisaniy_domov: {total} = снимок {key}")
    one = q(B, zapros("fakt_spisaniy_domov"), Дома=arr(B, h9), ВидыСубконто=vidy_subkonto(B),
            БезОтбораНоменклатуры=False, Номенклатура=arr(B, [card(B, "00000017651")]))
    assert all(same(B, r["Номенклатура"], card(B, "00000017651")) for r in one) and one, "отбор по карточке"

    s9 = ss(B, "000000009")
    kand = q(B, zapros("kandidaty_kartochek"), СС=s9, ПодразделениеОбъекта=s9.ПодразделениеОбъекта, ВидыСубконто=vidy_subkonto(B))
    green = [r for r in kand if r["Количество"] != 0 and int(r["КарточекОН"] or 0) == 1]
    assert len(kand) == 170 and len(green) == 17, (len(kand), len(green))
    assert all("Единица" in r for r in kand)
    print("OK kandidaty_kartochek: 170 позиций, 17 зелёных")

    cards = [card(B, c) for c in ("00-00010526", "00000017651", "00-00006019", "00-00012879")]
    ceny = q(B, zapros("ceny_posledney_zakupki"), Номенклатура=arr(B, cards), ПодразделениеЗакупки=s9.ПодразделениеЗакупки,
             НаДату=FAR)
    price = {B.String(r["Номенклатура"].Код): round(r["СуммаБезНДС"] / r["Количество"], 2) for r in ceny}
    assert price == {"00-00010526": 2679.16, "00000017651": 2679.17, "00-00006019": 625.0, "00-00012879": 2679.17}, price
    print("OK ceny_posledney_zakupki: септик как в [M5] §4.5")

    gr = B.NewObject("Граница", FAR, B.ВидГраницы.Включая)
    ost = q(B, zapros("ostatki_sklada_proekta"), ПодразделениеЗакупки=s9.ПодразделениеЗакупки, НаДату=gr,
            ВидыСубконто=vidy_subkonto(B), БезОтбораНоменклатуры=True, Номенклатура=arr(B, []))
    total = round(sum(r["Сумма"] or 0 for r in ost), 2)
    assert abs(total - e0["lies_project_uah_net"]) < 0.01, (total, e0["lies_project_uah_net"])
    print(f"OK ostatki_sklada_proekta: {total}")

    units = q(B, zapros("kesh_edinic"))
    first = {}
    for r in units:
        key = B.А_СтруктураСебестоимостиСервер.КлючЕдиницы(r["Наименование"])
        first.setdefault(key, r["Код"])
    assert first["т"] == "0306" and first["шт"] == "2009" and first["м"] == "0101", (first.get("т"), first.get("шт"), first.get("м"))
    print("OK kesh_edinic: т → 0306, шт → 2009, м → 0101")

    docs = q(B, zapros("dokumenty_komplektacii"), СС=s9, ПодразделениеОбъекта=s9.ПодразделениеОбъекта)
    assert any(int(r["НомерСтроки"]) == 170 and "000174" in B.String(r["ДокументКомплектации"]) for r in docs), "ІБ00-000174"
    sost = q(B, zapros("sostav_dokumenta_komplektacii"), ДокументКомплектации=kn(B, "ІБ00-000174"))
    assert sorted(float(r["Количество"]) for r in sost) == [1, 1, 1, 2], [r["Количество"] for r in sost]
    print("OK dokumenty_komplektacii / sostav: ІБ00-000174 → 1/2/1/1")

    s8 = ss(B, "000000008")
    nul = q(B, zapros("nulevye_pozicii"), СС=s8, ПодразделениеОбъекта=s8.ПодразделениеОбъекта, ВидыСубконто=vidy_subkonto(B))
    pos = {}
    for r in nul:
        pos.setdefault(int(r["НомерСтроки"]), False)
        if r["Номенклатура"] is not None:
            pos[int(r["НомерСтроки"])] = True
    assert len(pos) == 22 and sum(pos.values()) == 18, (len(pos), sum(pos.values()))
    print("OK nulevye_pozicii: 22 нулевые, 18 с фактом")

    pod = q(B, zapros("kartochki_dlya_podbora"), ТолькоОН=False, ОбщееНазвание=B.Справочники.ОбщиеНазванияНоменклатуры.ПустаяСсылка(),
            БезПоиска=False, Шаблон="%Мембрана супердифузійна%")
    assert any(B.String(r["Код"]) == "00-00006711" for r in pod), "мембрана вне ОН позиции"
    print(f"OK kartochki_dlya_podbora: {len(pod)} карточек")

    with tx(B):
        o = B.Справочники.А_СпецификацииДомов.СоздатьЭлемент()
        o.Наименование = "__ТЕСТ_П1 СС 009"
        o.ВидСпецификации = B.Перечисления.А_ВидыСпецификацийДомов.Дом
        o.Статус = B.Перечисления.А_СтатусыСпецификацийДомов.Черновик
        o.СтруктураСебестоимости = s9
        o.КоличествоВыпуска = 1
        o.ОбменДанными.Загрузка = True
        o.Записать()
        sp = q(B, zapros("specifikacii_ss"), СсылкаСС=s9)
        assert any(same(B, r["Ссылка"], o.Ссылка) for r in sp)
    print("OK specifikacii_ss")
    print("OK test_e2_zaprosy")


if __name__ == "__main__":
    main()
