# -*- coding: utf-8 -*-
"""Rule #-1 для запросов этапа 1: каждый файл zaprosy/*.sql исполняется на bas_industrialbud до переноса в BSL.
Данные — тестовые элементы внутри транзакции (откат). Печатает «OK, строк=N» по каждому запросу."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import arr, connect, q, same, tx, zapros  # noqa: E402

PFX = "__ТЕСТ_П1 "


def spec(B, kind, name, ss_ref=None, st="Черновик", unit_ref=None):
    o = B.Справочники.А_СпецификацииДомов.СоздатьЭлемент()
    o.Наименование = PFX + name
    o.ВидСпецификации = getattr(B.Перечисления.А_ВидыСпецификацийДомов, kind)
    o.Статус = getattr(B.Перечисления.А_СтатусыСпецификацийДомов, st)
    o.КоличествоВыпуска = 1
    if ss_ref is not None:
        o.СтруктураСебестоимости = ss_ref
    if unit_ref is not None:
        o.ЕдиницаВыпуска = unit_ref
    return o


def node_row(B, o, node, unit_ref):
    r = o.Состав.Добавить()
    r.КлючСтроки = B.NewObject("УникальныйИдентификатор")
    r.ВидСтроки = B.Перечисления.А_ВидыСтрокСпецификацииДомов.Узел
    r.Узел = node
    r.Количество = 1
    r.ЕдиницаСтроки = unit_ref
    r.Коэффициент = 1


def silent(o):
    o.ОбменДанными.Загрузка = True
    o.Записать()
    return o.Ссылка


def run(B, name, **params):
    rows = q(B, zapros(name), **params)
    print(f"OK {name}, строк={len(rows)}")
    return rows


def main():
    B = connect()
    s9 = B.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000009")
    u = B.Справочники.КлассификаторЕдиницИзмерения.НайтиПоКоду("2938")
    home = q(B, "ВЫБРАТЬ ПЕРВЫЕ 1 П.Ссылка КАК Ссылка ИЗ Справочник.Подразделения КАК П ГДЕ П.Родитель = &Р УПОРЯДОЧИТЬ ПО П.Код",
             Р=s9.ПодразделениеОбъекта)[0]["Ссылка"]
    with tx(B):
        septik = silent(spec(B, "Узел", "Септик", unit_ref=u))
        zm_o = spec(B, "Узел", "Зовнішні мережі", unit_ref=u)
        node_row(B, zm_o, septik, u)
        zm = silent(zm_o)
        h_o = spec(B, "Дом", "дім", s9, "Действует")
        node_row(B, h_o, zm, u)
        h_o.Дома.Добавить().Подразделение = home
        h = silent(h_o)
        v_o = spec(B, "Дом", "варіант", s9, "Действует")
        v_o.Дома.Добавить().Подразделение = home
        v = silent(v_o)
        z1 = silent(spec(B, "Закупочная", "закупівельна діє", s9, "Действует"))
        z2 = silent(spec(B, "Закупочная", "закупівельна чернетка", s9))

        z = B.NewObject("Запрос")
        z.Text = zapros("uroven_dereva")
        z.SetParameter("СпецификацииУровня", arr(B, [h, zm]))
        pak = z.ВыполнитьПакет()
        heads, rows = pak.Получить(0).Выгрузить(), pak.Получить(1).Выгрузить()
        assert heads.Количество() == 2 and rows.Количество() == 2, (heads.Количество(), rows.Количество())
        print("OK uroven_dereva, шапок=2, строк=2")

        par = run(B, "roditeli", Узлы=arr(B, [septik]))
        assert len(par) == 1 and same(B, par[0]["Спецификация"], zm)
        par = run(B, "roditeli", Узлы=arr(B, [zm]))
        assert len(par) == 1 and same(B, par[0]["Спецификация"], h)

        dv = run(B, "doma_v_variantah", Дома=arr(B, [home]), Ссылка=h, СсылкаСС=s9)
        assert len(dv) == 1 and same(B, dv[0]["Спецификация"], v)

        dz = run(B, "drugie_zakupochnye", Ссылка=z2, СсылкаСС=s9)
        assert len(dz) == 1 and same(B, dz[0]["Ссылка"], z1)

        sp = run(B, "specifikacii_dlya_proverki")
        names = {str(r["Наименование"]) for r in sp}
        assert {PFX + "Септик", PFX + "Зовнішні мережі", PFX + "дім"} <= names and PFX + "закупівельна діє" not in names

        sso = s9.ПолучитьОбъект()
        sso.ЭталоннаяСпецификация = h
        sso.ОбменДанными.Загрузка = True
        sso.ДополнительныеСвойства.Вставить("ОтключитьМеханизмРегистрацииОбъектов", True)
        sso.Записать()
        lst = [r for r in run(B, "spisok_formy") if same(B, r["Ссылка"], h)]
        assert len(lst) == 1 and same(B, lst[0]["ЭталонДляСС"], s9), "колонка «Еталон для СС»"
    print("OK test_e1_zaprosy")


if __name__ == "__main__":
    main()
