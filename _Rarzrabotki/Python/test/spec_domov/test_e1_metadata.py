# -*- coding: utf-8 -*-
"""Приёмка метаданных этапа 1 через COM (после загрузки 1A): состав, uk-синонимы, обмен, права, навигация, РК не проводится.
До загрузки падает: «перечисление А_ВидыСпецификацийДомов не найдено в базе». Списки ниже — единый источник для test_e1_files.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect  # noqa: E402

UK_ENUM = {
    "А_ВидыСпецификацийДомов": {"Дом": "Будинок", "Узел": "Вузол", "Закупочная": "Закупівельна"},
    "А_СтатусыСпецификацийДомов": {"Черновик": "Чернетка", "Действует": "Діє", "Закрыта": "Закрита"},
    "А_ВидыСтрокСпецификацииДомов": {"Деталь": "Деталь", "Узел": "Вузол",
                                     "ГруппаПоОбщемуНазванию": "Група за загальною назвою"},
    "А_ИсточникиКарточкиСпецификации": {"ЕдинственнаяКарточкаОН": "Єдина картка ОН",
                                        "СписаннаяНаДомаТипа": "Списувалась на будинки типу",
                                        "ИзДокументаКомплектации": "З документа комплектації",
                                        "КопияИзДругойСпецификации": "Скопійовано з іншої специфікації",
                                        "Вручную": "Вручну"},
    "А_ГруппыЗакупочнойСпецификации": {"Эталон": "Еталон", "ЭталонНеВыдавался": "Еталон — не видавалось",
                                       "ВнеЭталона": "Поза еталоном", "МШП": "МШП"},
    "А_ИсточникиЦеныЗакупки": {"ПоследняяЗакупка": "Остання закупівля",
                               "ПредыдущаяЗакупка": "Попередня закупівля (помилка одиниці)",
                               "ЦенаНормы": "Ціна норми (СС)", "СебестоимостьФакта": "Собівартість факту",
                               "Вручную": "Вручну"},
}
ENUMS = {name: list(values) for name, values in UK_ENUM.items()}
HEADER = ["ВидСпецификации", "СтруктураСебестоимости", "Основание", "Статус", "ДатаНачала", "КоличествоВыпуска",
          "ЕдиницаВыпуска", "НоменклатураВузла", "ДатаИзменения", "ДатаЦенНормы", "ДатаФакта", "Комментарий"]
TS = {
    "Состав": ["КлючСтроки", "ВидСтроки", "ЭтапСС", "ПозицияСС", "Номенклатура", "ИсточникКарточки", "Проверено",
               "НеПоставляется", "ОбщееНазвание", "Узел", "ПокупаетсяГотовым", "Количество", "ЕдиницаСтроки",
               "Коэффициент", "ИсточникКоэффициента", "ЦенаНормы", "Комментарий"],
    "Дома": ["Подразделение", "Комментарий"],
    "СоставОснования": ["КлючСтроки", "ВидСтроки", "ЭтапСС", "ПозицияСС", "Номенклатура", "ОбщееНазвание", "Узел",
                        "ПокупаетсяГотовым", "НеПоставляется", "Количество", "ЕдиницаСтроки", "Коэффициент", "ЦенаНормы"],
    "ДомаФакта": ["Подразделение", "ФактСумма", "Комментарий"],
    "ЗакупочныйСостав": ["КлючСтроки", "Номенклатура", "ГруппаЗакупки", "ОбщееНазвание", "Узел", "ПутьУзла", "ЭтапСС",
                         "ПозицияСС", "НормаЭталона", "СреднееФакта", "ДомовСФактом", "КоэффициентВариации",
                         "Количество", "ИзмененоВручную", "ПотребуетРешения", "Проверено", "ЦенаБезНДС", "ЦенаСНДС",
                         "ИсточникЦены", "ДокументЦены", "ЦенаНормы", "Комментарий"],
    "ФактПоДомам": ["Подразделение", "Номенклатура", "Списано", "СписаноСумма", "Лежит", "ЛежитСумма", "ИзСборки"],
}
RK_TS = {
    "НормаДома": ["КлючЛиста", "КлючСтроки", "КлючЭкземпляра", "Спецификация", "ВидСтроки", "Номенклатура",
                  "ОбщееНазвание", "ЭтапСС", "ПозицияСС", "ЕдиницаСтроки", "КоличествоЕдСтроки", "Количество",
                  "ЦенаНормы", "СуммаНормы", "ПутьУзла", "Уровень"],
    "ЭкземплярыУзлов": ["КлючЭкземпляра", "КлючЭкземпляраРодителя", "Узел", "НоменклатураВузла", "ПокупаетсяГотовым",
                        "ПутьУзла", "Уровень", "Количество", "СуммаНормы"],
}
ROLE_RIGHTS = {
    "ДобавлениеИзменениеДанныхБухгалтерии": ["Read", "Insert", "Update", "View", "InteractiveInsert", "Edit",
                                             "InteractiveSetDeletionMark", "InteractiveClearDeletionMark", "InputByString"],
    "ЧтениеДанныхБухгалтерии": ["Read", "View", "InputByString"],
    "ЗапускТонкогоКлиента": ["Read", "Insert", "Update", "View", "InteractiveInsert", "Edit",
                                             "InteractiveSetDeletionMark", "InteractiveClearDeletionMark", "InputByString"],
}


def names(coll):
    return [coll.Get(i).Name for i in range(coll.Count())]


def main():
    B = connect()
    md = B.Metadata
    fails = []

    def check(cond, text):
        if not cond:
            fails.append(text)

    for en, values in UK_ENUM.items():
        e = md.Enums.Find(en)
        if e is None:
            raise SystemExit(f"FAIL: перечисление {en} не найдено в базе")
        check(names(e.EnumValues) == list(values), f"{en}: значения {names(e.EnumValues)}")
        for i in range(e.EnumValues.Count()):
            v = e.EnumValues.Get(i)
            check(str(v.Synonym) == values.get(v.Name), f"{en}.{v.Name}: синоним uk «{v.Synonym}»")
    cat = md.Catalogs.Find("А_СпецификацииДомов")
    if cat is None:
        raise SystemExit("FAIL: справочник А_СпецификацииДомов не найден")
    check(str(cat.Synonym) == "Специфікації будинків і вузлів", f"синоним справочника «{cat.Synonym}»")
    check(bool(cat.Hierarchical), "справочник не иерархический")
    check(names(cat.Attributes) == HEADER, f"шапка {names(cat.Attributes)}")
    for i in range(cat.Attributes.Count()):
        check(str(cat.Attributes.Get(i).Synonym) != "", f"пустой uk-синоним {cat.Attributes.Get(i).Name}")
    check(names(cat.TabularSections) == list(TS), f"ТЧ {names(cat.TabularSections)}")
    for ts, cols in TS.items():
        t = cat.TabularSections.Find(ts)
        check(t is not None and names(t.Attributes) == cols, f"ТЧ {ts}: {names(t.Attributes) if t else None}")
        if t is not None:
            check(str(t.Synonym) != "", f"ТЧ {ts}: пустой uk-синоним")
    ss = md.Catalogs.Find("СтруктураСебестоимости")
    a = ss.Attributes.Find("ЭталоннаяСпецификация")
    check(a is not None and str(a.Synonym) == "Еталонна специфікація", "СС.ЭталоннаяСпецификация")
    rk = md.Documents.Find("РасчетКомплектаций")
    for n, syn in (("СпецификацияДома", "Специфікація будинку"), ("ДатаНормы", "Норма від"),
                   ("Спецификация", "Структура собівартості (СС)")):
        at = rk.Attributes.Find(n)
        check(at is not None and str(at.Synonym) == syn, f"РК.{n}: {at.Synonym if at else None}")
    for ts, cols in RK_TS.items():
        t = rk.TabularSections.Find(ts)
        check(t is not None and names(t.Attributes) == cols, f"РК.{ts}: {names(t.Attributes) if t else None}")
    ost = rk.TabularSections.Find("ТабличнаяЧастьОстатков")
    check(names(ost.Attributes)[-2:] == ["ПозицияСС", "НеИзНормы"], f"РК.ТабличнаяЧастьОстатков {names(ost.Attributes)[-2:]}")
    check(B.String(rk.Проведение) == B.String(B.Метаданные.СвойстваОбъектов.Проведение.Запретить), "РК стал проводимым")
    new_objs = [cat] + [md.Enums.Find(e) for e in UK_ENUM]
    for i in range(md.ExchangePlans.Count()):
        ep = md.ExchangePlans.Get(i)
        for o in new_objs:
            check(not ep.Content.Contains(o), f"{o.Name} в плане обмена {ep.Name}")
    for role, rights in ROLE_RIGHTS.items():
        r = md.Roles.Find(role)
        check(bool(B.ПравоДоступа("Чтение", cat, r)), f"{role}: нет Чтение")
        check(bool(B.ПравоДоступа("Изменение", cat, r)) == ("Update" in rights), f"{role}: Изменение")
    sub = md.Subsystems.Find("Склад").Subsystems.Find("А_СписанняБудинків")
    check(bool(sub.Content.Contains(cat)), "справочника нет в подсистеме А_СписанняБудинків")
    print(f"ПРОВАЛОВ: {len(fails)}")
    for f in fails:
        print("   ", f)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
