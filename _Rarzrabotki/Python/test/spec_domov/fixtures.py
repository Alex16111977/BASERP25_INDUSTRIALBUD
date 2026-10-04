# -*- coding: utf-8 -*-
"""Фикстуры приёмки П1. Тестовые элементы А_СпецификацииДомов создаются ТОЛЬКО внутри tx(B) и откатываются —
эталоны и закупочные в рабочую базу не пишутся. Имена начинаются с «__ТЕСТ_П1 ».
Коды единиц: 2938 компл, 2009 шт, 0101 м, 0123 м2, 5 рул (E3 §4.1)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import err, messages, q  # noqa: E402

PFX = "__ТЕСТ_П1 "


def enum(B, name, value):
    return getattr(getattr(B.Перечисления, name), value)


def unit(B, code):
    ref = B.Справочники.КлассификаторЕдиницИзмерения.НайтиПоКоду(code)
    assert not ref.Пустая(), f"единица {code}"
    return ref


def card(B, code):
    ref = B.Справочники.Номенклатура.НайтиПоКоду(code)
    assert not ref.Пустая(), f"карточка {code}"
    return ref


def ss(B, code):
    ref = B.Справочники.СтруктураСебестоимости.НайтиПоКоду(code)
    assert not ref.Пустая(), f"СС {code}"
    return ref


def ss_row(B, ss_code, n):
    rows = q(B, """ВЫБРАТЬ
	Компл.Этап КАК Этап, Компл.НоменклатураСС КАК НоменклатураСС, Компл.Количество КАК Количество,
	Компл.ЕдиницаСС КАК ЕдиницаСС, Компл.ОбщееНазвание КАК ОбщееНазвание, Компл.Номенклатура КАК Номенклатура,
	Компл.Сумма КАК Сумма, Компл.НомерСтроки КАК НомерСтроки
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код
	И Компл.НомерСтроки = &Н""", Код=ss_code, Н=n)
    assert len(rows) == 1, (ss_code, n)
    return rows[0]


def ss_first_row_with_card(B, ss_code):
    rows = q(B, """ВЫБРАТЬ ПЕРВЫЕ 1
	Компл.НомерСтроки КАК Н
ИЗ
	Справочник.СтруктураСебестоимости.Комплектующие КАК Компл
ГДЕ
	Компл.Ссылка.Код = &Код
	И Компл.Количество > 0
	И Компл.Номенклатура <> ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка)
УПОРЯДОЧИТЬ ПО
	Компл.НомерСтроки""", Код=ss_code)
    return ss_row(B, ss_code, int(rows[0]["Н"]))


def house(B, ss_code, n=1):
    rows = q(B, """ВЫБРАТЬ Подр.Ссылка КАК Ссылка ИЗ Справочник.Подразделения КАК Подр
ГДЕ Подр.Родитель = &Родитель И НЕ Подр.ПометкаУдаления УПОРЯДОЧИТЬ ПО Подр.Код""", Родитель=ss(B, ss_code).ПодразделениеОбъекта)
    return rows[n - 1]["Ссылка"]


def new_spec(B, kind, name, ss_ref=None, base_qty=1, unit_ref=None, st="Черновик"):
    obj = B.Справочники.А_СпецификацииДомов.СоздатьЭлемент()
    obj.Наименование = PFX + name
    obj.ВидСпецификации = enum(B, "А_ВидыСпецификацийДомов", kind)
    obj.Статус = enum(B, "А_СтатусыСпецификацийДомов", st)
    obj.КоличествоВыпуска = base_qty
    if ss_ref is not None:
        obj.СтруктураСебестоимости = ss_ref
    if unit_ref is not None:
        obj.ЕдиницаВыпуска = unit_ref
    return obj


def add_row(B, obj, kind="Деталь", card_ref=None, qty=1, unit_ref=None, node=None, stage=None, pos="", checked=False,
            ne_post=False, comment="", k=1, ksrc="ЕдиницыСовпадают", on=None, price=0, ready=False):
    r = obj.Состав.Добавить()
    r.ВидСтроки = enum(B, "А_ВидыСтрокСпецификацииДомов", kind)
    if card_ref is not None:
        r.Номенклатура = card_ref
        r.ОбщееНазвание = card_ref.ОбщееНазваниеНоменклатуры
    if on is not None:
        r.ОбщееНазвание = on
    if node is not None:
        r.Узел = node
    if stage is not None:
        r.ЭтапСС = stage
    r.ПозицияСС = pos
    r.Количество = qty
    if unit_ref is not None:
        r.ЕдиницаСтроки = unit_ref
    elif card_ref is not None:
        r.ЕдиницаСтроки = card_ref.БазоваяЕдиницаИзмерения
    r.Коэффициент = k
    r.ИсточникКоэффициента = enum(B, "А_ИсточникиКоэффициентаПересчетаСС", ksrc)
    r.Проверено = checked
    r.НеПоставляется = ne_post
    r.Комментарий = comment
    r.ЦенаНормы = price
    r.ПокупаетсяГотовым = ready
    return r


def write(obj, silent=False):
    if silent:
        obj.ОбменДанными.Загрузка = True
    obj.Записать()
    return obj.Ссылка


def expect_write_error(B, obj, *needles):
    """Запись должна упасть, в ошибке и сообщениях — все needles. После падения транзакция ошибочна: кейс завершать."""
    messages(B)
    try:
        obj.Записать()
    except Exception as e:
        text = err(e) + "\n" + "\n".join(messages(B))
        for n in needles:
            assert n in text, f"нет «{n}» в:\n{text}"
        return text
    raise AssertionError("запись прошла, а должна была упасть: " + ", ".join(needles))


def kn(B, number):
    rows = q(B, """ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка ИЗ Документ.КомплектацияНоменклатуры КАК Док
ГДЕ Док.Номер = &Номер И Док.Дата >= ДАТАВРЕМЯ(2026, 1, 1)""", Номер=number)
    assert len(rows) == 1, number
    return rows[0]["Ссылка"]


def srv(B):
    return B.А_СпецификацииДомовСервер


def row_index_by_ss(B, obj, ss_code, n):
    """Индекс (с 0) строки «Состав», у которой позиция = n-я строка СС (по ключу этап + текст)."""
    row = ss_row(B, ss_code, n)
    m = B.А_СтруктураСебестоимостиСервер
    key = m.КлючПозиции(row["Этап"], row["НоменклатураСС"])
    for i in range(obj.Состав.Количество()):
        r = obj.Состав.Получить(i)
        if r.ПозицияСС and m.КлючПозиции(r.ЭтапСС, r.ПозицияСС) == key:
            return i
    raise AssertionError(f"нет строки с позицией {ss_code}:{n}")


def col(vt, name):
    return [getattr(vt.Получить(i), name) for i in range(vt.Количество())]


def arr_list(a):
    return [a.Получить(i) for i in range(a.Количество())]
