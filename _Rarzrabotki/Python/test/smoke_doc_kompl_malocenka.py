# -*- coding: utf-8 -*-
"""Smoke 2b: РасчетКомплектаций — маршрутизация малоценки (счёт 22 -> ПередачаМалоценныхАктивовВЭксплуатацию).
Отдельный документ-расчёт по складу, где ЕСТЬ остатки сч.22. Get-or-create, без проведения/удаления."""
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER = "SMOKE_DOC_KOMPL_MAL_v1"

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')


def find_by_comment(doc_type, marker):
    q = buh.NewObject("Запрос")
    q.Text = (f"ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.{doc_type} КАК Д "
              "ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления")
    q.SetParameter("М", marker)
    r = q.Execute().Выгрузить()
    return r.Получить(0).Ссылка if r.Количество() > 0 else None


# склад с максимумом остатков сч.22
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1
	ВЫРАЗИТЬ(Ост.Субконто3 КАК Справочник.Склады) КАК Склад,
	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Ост.Субконто1) КАК Позиций
ИЗ
	РегистрБухгалтерии.Хозрасчетный.Остатки(&НаДату, Счет В ИЕРАРХИИ (&Счета), , ) КАК Ост
ГДЕ
	Ост.КоличествоОстаток > 0

СГРУППИРОВАТЬ ПО
	ВЫРАЗИТЬ(Ост.Субконто3 КАК Справочник.Склады)

УПОРЯДОЧИТЬ ПО
	Позиций УБЫВ
"""
q.SetParameter("НаДату", datetime.datetime(2026, 7, 7, 23, 59, 59))
arr = buh.NewObject("Массив")
arr.Добавить(buh.ПланыСчетов.Хозрасчетный.НайтиПоКоду("22"))
q.SetParameter("Счета", arr)
rows = q.Execute().Выгрузить()
assert rows.Количество() > 0, "нет складов с остатками сч.22"
skl = rows.Получить(0).Склад
print("Склад малоценки:", skl.Наименование, "| позиций:", rows.Получить(0).Позиций)

spec = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000004")

ref = find_by_comment("РасчетКомплектаций", MARKER)
if ref is not None:
    doc = ref.ПолучитьОбъект()
else:
    doc = buh.Документы.РасчетКомплектаций.СоздатьДокумент()
    doc.Дата = datetime.datetime.now().replace(microsecond=0)
    doc.Заполнить(None)
    doc.Комментарий = MARKER
doc.Спецификация = spec
doc.Период = datetime.datetime(2026, 7, 7)
doc.СкладыОстатков.Очистить()
doc.СкладыОстатков.Добавить().Склад = skl

doc.РассчитатьАнализ()
n = doc.ТабличнаяЧастьОстатков.Количество()
rows22 = 0
for i in range(n):
    r = doc.ТабличнаяЧастьОстатков.Получить(i)
    if str(r.Счет.Код).startswith("22"):
        rows22 += 1
print(f"Анализ: {n} строк, из них сч.22*: {rows22}")
assert rows22 > 0, "нет строк 22 — склад выбран неудачно"

# целевые малоценки + комплектации (get-or-create)
qorg = buh.NewObject("Запрос")
qorg.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 О.Ссылка КАК Ссылка ИЗ Справочник.Организации КАК О ГДЕ НЕ О.ПометкаУдаления"
org = qorg.Execute().Выгрузить().Получить(0).Ссылка


def get_or_create_target(doc_type, marker):
    r = find_by_comment(doc_type, marker)
    if r is not None:
        return r
    d = getattr(buh.Документы, doc_type).СоздатьДокумент()
    d.Дата = datetime.datetime.now().replace(microsecond=0)
    d.Организация = org
    d.Комментарий = marker
    d.Записать()
    return d.Ссылка


mal_n = get_or_create_target("ПередачаМалоценныхАктивовВЭксплуатацию", MARKER + "_Н")
mal_o = get_or_create_target("ПередачаМалоценныхАктивовВЭксплуатацию", MARKER + "_П")
kn_n = get_or_create_target("КомплектацияНоменклатуры", MARKER + "_КНН")
kn_o = get_or_create_target("КомплектацияНоменклатуры", MARKER + "_КНП")
for i in range(doc.ДокументиМалоценки.Количество()):
    row = doc.ДокументиМалоценки.Получить(i)
    row.ДокументПоНормам = mal_n
    row.ДокументДодаткова = mal_o
for i in range(doc.ДокументиКомплектації.Количество()):
    row = doc.ДокументиКомплектації.Получить(i)
    row.ДокументПоНормам = kn_n
    row.ДокументДодаткова = kn_o

res = doc.ЗаполнитьДокументыПоАнализу(True, True)
print(f"Результат: Заповнено={res.Заповнено}, Пропущено={res.Пропущено}, БезНазначення={res.БезНазначення}, Помилок={res.Помилки.Количество()}")
for i in range(res.Помилки.Количество()):
    print("  Помилка:", res.Помилки.Получить(i))
assert res.Помилки.Количество() == 0

total_mal = 0
for tref in (mal_n, mal_o):
    obj = tref.ПолучитьОбъект()
    for i in range(obj.МалоценныеАктивы.Количество()):
        r = obj.МалоценныеАктивы.Получить(i)
        assert str(r.СчетУчетаБУ.Код).startswith("22"), f"чужой счёт в малоценке: {r.СчетУчетаБУ.Код}"
        assert not r.НалоговоеНазначение.Пустая(), "пустое НалоговоеНазначение"
        assert float(r.Количество) > 0
    total_mal += obj.МалоценныеАктивы.Количество()
assert total_mal == rows22 or total_mal > 0, f"малоценка: {total_mal} строк при {rows22} строках 22"
# в комплектации не должно быть 22-х
total_kn22 = 0
for tref in (kn_n, kn_o):
    obj = tref.ПолучитьОбъект()
    for i in range(obj.Комплектующие.Количество()):
        if str(obj.Комплектующие.Получить(i).СчетУчетаБУ.Код).startswith("22"):
            total_kn22 += 1
assert total_kn22 == 0, f"строки сч.22 попали в комплектацию: {total_kn22}"
print(f"Малоценка: {total_mal} строк в МалоценныеАктивы (все 22*, налог.назначение есть); в комплектации 22-х нет")

doc.Записать()
print("MALOCENKA PASS")
